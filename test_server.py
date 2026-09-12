import pytest
import tempfile
import pandas as pd
from pathlib import Path
import openpyxl
from openpyxl import Workbook
from datetime import datetime, timedelta

import sys
SERVER_DIR = Path(__file__).resolve().parent
if str(SERVER_DIR) not in sys.path:
    sys.path.insert(0, str(SERVER_DIR))

from app import create_app
from fastapi.testclient import TestClient
from generators import (
    generate_ei_report,
    generate_forward_pendency_report,
    generate_reverse_pendency_report,
    generate_conversion_report,
    generate_nps_report,
    generate_tat_report,
    generate_vms_adherence_report,
    generate_second_attempt_adherence_report,
    generate_eob_report,
    generate_untraceable_report,
    generate_cpd_breach_report,
    generate_weekly_scm_tat_report
)

client = TestClient(create_app())
API_KEY = "OoV81VZ6ugIQ5qu_JNKfDM0jEp0SQyhpuZMaPTv5BbQ"
HEADERS = {"X-API-KEY": API_KEY}

def test_root():
    res = client.get("/")
    assert res.status_code == 200
    assert res.json()["server"] == "EI Stream Server"

def test_health():
    res = client.get("/health", headers=HEADERS)
    assert res.status_code == 200
    assert res.json()["status"] == "ok"

def test_health_invalid_key():
    res = client.get("/health", headers={"X-API-KEY": "wrong_key"})
    assert res.status_code == 403

def test_forward_pendency_stream():
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        src_xlsx = tmp_path / "fwd_src.xlsx"
        out_xlsx = tmp_path / "fwd_out.xlsx"
        
        wb = Workbook()
        ws = wb.active
        ws.title = "raw_data_North"
        headers = ["PendingShipments", "Source_DC", "Aging", "CustomerPriorityV2", "Attempt_Status"]
        ws.append(headers)
        ws.append(["SHIP1001", "ALG", 1, "P2", "Attempted"])
        ws.append(["SHIP1002", "AYP", 4, "P3", "Unattempted"])
        ws.append(["SHIP1003", "DEO", 7, "P4", "Attempted"])
        ws.append(["SHIP1004", "ALG", 0, "P0", "Attempted"])
        ws.append(["SHIP1005", "AYP", 2, "P1", "Unattempted"])
        ws.append(["SHIP1006", "CAR-KHR", 2, "P0", "Attempted"])
        wb.save(src_xlsx)
        
        generate_forward_pendency_report(src_xlsx, out_xlsx)
        assert out_xlsx.exists()

        res_wb = openpyxl.load_workbook(out_xlsx)
        assert "Summary" in res_wb.sheetnames
        assert "CPD-DID pendency" in res_wb.sheetnames
        assert "RAW" in res_wb.sheetnames

        # Check Summary Priority Table headers (row 3, columns 9 to 15)
        sum_ws = res_wb["Summary"]
        prio_headers = [sum_ws.cell(row=3, column=c).value for c in range(9, 16)]
        assert prio_headers == ["Source DC", "P0", "P1", "P2", "P3", "P4", "Total Pendency"]

        # Check CAR is in summary and CAR-KHR is not
        dcs_in_summary = [sum_ws.cell(row=r, column=9).value for r in range(4, sum_ws.max_row + 1)]
        assert "CAR" in dcs_in_summary
        assert "CAR-KHR" not in dcs_in_summary

        # Check RAW sheet contains CAR, not CAR-KHR
        raw_ws = res_wb["RAW"]
        raw_dcs = [raw_ws.cell(row=r, column=2).value for r in range(2, raw_ws.max_row + 1)]
        assert "CAR" in raw_dcs
        assert "CAR-KHR" not in raw_dcs

        # Check CPD-DID sheet contains P0, P1, P2, P3 rows and CAR
        cpd_ws = res_wb["CPD-DID pendency"]
        cpd_rows = list(cpd_ws.iter_rows(values_only=True))
        assert len(cpd_rows) == 6  # header + 5 data rows (P2, P3, P0, P1, P0 from CAR-KHR)
        priorities_in_cpd = [r[4] for r in cpd_rows[1:]]
        assert set(priorities_in_cpd) == {"P0", "P1", "P2", "P3"}
        cpd_dcs = [r[1] for r in cpd_rows[1:]]
        assert "CAR" in cpd_dcs
        assert "CAR-KHR" not in cpd_dcs
        res_wb.close()

def test_reverse_pendency_stream():
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        src_xlsx = tmp_path / "rev_src.xlsx"
        out_xlsx = tmp_path / "rev_out.xlsx"
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Raw"
        headers = ["tracking_number", "Source DC", "Region", "Aging", "Age_Bucket", "Attempt_Status"]
        ws.append(headers)
        ws.append(["TRACK101", "ALG", "North", 1, "", "Done"])
        ws.append(["TRACK102", "AYP", "North", 3, "", "Pending"])
        ws.append(["TRACK103", "CAR-KHR", "North", 4, "", "Pending"])
        wb.save(src_xlsx)
        
        generate_reverse_pendency_report(src_xlsx, out_xlsx)
        assert out_xlsx.exists()

        res_wb = openpyxl.load_workbook(out_xlsx)
        assert "Summary" in res_wb.sheetnames
        assert "Critical P0" in res_wb.sheetnames
        assert "Raw" in res_wb.sheetnames

        # Check Raw sheet contains CAR, not CAR-KHR
        raw_ws = res_wb["Raw"]
        raw_dcs = [raw_ws.cell(row=r, column=2).value for r in range(2, raw_ws.max_row + 1)]
        assert "CAR" in raw_dcs
        assert "CAR-KHR" not in raw_dcs

        # Check Critical P0 contains CAR
        p0_ws = res_wb["Critical P0"]
        p0_dcs = [p0_ws.cell(row=r, column=2).value for r in range(2, p0_ws.max_row + 1)]
        assert "CAR" in p0_dcs
        assert "CAR-KHR" not in p0_dcs
        res_wb.close()

def test_conversion_stream():
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        src_xlsx = tmp_path / "conv_src.xlsx"
        out_sameday = tmp_path / "conv_sameday_out.xlsx"
        out_d1 = tmp_path / "conv_d1_out.xlsx"
        
        with pd.ExcelWriter(src_xlsx, engine='openpyxl') as writer:
            dc_df = pd.DataFrame([
                {
                    'Source_DC': 'ALG', 'Picked-up': 80.0, 'OFP': 100.0, 'Succ_pickup%': '80%', 'Succ_del%': '85%',
                    'COD_Succ_del%': '80%', 'PP_Succ_del%': '90%'
                },
                {
                    'Source_DC': 'CAR', 'Picked-up': 30.0, 'OFP': 50.0, 'Succ_pickup%': '60%', 'Succ_del%': '70%',
                    'COD_Succ_del%': '70%', 'PP_Succ_del%': '70%'
                },
                {
                    'Source_DC': 'CAR-KHR', 'Picked-up': 20.0, 'OFP': 50.0, 'Succ_pickup%': '40%', 'Succ_del%': '80%',
                    'COD_Succ_del%': '80%', 'PP_Succ_del%': '80%'
                }
            ])
            agent_df = pd.DataFrame([
                {
                    'Source_DC': 'ALG', 'Picked-up': 80.0, 'OFP': 100.0, 'del_update': 85.0, 'OFD': 100.0
                },
                {
                    'Source_DC': 'CAR-KHR', 'Picked-up': 20.0, 'OFP': 50.0, 'del_update': 40.0, 'OFD': 50.0
                }
            ])
            e2e_cols = [f"col_{i}" for i in range(25)]
            e2e_cols[22] = "Source_DC"
            row = ["val"] * 25
            row[22] = "ALG"
            e2e_df = pd.DataFrame([row], columns=e2e_cols)
            
            dc_df.to_excel(writer, sheet_name='E2E_DC', index=False)
            agent_df.to_excel(writer, sheet_name='Agent_view', index=False)
            e2e_df.to_excel(writer, sheet_name='E2E_Raw', index=False)
            
        generate_conversion_report(src_xlsx, out_sameday, sub_type='sameday')
        assert out_sameday.exists()
        from openpyxl import load_workbook
        wb_same = load_workbook(out_sameday)
        assert wb_same.sheetnames == ['Sameday DC_View', 'Sameday Agent_View']
        ws_dc = wb_same['Sameday DC_View']
        # Check whole number formatting
        for col_idx in range(1, ws_dc.max_column + 1):
            header = ws_dc.cell(row=1, column=col_idx).value
            cell_val = ws_dc.cell(row=2, column=col_idx).value
            num_fmt = ws_dc.cell(row=2, column=col_idx).number_format
            if header in ('Picked-up', 'OFP'):
                assert isinstance(cell_val, int)
                assert num_fmt == '0'
            elif header and '%' in str(header):
                assert num_fmt == '0.0%'
        wb_same.close()

        generate_conversion_report(src_xlsx, out_d1, sub_type='d-1')
        assert out_d1.exists()
        wb_d1 = load_workbook(out_d1)
        assert wb_d1.sheetnames == ['D-1 DC_View', 'D-1 Agent_View']
        ws_d1_dc = wb_d1['D-1 DC_View']
        # CAR and CAR-KHR should be consolidated into a single 'CAR' row
        dc_names = [ws_d1_dc.cell(row=r, column=1).value for r in range(2, ws_d1_dc.max_row + 1)]
        assert "CAR" in dc_names
        assert "CAR-KHR" not in dc_names

        # Agent view should have CAR (from CAR-KHR)
        ws_d1_agent = wb_d1['D-1 Agent_View']
        agent_dcs = [ws_d1_agent.cell(row=r, column=1).value for r in range(2, ws_d1_agent.max_row + 1)]
        assert "CAR" in agent_dcs
        assert "CAR-KHR" not in agent_dcs
        wb_d1.close()

def test_nps_stream():
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        src_xlsx = tmp_path / "nps_src.xlsx"
        out_xlsx = tmp_path / "nps_out.xlsx"
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        headers = ["Col0", "Col1", "Col2", "Col3", "Option", "Col5", "Col6", "Col7", "Col8", "Col9",
                   "Col10", "Col11", "Col12", "Col13", "Col14", "Col15", "Col16", "Col17", "Col18", "Col19",
                   "Col20", "Col21", "Agent_Name", "Col23", "Col24", "Col25", "Col26", "Col27", "Source_DC"]
        ws.append(headers)
        row1 = [""] * 29
        row1[4] = "Promoter"
        row1[22] = "Agent A"
        row1[28] = "ALG"
        ws.append(row1)

        row2 = [""] * 29
        row2[4] = "Neutral"
        row2[22] = "Agent B"
        row2[28] = "CAR-KHR"
        ws.append(row2)
        wb.save(src_xlsx)
        
        generate_nps_report(src_xlsx, out_xlsx)
        assert out_xlsx.exists()

        wb_out = openpyxl.load_workbook(out_xlsx)
        assert "Raw" in wb_out.sheetnames
        raw_ws = wb_out["Raw"]
        raw_dcs = [raw_ws.cell(row=r, column=29).value for r in range(2, raw_ws.max_row + 1)]
        assert "CAR" in raw_dcs
        assert "CAR-KHR" not in raw_dcs
        wb_out.close()

def test_tat_stream():
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        src_xlsx = tmp_path / "tat_src.xlsx"
        out_xlsx = tmp_path / "tat_out.xlsx"
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        headers = ["Col0", "Col1", "Col2", "Col3", "Status", "Col5", "Col6", "Col7", "Col8", "Col9",
                   "Col10", "Col11", "Col12", "Col13", "Col14", "Col15", "Col16", "Col17", "Col18", "Col19",
                   "Col20", "Col21", "Col22", "Col23", "Col24", "Col25", "Col26", "Col27", "Source_DC"]
        ws.append(headers)
        row1 = [""] * 29
        row1[4] = "Complete"
        row1[28] = "ALG"
        ws.append(row1)

        row2 = [""] * 29
        row2[4] = "Pending"
        row2[28] = "CAR-KHR"
        ws.append(row2)
        wb.save(src_xlsx)
        
        generate_tat_report(src_xlsx, out_xlsx)
        assert out_xlsx.exists()

        wb_out = openpyxl.load_workbook(out_xlsx)
        assert "SCM TAT raw data" in wb_out.sheetnames
        raw_ws = wb_out["SCM TAT raw data"]
        raw_dcs = [raw_ws.cell(row=r, column=29).value for r in range(2, raw_ws.max_row + 1)]
        assert "CAR" in raw_dcs
        assert "CAR-KHR" not in raw_dcs
        wb_out.close()

def test_vms_stream():
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        src_xlsx = tmp_path / "vms_src.xlsx"
        out_xlsx = tmp_path / "vms_out.xlsx"
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Raw"
        headers = ["Source_DC", "VMS_Status", "Other"]
        ws.append(headers)
        ws.append(["ALG", "Done", "Val"])
        ws.append(["AYP", "Not Done", "Val"])
        ws.append(["KNP", "Adherence", "Val"])
        ws.append(["LKO", "Non-Adherence", "Val"])
        ws.append(["GZB", "Non Adherence", "Val"])
        ws.append(["CAR-KHR", "Done", "Val"])
        wb.save(src_xlsx)
        
        generate_vms_adherence_report(src_xlsx, out_xlsx)
        assert out_xlsx.exists()

        wb_out = openpyxl.load_workbook(out_xlsx, data_only=True)
        assert "Summary" in wb_out.sheetnames
        assert "Raw" in wb_out.sheetnames
        ws_sum = wb_out["Summary"]

        # Check headers
        headers_read = [ws_sum.cell(row=3, column=c).value for c in range(1, 6)]
        assert headers_read == ['Source DC', 'Total', 'Adherence', 'Non-Adherence', 'Adherence %']

        # Check rows content
        rows_data = {}
        for r in range(4, 10):
            dc = ws_sum.cell(row=r, column=1).value
            if dc and dc != 'TOTAL / SUMMARY':
                rows_data[dc] = {
                    "total": ws_sum.cell(row=r, column=2).value,
                    "adh": ws_sum.cell(row=r, column=3).value,
                    "non_adh": ws_sum.cell(row=r, column=4).value,
                    "pct": ws_sum.cell(row=r, column=5).value,
                }
        assert rows_data["KNP"]["adh"] == 1 and rows_data["KNP"]["non_adh"] == 0
        assert rows_data["LKO"]["adh"] == 0 and rows_data["LKO"]["non_adh"] == 1
        assert rows_data["GZB"]["adh"] == 0 and rows_data["GZB"]["non_adh"] == 1
        assert rows_data["ALG"]["adh"] == 1 and rows_data["ALG"]["non_adh"] == 0
        assert rows_data["AYP"]["adh"] == 0 and rows_data["AYP"]["non_adh"] == 1
        assert rows_data["CAR"]["adh"] == 1 and rows_data["CAR"]["non_adh"] == 0
        assert "CAR-KHR" not in rows_data

        raw_ws = wb_out["Raw"]
        raw_dcs = [raw_ws.cell(row=r, column=1).value for r in range(2, raw_ws.max_row + 1)]
        assert "CAR" in raw_dcs
        assert "CAR-KHR" not in raw_dcs
        wb_out.close()

def test_second_attempt_stream():
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        src_xlsx = tmp_path / "second_attempt_src.xlsx"
        out_xlsx = tmp_path / "second_attempt_out.xlsx"
        
        wb = Workbook()
        
        # 1. Summary sheet
        ws_sum = wb.active
        ws_sum.title = "Summary"
        ws_sum.append(["FWD Metrics", "", "", "", "", "", "REV Metrics", "", "", "", ""])
        ws_sum.append(["Source DC", "Non Adherence", "Adherence", "Grand Total", "Adherence %", "", "Source DC", "Non Adherence", "Adherence", "Grand Total", "Adherence %"])
        ws_sum.append(["ALG", 5, 45, 50, 0.90, "", "ALG", 2, 18, 20, 0.90])
        ws_sum.append(["AYP", 10, 15, 25, 0.60, "", "AYP", 1, 9, 10, 0.90])
        ws_sum.append(["CAR-KHR", 2, 8, 10, 0.80, "", "CAR-KHR", 1, 4, 5, 0.80])
        
        # 2. FWD sheet
        ws_fwd = wb.create_sheet("FWD")
        ws_fwd.append(["Tracking_No", "Source_DC", "Agent", "Status"])
        ws_fwd.append(["TRK001", "ALG", "Agent 1", "Delivered"])
        ws_fwd.append(["TRK002", "AYP", "Agent 2", "Undelivered"])
        ws_fwd.append(["TRK004", "CAR-KHR", "Agent 4", "Delivered"])
        
        # 3. REV sheet
        ws_rev = wb.create_sheet("REV")
        ws_rev.append(["Tracking_No", "Source_DC", "Agent", "Status"])
        ws_rev.append(["TRK003", "ALG", "Agent 3", "Returned"])
        
        wb.save(src_xlsx)
        
        generate_second_attempt_adherence_report(src_xlsx, out_xlsx)
        assert out_xlsx.exists()

        wb_out = openpyxl.load_workbook(out_xlsx)
        assert "Raw" in wb_out.sheetnames
        raw_ws = wb_out["Raw"]
        # Source_DC is at column 3 (since col 1 is Flow_Type)
        raw_dcs = [raw_ws.cell(row=r, column=3).value for r in range(2, raw_ws.max_row + 1)]
        assert "CAR" in raw_dcs
        assert "CAR-KHR" not in raw_dcs
        wb_out.close()

def test_eob_stream():
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        src_xlsx = tmp_path / "eob_src.xlsx"
        out_xlsx = tmp_path / "eob_out.xlsx"
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Raw"
        headers = ["Tracking No", "Source DC", "Latest Status", "Ageing Bucket"]
        ws.append(headers)
        ws.append(["TRACK1", "ALG", "Out_For_Delivery", "1-2 days"])
        ws.append(["TRACK2", "AYP", "Undelivered_Attempted", "3-5 days"])
        ws.append(["TRACK3", "CAR-KHR", "Out_For_Delivery", "1-2 days"])
        wb.save(src_xlsx)
        
        generate_eob_report(src_xlsx, out_xlsx)
        assert out_xlsx.exists()

        wb_out = openpyxl.load_workbook(out_xlsx)
        assert "Raw" in wb_out.sheetnames
        raw_ws = wb_out["Raw"]
        raw_dcs = [raw_ws.cell(row=r, column=2).value for r in range(2, raw_ws.max_row + 1)]
        assert "CAR" in raw_dcs
        assert "CAR-KHR" not in raw_dcs
        wb_out.close()

def test_untraceable_stream():
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        src_xlsx = tmp_path / "ut_src.xlsx"
        out_xlsx = tmp_path / "ut_out.xlsx"
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Raw"
        headers = ["ShipmentId", "Source DC", "Age Bucket", "Amount"]
        ws.append(headers)
        ws.append(["SHIP1", "ALG", "0-2 Days", 500])
        ws.append(["SHIP2", "AYP", "6-10 Days", 1200])
        ws.append(["SHIP3", "CAR-KHR", "0-2 Days", 800])
        wb.save(src_xlsx)
        
        generate_untraceable_report(src_xlsx, out_xlsx)
        assert out_xlsx.exists()

        wb_out = openpyxl.load_workbook(out_xlsx)
        assert "Raw" in wb_out.sheetnames
        raw_ws = wb_out["Raw"]
        raw_dcs = [raw_ws.cell(row=r, column=2).value for r in range(2, raw_ws.max_row + 1)]
        assert "CAR" in raw_dcs
        assert "CAR-KHR" not in raw_dcs
        wb_out.close()

def test_ei_stream():
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        src_xlsx = tmp_path / "ei_src.xlsx"
        out_xlsx = tmp_path / "ei_out.xlsx"

        wb = Workbook()
        ws_task = wb.active
        ws_task.title = "Task_per_1k"

        yesterday = datetime.now() - timedelta(days=1)
        ws_task.cell(1, 1, "DC")
        ws_task.cell(1, 2, "Region")
        ws_task.cell(1, 3, "City")
        ws_task.cell(1, 4, yesterday)
        ws_task.cell(1, 10, "WTD")

        ws_task.cell(3, 1, "ALG")
        ws_task.cell(3, 2, "North")
        ws_task.cell(3, 3, "Aligarh")
        # daily
        ws_task.cell(3, 4, 100) # OFD
        ws_task.cell(3, 5, 2)   # FWD Task
        ws_task.cell(3, 6, 20)  # FWD 1k
        ws_task.cell(3, 7, 50)  # OFP
        ws_task.cell(3, 8, 1)   # REV Task
        ws_task.cell(3, 9, 20)  # REV 1k
        # WTD
        ws_task.cell(3, 10, 500) # OFD
        ws_task.cell(3, 11, 10)  # FWD Task
        ws_task.cell(3, 12, 20)  # FWD 1k
        ws_task.cell(3, 13, 250) # OFP
        ws_task.cell(3, 14, 5)   # REV Task
        ws_task.cell(3, 15, 20)  # REV 1k

        ws_task.cell(4, 1, "CAR-KHR")
        ws_task.cell(4, 2, "North")
        ws_task.cell(4, 3, "Khurja")
        # daily
        ws_task.cell(4, 4, 50)
        ws_task.cell(4, 5, 1)
        ws_task.cell(4, 6, 20)
        ws_task.cell(4, 7, 25)
        ws_task.cell(4, 8, 1)
        ws_task.cell(4, 9, 40)
        # WTD
        ws_task.cell(4, 10, 250)
        ws_task.cell(4, 11, 5)
        ws_task.cell(4, 12, 20)
        ws_task.cell(4, 13, 125)
        ws_task.cell(4, 14, 2)
        ws_task.cell(4, 15, 16)

        ws_raw = wb.create_sheet("Raw")
        ws_raw.append(["Source_DC", "Final_tracking_no", "fwd_agent name", "rev_agent name"])
        ws_raw.append(["ALG", "MYSC12345", "Agent A", ""])
        ws_raw.append(["ALG", "MYSR12345", "", "Agent B"])
        ws_raw.append(["CAR-KHR", "MYSC99999", "Agent C", ""])

        wb.save(src_xlsx)

        generate_ei_report(str(src_xlsx), str(out_xlsx))
        assert out_xlsx.exists()

        import openpyxl
        wb_out = openpyxl.load_workbook(out_xlsx)
        assert wb_out.sheetnames == ["SUMMARY", "Filtered_Source_DC", "FWD EI", "REVERSE EI", "Agent Summary"]
        ws_summary = wb_out["SUMMARY"]

        # Check side-by-side title headers in row 1
        assert ws_summary.cell(1, 1).value == "Forward EI"
        assert ws_summary.cell(1, 7).value == "Reverse EI"
        assert ws_summary.cell(1, 13).value == "Weekly Forward EI"
        assert ws_summary.cell(1, 19).value == "Weekly Reverse EI"

        # Check column headers in row 2
        assert ws_summary.cell(2, 1).value == "Date"
        assert ws_summary.cell(2, 2).value == "Source_DC"
        assert ws_summary.cell(2, 7).value == "Date"
        assert ws_summary.cell(2, 8).value == "Source_DC"
        assert ws_summary.cell(2, 13).value == "Date"
        assert ws_summary.cell(2, 14).value == "Source_DC"
        assert ws_summary.cell(2, 19).value == "Date"
        assert ws_summary.cell(2, 20).value == "Source_DC"

        # Check Filtered_Source_DC sheet contains CAR, not CAR-KHR
        ws_filt = wb_out["Filtered_Source_DC"]
        filt_dcs = [ws_filt.cell(row=r, column=1).value for r in range(2, ws_filt.max_row + 1)]
        assert "CAR" in filt_dcs
        assert "CAR-KHR" not in filt_dcs


def test_cpd_breach_stream():
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        src_xlsx = tmp_path / "breach_src.xlsx"
        out_xlsx = tmp_path / "breach_out.xlsx"

        wb = Workbook()
        ws = wb.active
        ws.title = "LM"
        headers = [
            'tracking_number', 'order_created_date', 'customer_promise_date',
            'New_Final_delay_tag', 'rto_ic_flag', 'Latest Status', 'Current Location',
            'Latest Update Time', 'Cs Notes', 'No. of Attempts', 'DC Code', 'Source DC', 'Region'
        ]
        ws.append(headers)
        ws.append(["TRK001", "01/09/26 10:00", "09/09/2026", "Customer Attributed", 0, "Delivered", "Hub1", "46275.0", "", 1, "ALG", "ALG", "North"])
        ws.append(["TRK002", "02/09/26 11:00", "09/09/2026", "Last Mile delay", 0, "Out_For_Delivery", "Hub2", "46275.1", "", 2, "ALG", "ALG", "North"])
        ws.append(["TRK003", "03/09/26 12:00", "09/09/2026", "RTO/IC - NCD", 1, "RTO", "Hub3", "46275.2", "", 0, "ALL", "ALL", "North"])
        ws.append(["TRK004", "04/09/26 13:00", "09/09/2026", "Last Mile delay", 0, "Out_For_Delivery", "Hub4", "46275.3", "", 1, "XYZ_NON_ALLOWED", "XYZ_NON_ALLOWED", "South"])
        ws.append(["TRK005", "05/09/26 14:00", "09/09/2026", "Customer Attributed", 0, "Delivered", "Hub5", "46275.4", "", 1, "CAR-KHR", "CAR-KHR", "North"])
        wb.save(src_xlsx)

        generate_cpd_breach_report(src_xlsx, out_xlsx)
        assert out_xlsx.exists()

        wb_out = openpyxl.load_workbook(out_xlsx)
        assert wb_out.sheetnames == ["summary", "raw"]

        ws_sum = wb_out["summary"]
        assert ws_sum.cell(1, 1).value == "Source DC"
        assert ws_sum.cell(1, 2).value == "Customer Attributed"
        assert ws_sum.cell(1, 3).value == "Last Mile delay"
        assert ws_sum.cell(1, 4).value == "RTO/IC - NCD"
        assert ws_sum.cell(1, 5).value == "Total CPD Breach"

        # Check Total Result row at bottom
        last_row = ws_sum.max_row
        assert ws_sum.cell(last_row, 1).value == "Total Result"
        assert ws_sum.cell(last_row, 2).value == 2  # Customer Attributed (ALG + CAR-KHR)
        assert ws_sum.cell(last_row, 3).value == 1  # Last Mile delay
        assert ws_sum.cell(last_row, 4).value == 1  # RTO/IC - NCD
        assert ws_sum.cell(last_row, 5).value == 4  # Total CPD Breach

        # Check summary has CAR and not CAR-KHR
        sum_dcs = [ws_sum.cell(row=r, column=1).value for r in range(2, last_row)]
        assert "CAR" in sum_dcs
        assert "CAR-KHR" not in sum_dcs

        # Check raw sheet: should have 4 rows (excluding non-allowed DC)
        ws_raw = wb_out["raw"]
        assert ws_raw.max_row == 5  # header + 4 data rows
        raw_dcs = [ws_raw.cell(row=r, column=12).value for r in range(2, ws_raw.max_row + 1)]
        assert "CAR" in raw_dcs
        assert "CAR-KHR" not in raw_dcs
        wb_out.close()


def test_weekly_scm_tat_stream():
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        src_xlsx = tmp_path / "weekly_tat_src.xlsx"
        out_xlsx = tmp_path / "weekly_tat_out.xlsx"

        wb = Workbook()
        
        # 1. DC sheet
        ws_dc = wb.active
        ws_dc.title = "DC"
        ws_dc.append(["Overall DC"] + [""] * 8 + ["Forward DC"] + [""] * 8 + ["Reverse DC"])
        ws_dc.append([None, "06-Sep-26", "07-Sep-26", "08-Sep-26", "09-Sep-26", "10-Sep-26", "11-Sep-26", "Grand Total", None, "Forward DC", "06-Sep-26", "07-Sep-26", "08-Sep-26", "09-Sep-26", "10-Sep-26", "11-Sep-26", "Grand Total", None, "Reverse DC", "06-Sep-26", "07-Sep-26", "08-Sep-26", "09-Sep-26", "10-Sep-26", "11-Sep-26", "Grand Total"])
        ws_dc.append(["ALG", 0.96, 0.92, 0.85, 0.98, 0.95, 0.91, 0.928, None, "ALG", 0.96, 0.92, 0.85, 0.98, 0.95, 0.91, 0.928, None, "ALG", 0.96, 0.92, 0.85, 0.98, 0.95, 0.91, 0.928])
        ws_dc.append(["AYP", 0.99, 0.99, 0.97, 0.98, 0.99, 0.96, 0.98, None, "AYP", 0.99, 0.99, 0.97, 0.98, 0.99, 0.96, 0.98, None, "AYP", 0.99, 0.99, 0.97, 0.98, 0.99, 0.96, 0.98])
        
        # 2. Data sheet
        ws_data = wb.create_sheet(title="Data")
        headers = ["dummy"] * 37
        headers[4] = "status_status"
        headers[11] = "l4_name"
        headers[12] = "l5_name"
        headers[23] = "Attribute"
        headers[26] = "Source DC"
        headers[34] = "Aging"
        ws_data.append(headers)

        row1 = ["val"] * 37
        row1[4] = "Open"
        row1[11] = "Delay"
        row1[12] = "Traffic"
        row1[23] = "Forward"
        row1[26] = "ALG"
        row1[34] = "1"
        ws_data.append(row1)

        row2 = ["val"] * 37
        row2[4] = "Closed"
        row2[11] = "Damage"
        row2[12] = "Handling"
        row2[23] = "Reverse"
        row2[26] = "ALG"
        row2[34] = "2"
        ws_data.append(row2)

        row3 = ["val"] * 37
        row3[4] = "Pending"
        row3[11] = "Delay"
        row3[12] = "Weather"
        row3[23] = "Reverse"
        row3[26] = "AYP"
        row3[34] = "3"
        ws_data.append(row3)

        row4 = ["val"] * 37
        row4[4] = "Open"
        row4[11] = "Other"
        row4[12] = "Other"
        row4[23] = "Forward"
        row4[26] = "NON_ALLOWED_DC"
        row4[34] = "1"
        ws_data.append(row4)

        row5 = ["val"] * 37
        row5[4] = "Open"
        row5[11] = "Delay"
        row5[12] = "Weather"
        row5[23] = "Forward"
        row5[26] = "CAR-KHR"
        row5[34] = "1"
        ws_data.append(row5)

        wb.save(src_xlsx)

        generate_weekly_scm_tat_report(src_xlsx, out_xlsx)
        assert out_xlsx.exists()

        wb_out = openpyxl.load_workbook(out_xlsx)
        assert wb_out.sheetnames == ["summary", "raw data", "todays tasks", "today's task summary"]

        # Check KPI cards on today's task summary
        ws_kpi = wb_out["today's task summary"]
        assert ws_kpi.cell(2, 1).value == "OPEN TASKS"
        assert str(ws_kpi.cell(3, 1).value).replace(",", "") == "3"
        assert ws_kpi.cell(2, 4).value == "FORWARD FLOW"
        assert str(ws_kpi.cell(3, 4).value).replace(",", "") == "2"
        assert ws_kpi.cell(2, 7).value == "REVERSE FLOW"
        assert str(ws_kpi.cell(3, 7).value).replace(",", "") == "1"

        # Check raw data has 4 rows (excluding non-allowed DC)
        ws_raw = wb_out["raw data"]
        assert ws_raw.max_row == 5  # header + 4 data rows
        raw_dcs = [ws_raw.cell(row=r, column=27).value for r in range(2, ws_raw.max_row + 1)]
        assert "CAR" in raw_dcs
        assert "CAR-KHR" not in raw_dcs

        # Check todays tasks has 3 rows (only non-closed allowed rows)
        ws_tasks = wb_out["todays tasks"]
        assert ws_tasks.max_row == 4  # header + 3 open tasks
        task_dcs = [ws_tasks.cell(row=r, column=27).value for r in range(2, ws_tasks.max_row + 1)]
        assert "CAR" in task_dcs
        assert "CAR-KHR" not in task_dcs

        wb_out.close()
