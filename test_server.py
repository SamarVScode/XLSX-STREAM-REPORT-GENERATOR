import pytest
import tempfile
import pandas as pd
from pathlib import Path
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
    generate_untraceable_report
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
        wb.save(src_xlsx)
        
        generate_forward_pendency_report(src_xlsx, out_xlsx)
        assert out_xlsx.exists()

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
        wb.save(src_xlsx)
        
        generate_reverse_pendency_report(src_xlsx, out_xlsx)
        assert out_xlsx.exists()

def test_conversion_stream():
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        src_xlsx = tmp_path / "conv_src.xlsx"
        out_xlsx = tmp_path / "conv_out.xlsx"
        
        with pd.ExcelWriter(src_xlsx, engine='openpyxl') as writer:
            dc_df = pd.DataFrame([{
                'Source_DC': 'ALG', 'Succ_pickup%': '80%', 'Succ_del%': '85%',
                'COD_Succ_del%': '80%', 'PP_Succ_del%': '90%'
            }])
            agent_df = pd.DataFrame([{
                'Source_DC': 'ALG', 'Picked-up': 80, 'OFP': 100, 'del_update': 85, 'OFD': 100
            }])
            e2e_cols = [f"col_{i}" for i in range(25)]
            e2e_cols[22] = "Source_DC"
            row = ["val"] * 25
            row[22] = "ALG"
            e2e_df = pd.DataFrame([row], columns=e2e_cols)
            
            dc_df.to_excel(writer, sheet_name='E2E_DC', index=False)
            agent_df.to_excel(writer, sheet_name='Agent_view', index=False)
            e2e_df.to_excel(writer, sheet_name='E2E_Raw', index=False)
            
        generate_conversion_report(src_xlsx, out_xlsx)
        assert out_xlsx.exists()

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
        wb.save(src_xlsx)
        
        generate_nps_report(src_xlsx, out_xlsx)
        assert out_xlsx.exists()

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
        wb.save(src_xlsx)
        
        generate_tat_report(src_xlsx, out_xlsx)
        assert out_xlsx.exists()

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
        wb.save(src_xlsx)
        
        generate_vms_adherence_report(src_xlsx, out_xlsx)
        assert out_xlsx.exists()

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
        wb.save(src_xlsx)
        
        generate_eob_report(src_xlsx, out_xlsx)
        assert out_xlsx.exists()

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
        wb.save(src_xlsx)
        
        generate_untraceable_report(src_xlsx, out_xlsx)
        assert out_xlsx.exists()

def test_ei_stream():
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        src_xlsx = tmp_path / "ei_src.xlsx"
        out_xlsx = tmp_path / "ei_out.xlsx"
        
        wb = Workbook()
        ws_task = wb.active
        ws_task.title = "Task_per_1k"
        
        today = datetime.now().date()
        yday = datetime.combine(today - timedelta(days=1), datetime.min.time())
        
        row1 = ["Source_DC", "Region", "City", yday, "", "", "", "", "", "WTD", "", "", "", "", ""]
        ws_task.append(row1)
        row2 = ["", "", "", "OFD", "Forward_Task", "Fwd_Task_per_1k", "OFP", "Reverse_Task", "Rev_Task_per_1k",
                "OFD", "Forward_Task", "Fwd_Task_per_1k", "OFP", "Reverse_Task", "Rev_Task_per_1k"]
        ws_task.append(row2)
        row3 = ["ALG", "North", "Aligarh", 100, 5, 50.0, 50, 2, 40.0, 500, 20, 40.0, 200, 10, 50.0]
        ws_task.append(row3)
        
        ws_raw = wb.create_sheet("Raw")
        raw_headers = ["source_dc", "tracking_no", "agent_name"]
        ws_raw.append(raw_headers)
        ws_raw.append(["ALG", "MYSC12345", "Agent 1"])
        ws_raw.append(["ALG", "MYSR12345", "Agent 2"])
        
        wb.save(src_xlsx)
        
        generate_ei_report(str(src_xlsx), str(out_xlsx))
        assert out_xlsx.exists()
