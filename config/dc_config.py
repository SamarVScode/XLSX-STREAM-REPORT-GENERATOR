# ============================================================
# DC Configuration — ei_stream_server
# ============================================================
# Allowed Source DCs across ALL server report generators
# ============================================================

ALLOWED_SOURCE_DCS = [
    'ALG', 'AYP', 'DEO', 'JHS', 'JNP',
    'KNP', 'MAU', 'MRZ', 'MTH', 'MZN',
    'RBR', 'SPR', 'VNS', 'ALL', 'NDL',
    'JAI', 'GUR', 'WDL', 'LKO', 'LXR',
    'NDA', 'GRM', 'SDL', 'GZB', 'GND',
    'CAR', 'FAR', 'MPR', 'MHP', 'MEE',
    'DRD', 'UDR', 'JLD', 'LDH', 'JMU',
    'ALW', 'JDH', 'AGR', 'GKP', 'MOR',
    'ATQ', 'KRN', 'BRL', 'KOT', 'ROH',
    'HDN', 'BKR', 'SIK', 'SGG', 'HIS',
    'PTL', 'BLW', 'AJM', 'AMB', 'RKR',
    'HRD', 'RDP', 'RUP', 'YMG', 'BTD',
    'SRG', 'PPT', 'KRK', 'SHM', 'RSH',
    'SON', 'PTK', 'LUD', 'FZD'
]

ALLOWED_DCS_SET = set(ALLOWED_SOURCE_DCS)

# Lowercase variant used by generators (compares sdc.lower())
ALLOWED_DCS_SET_LOWER = set(dc.lower() for dc in ALLOWED_SOURCE_DCS)
