# ============================================================
# DC Configuration — ei_stream_server
# ============================================================
# Allowed Source DCs across ALL server report generators
# ============================================================

ALLOWED_SOURCE_DCS = [
    'GZB', 'NDA', 'GND', 'WDL', 'MEE',
    'AGR', 'MTH', 'SPR', 'MZN', 'HPA',
    'FZD', 'LKO', 'BRL', 'MOR', 'GKP',
    'DRD', 'HDN', 'HRD', 'RDP', 'RSH',
    'RKR', 'GRM', 'GUR', 'KOT', 'JDH',
    'UDR', 'AJM', 'BKR', 'BLW', 'SIK',
    'SGG', 'ALW', 'HIS', 'ROH', 'SON',
    'PPT', 'KRN', 'AMB', 'YMG', 'KRK',
    'JMU', 'ATQ', 'SRG', 'PTK', 'NBZ',
    'LXR', 'FAR', 'SDL', 'JAI', 'ALL',
    'KNP', 'VNS', 'MAU', 'MRZ', 'AYP',
    'ALG', 'DEO', 'JNP', 'JHS', 'RBR',
    'BTD', 'CAR', 'JLD', 'LDH', 'LUD',
    'PTL', 'RUP', 'SHM', 'MPR', 'MHP',
    'NDL'
]

ALLOWED_DCS_SET = set(ALLOWED_SOURCE_DCS)

# Lowercase variant used by generators (compares sdc.lower())
ALLOWED_DCS_SET_LOWER = set(dc.lower() for dc in ALLOWED_SOURCE_DCS)
