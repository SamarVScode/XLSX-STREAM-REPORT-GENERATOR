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

# Central registry for Mini-DCs, Spokes, and Hub Aliases mapped to Canonical Parent DCs
DC_ALIASES = {
    'CAR-KHR': 'CAR',
    'CAR_KHR': 'CAR',
    'CARKHR':  'CAR',
}

DC_ALIASES_MAP = {k.strip().upper(): v.strip().upper() for k, v in DC_ALIASES.items()}


def normalize_dc_code(raw_dc) -> str:
    """
    Normalizes any DC code or mini-hub alias to its canonical parent DC code.
    """
    if raw_dc is None:
        return ""
    s = str(raw_dc).strip().upper()
    if not s:
        return ""
    if s in DC_ALIASES_MAP:
        return DC_ALIASES_MAP[s]
    alt1 = s.replace('_', '-')
    if alt1 in DC_ALIASES_MAP:
        return DC_ALIASES_MAP[alt1]
    alt2 = s.replace('-', '_')
    if alt2 in DC_ALIASES_MAP:
        return DC_ALIASES_MAP[alt2]
    return s


def is_allowed_dc(raw_dc) -> bool:
    """
    Returns True if the raw DC (or its alias) resolves to an allowed server DC.
    """
    canonical = normalize_dc_code(raw_dc)
    return canonical in ALLOWED_DCS_SET if canonical else False

