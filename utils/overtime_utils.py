"""
utils/overtime_utils.py - small, robust helpers for parsing shifts/time values.

Functions:
- parse_shift(cell_value) -> tuple[start_time, end_time] or None
- to_time(val) -> datetime.time or None

Constants:
- allowed_by_type: mapping work-type -> allowed hours
- day_names: greek day names
- ot_columns: mapping greek day name -> (start_col, end_col)
- DAY_TO_COLS_INT: mapping weekday int (0=Mon..6=Sun) -> (start_col, end_col)
"""
from datetime import datetime, time, timedelta
from typing import Optional, Tuple, Union

Number = Union[int, float]

allowed_by_type = {
    "5ΗΜΕΡΟΣ": 8.0,
    "6ΗΜΕΡΟΣ": 6.67
}

day_names = ["ΔΕΥΤΕΡΑ", "ΤΡΙΤΗ", "ΤΕΤΑΡΤΗ", "ΠΕΜΠΤΗ", "ΠΑΡΑΣΚΕΥΗ", "ΣΑΒΒΑΤΟ", "ΚΥΡΙΑΚΗ"]

ot_columns = {
    "ΔΕΥΤΕΡΑ": ("C", "D"),
    "ΤΡΙΤΗ": ("H", "I"),
    "ΤΕΤΑΡΤΗ": ("M", "N"),
    "ΠΕΜΠΤΗ": ("R", "S"),
    "ΠΑΡΑΣΚΕΥΗ": ("W", "X"),
    "ΣΑΒΒΑΤΟ": ("AB", "AC"),
    "ΚΥΡΙΑΚΗ": ("AG", "AH")
}

# Also export an int-keyed mapping for weekday indices (0=Monday ... 6=Sunday)
DAY_TO_COLS_INT = {
    0: ot_columns["ΔΕΥΤΕΡΑ"],
    1: ot_columns["ΤΡΙΤΗ"],
    2: ot_columns["ΤΕΤΑΡΤΗ"],
    3: ot_columns["ΠΕΜΠΤΗ"],
    4: ot_columns["ΠΑΡΑΣΚΕΥΗ"],
    5: ot_columns["ΣΑΒΒΑΤΟ"],
    6: ot_columns["ΚΥΡΙΑΚΗ"],
}


def _excel_fraction_to_time(frac: float) -> Optional[time]:
    """
    Convert Excel time (fraction of day, or serial) to datetime.time.
    Accepts values like 0.5 (12:00), or floats >1 (will wrap modulo 1).
    """
    try:
        f = float(frac)
    except Exception:
        return None
    # keep only fractional part (time-of-day)
    frac_day = f % 1.0
    total_minutes = int(round(frac_day * 24 * 60))
    total_minutes = total_minutes % (24 * 60)
    hh = total_minutes // 60
    mm = total_minutes % 60
    return time(hour=hh, minute=mm)


def _parse_hhmm(s: str) -> Optional[time]:
    """
    Parse a single HH:MM-like string into time.
    Accepts separators ':' or '.' and optional seconds.
    Returns None on failure.
    """
    if not isinstance(s, str):
        return None
    s = s.strip()
    if not s:
        return None
    # normalize common separators
    s = s.replace(".", ":")
    # try multiple formats
    fmts = ("%H:%M", "%H:%M:%S", "%H:%M.%f")
    for fmt in fmts:
        try:
            dt = datetime.strptime(s, fmt)
            return dt.time()
        except Exception:
            continue
    # fallback: if looks like H:MM or HH:MM without zero-pad
    parts = s.split(":")
    if len(parts) >= 2:
        try:
            hh = int(parts[0])
            mm = int(parts[1])
            hh = hh % 24
            mm = max(0, min(59, mm))
            return time(hour=hh, minute=mm)
        except Exception:
            return None
    return None


def to_time(val: Union[str, time, datetime, Number, None]) -> Optional[time]:
    """
    Normalize various types into a datetime.time:
    - datetime.time -> returned
    - datetime.datetime -> .time()
    - string "HH:MM" or "HH.MM" -> parsed
    - numeric (int/float) -> treated as Excel fraction-of-day
    Returns None if value can't be converted.
    """
    if val is None:
        return None
    if isinstance(val, time):
        return val
    if isinstance(val, datetime):
        return val.time()
    # numeric (Excel)
    if isinstance(val, (int, float)):
        return _excel_fraction_to_time(float(val))
    # str
    if isinstance(val, str):
        s = val.strip()
        # empty or common invalid tokens
        if s == "":
            return None
        # parse
        return _parse_hhmm(s)
    return None


def parse_shift(cell_value: object) -> Optional[Tuple[time, time]]:
    """
    Parse a shift value expressed as 'HH:MM-HH:MM' (possibly with spaces/dots)
    and return (start_time, end_time) as datetime.time objects.
    If parsing fails, returns None.

    Also handles:
    - already parsed tuples/list of two time/datetime objects
    - numeric Excel fractions for start/end (if passed in a tuple/list)
    - strings like '08:00 - 16:00', '08.00-00:30'
    - crossing-midnight shifts are represented as times (caller may combine with date)
    """
    # If already a pair/sequence, try to convert both items
    if isinstance(cell_value, (list, tuple)) and len(cell_value) == 2:
        start = to_time(cell_value[0])
        end = to_time(cell_value[1])
        if start and end:
            return start, end
        return None

    # If it's a string, split by '-' (allow spaces around)
    try:
        s = str(cell_value).strip()
    except Exception:
        return None
    if not s:
        return None

    # common separator is '-', but protect against other dashes / unicode
    # normalize different dash characters
    for dash in ("\u2013", "\u2014", "\u2212"):  # –, —, −
        s = s.replace(dash, "-")

    parts = [p.strip() for p in s.split("-") if p.strip() != ""]
    if len(parts) != 2:
        return None

    start = to_time(parts[0])
    end = to_time(parts[1])
    if not start or not end:
        return None

    return start, end