import datetime
import pandas as pd

def resolve_file_bytes(file_bytes):
    """Resolves raw bytes regardless of input type (bytes or file-like object)."""
    if hasattr(file_bytes, "getvalue"):
        return file_bytes.getvalue()
    elif hasattr(file_bytes, "read"):
        try:
            file_bytes.seek(0)
        except Exception:
            pass
        return file_bytes.read()
    else:
        return file_bytes

def parse_date(date_val, default=None):
    """Helper to parse date values into datetime.date, with fallbacks."""
    if date_val is None or pd.isna(date_val):
        return default
    if isinstance(date_val, (datetime.date, datetime.datetime)):
        return date_val.date() if isinstance(date_val, datetime.datetime) else date_val
    try:
        return pd.to_datetime(str(date_val).strip()).date()
    except Exception:
        return default

def get_value_by_aliases(raw_params, aliases, default=None):
    """Helper to check key and all aliases (case-insensitive and stripped) in raw_params."""
    for alias in aliases:
        alias_clean = alias.lower().strip()
        for key, val in raw_params.items():
            if key.lower().strip() == alias_clean:
                return val
    return default

def get_str(raw_params, name, aliases, default):
    """Safely extracts string values from raw_params with aliases and fallback."""
    v = get_value_by_aliases(raw_params, [name] + aliases)
    if v is None or pd.isna(v) or str(v).strip() == "":
        return default
    return str(v).strip()

def get_float(raw_params, name, aliases, default):
    """Safely extracts float values from raw_params with aliases and fallback."""
    v = get_value_by_aliases(raw_params, [name] + aliases)
    if v is None or pd.isna(v):
        return default
    try:
        return float(v)
    except Exception:
        return default

def get_int(raw_params, name, aliases, default):
    """Safely extracts integer values from raw_params with aliases and fallback."""
    v = get_value_by_aliases(raw_params, [name] + aliases)
    if v is None or pd.isna(v):
        return default
    try:
        return int(float(v))
    except Exception:
        return default
