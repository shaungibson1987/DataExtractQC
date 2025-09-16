import pandas as pd

def get_open_ends(df):
    cols = list(df.columns)
    if 'TESTJUMP' in cols:
        start_idx = cols.index('TESTJUMP') + 1
    elif 'TESTLANG' in cols:
        start_idx = cols.index('TESTLANG') + 1
    elif 'ReDemHasRun' in cols:
        start_idx = cols.index('ReDemHasRun') + 1
    elif 'ErrorLog' in cols:
        start_idx = cols.index('ErrorLog') + 1
    else:
        return []
    open_end_cols = []
    ignore_set = {"yes", "no", "dontknow", "_ref"}
    for col in cols[start_idx:]:
        col_lower = col.lower()
        values = df[col].dropna().astype(str)
        values = [v.strip() for v in values if v.strip() != '']
        if not values:
            continue
        if all(v.lower() in ignore_set for v in values):
            continue
        if '.oth' in col_lower or '._oth' in col_lower:
            open_end_cols.append(col)
            continue
        # Only consider as open end if at least one value is longer than 3 characters
        has_long_value = any(len(v) > 3 for v in values)
        if has_long_value and any(not (v.startswith('_') or v[0].isdigit()) for v in values):
            open_end_cols.append(col)
    return open_end_cols