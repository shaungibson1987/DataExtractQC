import pandas as pd

def print_open_ends(file_path, incentive_col):
    df = pd.read_excel(file_path, dtype=str)
    cols = list(df.columns)
    try:
        start_idx = cols.index(incentive_col) + 1
    except ValueError:
        print(f"Column '{incentive_col}' not found.")
        return
    open_end_cols = []
    for col in cols[start_idx:]:
        col_lower = col.lower()
        # Get non-blank values only
        values = df[col].dropna().astype(str)
        values = [v for v in values if v.strip() != '']
        # Skip column if all cells are blank
        if not values:
            continue
        # Include columns with .oth or ._oth in the name, but only if not all blank
        if '.oth' in col_lower or '._oth' in col_lower:
            open_end_cols.append(col)
            continue
        # Check if any value does NOT start with a digit or '_'
        if any(not (v.startswith('_') or v[0].isdigit()) for v in values):
            open_end_cols.append(col)
    # Print in include.txt format (one column per line, no semicolons)
    for col in open_end_cols:
        print(col)

# Example usage:
print_open_ends(r'C:\Yonder\Box\Yonder Data Solutions\DataExtractQC\P026776\P026776.xlsx', 'TESTJUMP')