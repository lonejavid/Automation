"""
Data Transformation Module
Replicates Ctrl+D macro - works for UNLIMITED rows!
"""

import pandas as pd
from datetime import datetime


def transform_raw_car_data(input_file):
    """
    Transform raw HME car data to template format
    
    Args:
        input_file: Path to raw Excel file from HMECloud
    
    Returns:
        DataFrame with transformed data
    """
    print(f"\n   Transforming: {input_file}")
    
    # Read raw file
    raw_df = pd.read_excel(input_file, sheet_name=0, header=None)
    
    # Extract store name from row 3, column 1
    store_name = raw_df.iloc[3, 1]
    print(f"   Store: {store_name}")
    
    # Column mapping (0-indexed)
    column_mapping = {
        0: 'Daypart',           # Column A
        2: 'Departure Time',    # Column C
        4: 'Event Name',        # Column E
        7: 'Cars in Queue',     # Column H
        11: 'Menu Board',       # Column L
        15: 'Greet',            # Column P
        18: 'Service',          # Column S
        19: 'Lane Queue',       # Column T
        22: 'Lane Total',       # Column W
        23: 'Lane Total 2'      # Column X
    }
    
    # Extract data starting from row 7
    transformed_data = []
    for idx in range(7, len(raw_df)):
        row_data = {}
        for col_idx, col_name in column_mapping.items():
            row_data[col_name] = raw_df.iloc[idx, col_idx]
        transformed_data.append(row_data)
    
    # Create DataFrame
    df = pd.DataFrame(transformed_data)
    
    # Add Store Name column
    df.insert(1, 'Store Name', store_name)
    
    # Fill down Daypart column
    df['Daypart'] = df['Daypart'].ffill()
    
    # Remove rows where Event Name is NaN
    df = df[df['Event Name'].notna()].copy()
    df = df.reset_index(drop=True)
    
    print(f"   ✅ Transformed {len(df)} rows (ALL rows, not just 197!)")
    
    return df


if __name__ == "__main__":
    # Test
    test_file = "/Users/sanctum/Desktop/Automation/hme_stores/test.xlsx"
    if os.path.exists(test_file):
        df = transform_raw_car_data(test_file)
        print(df.head())



