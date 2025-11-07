"""
Create sample Raw Car Data file for testing
"""

import pandas as pd
from datetime import datetime

# Create sample data matching HMECloud format
data = []

# Header rows (0-6)
data.append(['Raw Car Data Report'] + [None]*23)
data.append([None]*24)
data.append([None]*24)
data.append(['Store:', '(Ungrouped) 5 Mandela - KFC', None, 'Start Time:', None, 'Nov 04, 2025 09:00 AM'] + [None]*18)
data.append(['Brand:', 'KFC'] + [None]*22)
data.append([None]*24)
data.append(['Daypart', None, 'Departure Time', None, 'Event Name', None, None, 'Cars in Queue', None, None, None, 'Menu Board', None, None, None, 'Greet', None, None, 'Service', 'Lane Queue', None, None, 'Lane Total', 'Lane Total 2'])

# Data rows (starting from row 7)
dayparts = ['6:00AM - 10:59AM', '11:00AM - 1:59PM', '2:00PM - 4:59PM', '5:00PM - 7:59PM', '8:00PM - 3:59AM']

for daypart_idx, daypart in enumerate(dayparts):
    # Add 20 rows per daypart
    for i in range(20):
        if i == 0:
            # First row of daypart has the daypart name
            row = [daypart]
        else:
            # Other rows have empty first column
            row = [None]
        
        # Add other columns
        row.extend([
            None,  # Col 1
            f'2025-11-04 {10+daypart_idx}:{15+i}:00 AM',  # Departure Time (col 2)
            None,  # Col 3
            'Car_Departure',  # Event Name (col 4)
            None, None,  # Cols 5-6
            3,  # Cars in Queue (col 7)
            None, None, None,  # Cols 8-10
            150 + i*10,  # Menu Board (col 11)
            None, None, None,  # Cols 12-14
            150 + i*10,  # Greet (col 15)
            None, None,  # Cols 16-17
            200 + i*15,  # Service (col 18)
            20 + i*5,  # Lane Queue (col 19)
            None, None,  # Cols 20-21
            370 + i*20,  # Lane Total (col 22)
            370 + i*20   # Lane Total 2 (col 23)
        ])
        
        data.append(row)

# Create DataFrame
df = pd.DataFrame(data)

# Save to Excel
output_file = "/Users/sanctum/Desktop/Automation/downloads/Sample Raw Car Data - Mandela.xlsx"
df.to_excel(output_file, index=False, header=False)

print(f"✅ Created sample file: {output_file}")
print(f"   Rows: {len(df)}")
print(f"   Columns: {len(df.columns)}")



