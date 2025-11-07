#!/usr/bin/env python3
"""
Create a sample Excel file with postcodes for testing.
"""

import pandas as pd

# Sample postcodes from various UK locations
sample_postcodes = [
    'SW1A 1AA',  # Westminster, London
    'M1 1AE',    # Manchester
    'B1 1AA',    # Birmingham
    'EH1 1YZ',   # Edinburgh
    'CF10 1DD',  # Cardiff
    'BT1 1AA',   # Belfast
    'LS1 1AA',   # Leeds
    'L1 1AA',    # Liverpool
    'NE1 1AA',   # Newcastle
    'BS1 1AA',   # Bristol
]

# Create DataFrame
df = pd.DataFrame({
    'postcode': sample_postcodes,
    'description': [
        'Westminster Parliament',
        'Manchester City Centre',
        'Birmingham City Centre',
        'Edinburgh City Centre',
        'Cardiff City Centre',
        'Belfast City Centre',
        'Leeds City Centre',
        'Liverpool City Centre',
        'Newcastle City Centre',
        'Bristol City Centre'
    ]
})

# Save to Excel
output_file = 'sample_postcodes.xlsx'
df.to_excel(output_file, index=False)
print(f"Created {output_file} with {len(df)} sample postcodes")
