# CouncilWards

A Python tool for looking up council ward information for UK postcodes using the postcodes.io API.

## Features

- Read postcodes from Excel files
- Batch lookup using postcodes.io API (efficient processing of large datasets)
- Extract council ward, district, parliamentary constituency, and geographic data
- Export results to Excel with all original data preserved

## Installation

1. Clone this repository
2. Install dependencies:

```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

```bash
python postcode_lookup.py input.xlsx output.xlsx
```

This will:
1. Read postcodes from `input.xlsx` (expects a column named "postcode")
2. Look up each postcode using the postcodes.io API
3. Save results to `output.xlsx` with additional columns for ward data

### Advanced Options

Specify a custom postcode column name:

```bash
python postcode_lookup.py input.xlsx output.xlsx --postcode-column "Postal Code"
```

Adjust API request delay (in seconds):

```bash
python postcode_lookup.py input.xlsx output.xlsx --delay 0.2
```

### Output Fields

The script adds the following columns to your Excel file:

- `admin_ward` - Administrative ward name
- `admin_district` - Administrative district name
- `parliamentary_constituency` - Parliamentary constituency name
- `region` - Region name
- `country` - Country (England, Scotland, Wales, Northern Ireland)
- `postcode_formatted` - Properly formatted postcode
- `latitude` - Latitude coordinate
- `longitude` - Longitude coordinate

## Creating a Sample File

To test the tool with sample data:

```bash
python create_sample.py
python postcode_lookup.py sample_postcodes.xlsx results.xlsx
```

## API Information

This tool uses the free [postcodes.io API](https://postcodes.io/):
- No API key required
- Batch lookups (up to 100 postcodes per request)
- Comprehensive UK postcode data
- Open source and free to use

## Example

Input Excel file (`input.xlsx`):

| postcode | customer_name |
|----------|---------------|
| SW1A 1AA | John Smith    |
| M1 1AE   | Jane Doe      |

Output Excel file (`output.xlsx`):

| postcode | customer_name | admin_ward | admin_district | parliamentary_constituency |
|----------|---------------|------------|----------------|----------------------------|
| SW1A 1AA | John Smith    | St James's | Westminster    | Cities of London and Westminster |
| M1 1AE   | Jane Doe      | Piccadilly | Manchester     | Manchester Central         |

## Requirements

- Python 3.7+
- pandas
- openpyxl
- requests

## License

MIT
