#!/usr/bin/env python3
"""
Postcode to Council Ward Lookup Tool

This script reads postcodes from an Excel file and looks up council ward information
using the postcodes.io API. Results are saved to a new Excel file.

Usage:
    python postcode_lookup.py input.xlsx output.xlsx [--postcode-column COLUMN_NAME]
"""

import argparse
import sys
from typing import List, Dict, Optional
import pandas as pd
import requests
from time import sleep


class PostcodeLookup:
    """Handle postcode lookups using the postcodes.io API."""

    BASE_URL = "https://api.postcodes.io"
    BATCH_SIZE = 100  # API allows up to 100 postcodes per batch request

    def __init__(self, delay: float = 0.1):
        """
        Initialize the PostcodeLookup.

        Args:
            delay: Delay in seconds between API requests (to be respectful)
        """
        self.delay = delay
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'CouncilWards-Lookup/1.0',
            'Content-Type': 'application/json'
        })

    def normalize_postcode(self, postcode: str) -> str:
        """Normalize postcode by removing spaces and converting to uppercase."""
        if pd.isna(postcode):
            return ""
        return str(postcode).replace(" ", "").strip().upper()

    def lookup_single(self, postcode: str) -> Optional[Dict]:
        """
        Look up a single postcode.

        Args:
            postcode: The postcode to look up

        Returns:
            Dictionary with postcode data or None if not found
        """
        normalized = self.normalize_postcode(postcode)
        if not normalized:
            return None

        try:
            response = self.session.get(
                f"{self.BASE_URL}/postcodes/{normalized}",
                timeout=10
            )

            if response.status_code == 200:
                data = response.json()
                if data.get('status') == 200 and data.get('result'):
                    return data['result']
            elif response.status_code == 404:
                return None
            else:
                print(f"Warning: API returned status {response.status_code} for {postcode}")
                return None

        except requests.exceptions.RequestException as e:
            print(f"Error looking up {postcode}: {e}")
            return None

    def lookup_batch(self, postcodes: List[str]) -> Dict[str, Optional[Dict]]:
        """
        Look up multiple postcodes in a single API call.

        Args:
            postcodes: List of postcodes to look up (max 100)

        Returns:
            Dictionary mapping postcodes to their data
        """
        if len(postcodes) > self.BATCH_SIZE:
            raise ValueError(f"Batch size cannot exceed {self.BATCH_SIZE}")

        # Normalize postcodes
        normalized_map = {self.normalize_postcode(pc): pc for pc in postcodes}
        normalized_postcodes = [pc for pc in normalized_map.keys() if pc]

        if not normalized_postcodes:
            return {pc: None for pc in postcodes}

        results = {}

        try:
            response = self.session.post(
                f"{self.BASE_URL}/postcodes",
                json={"postcodes": normalized_postcodes},
                timeout=30
            )

            if response.status_code == 200:
                data = response.json()
                if data.get('status') == 200 and data.get('result'):
                    for item in data['result']:
                        query = item.get('query', '').upper()
                        original = normalized_map.get(query)
                        if original:
                            results[original] = item.get('result')
            else:
                print(f"Warning: Batch API returned status {response.status_code}")
                # Fall back to individual lookups
                for pc in postcodes:
                    sleep(self.delay)
                    results[pc] = self.lookup_single(pc)

        except requests.exceptions.RequestException as e:
            print(f"Error in batch lookup: {e}")
            print("Falling back to individual lookups...")
            for pc in postcodes:
                sleep(self.delay)
                results[pc] = self.lookup_single(pc)

        return results

    def lookup_all(self, postcodes: List[str], show_progress: bool = True) -> List[Dict]:
        """
        Look up all postcodes with batch processing.

        Args:
            postcodes: List of all postcodes to look up
            show_progress: Whether to show progress messages

        Returns:
            List of dictionaries with lookup results
        """
        results = []
        total = len(postcodes)

        # Process in batches
        for i in range(0, total, self.BATCH_SIZE):
            batch = postcodes[i:i + self.BATCH_SIZE]
            batch_num = i // self.BATCH_SIZE + 1
            total_batches = (total + self.BATCH_SIZE - 1) // self.BATCH_SIZE

            if show_progress:
                print(f"Processing batch {batch_num}/{total_batches} ({len(batch)} postcodes)...")

            batch_results = self.lookup_batch(batch)
            results.extend(batch_results.values())

            # Be respectful to the API
            if i + self.BATCH_SIZE < total:
                sleep(self.delay)

        return results


def extract_fields(result: Optional[Dict]) -> Dict[str, Optional[str]]:
    """
    Extract relevant fields from API result.

    Args:
        result: API result dictionary

    Returns:
        Dictionary with extracted fields
    """
    if not result:
        return {
            'admin_ward': None,
            'admin_district': None,
            'parliamentary_constituency': None,
            'region': None,
            'country': None,
            'postcode_formatted': None,
            'latitude': None,
            'longitude': None
        }

    return {
        'admin_ward': result.get('admin_ward'),
        'admin_district': result.get('admin_district'),
        'parliamentary_constituency': result.get('parliamentary_constituency'),
        'region': result.get('region'),
        'country': result.get('country'),
        'postcode_formatted': result.get('postcode'),
        'latitude': result.get('latitude'),
        'longitude': result.get('longitude')
    }


def process_excel(
    input_file: str,
    output_file: str,
    postcode_column: str = 'postcode',
    delay: float = 0.1
) -> None:
    """
    Process an Excel file and add council ward information.

    Args:
        input_file: Path to input Excel file
        output_file: Path to output Excel file
        postcode_column: Name of the column containing postcodes
        delay: Delay between API requests in seconds
    """
    print(f"Reading Excel file: {input_file}")

    try:
        df = pd.read_excel(input_file)
    except FileNotFoundError:
        print(f"Error: File '{input_file}' not found.")
        sys.exit(1)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        sys.exit(1)

    # Check if postcode column exists
    if postcode_column not in df.columns:
        print(f"Error: Column '{postcode_column}' not found in Excel file.")
        print(f"Available columns: {', '.join(df.columns)}")
        sys.exit(1)

    print(f"Found {len(df)} rows with postcodes")

    # Initialize lookup service
    lookup = PostcodeLookup(delay=delay)

    # Get postcodes
    postcodes = df[postcode_column].tolist()

    # Perform lookups
    print("\nLooking up postcodes...")
    results = lookup.lookup_all(postcodes, show_progress=True)

    # Extract fields and add to dataframe
    print("\nProcessing results...")
    extracted = [extract_fields(r) for r in results]
    result_df = pd.DataFrame(extracted)

    # Combine with original data
    output_df = pd.concat([df, result_df], axis=1)

    # Count successful lookups
    successful = sum(1 for r in results if r is not None)
    print(f"\nSuccessfully looked up {successful}/{len(postcodes)} postcodes")

    # Save to Excel
    print(f"Saving results to: {output_file}")
    output_df.to_excel(output_file, index=False)
    print("Done!")


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description='Look up council ward information for postcodes in an Excel file'
    )
    parser.add_argument(
        'input_file',
        help='Input Excel file containing postcodes'
    )
    parser.add_argument(
        'output_file',
        help='Output Excel file to save results'
    )
    parser.add_argument(
        '--postcode-column',
        default='postcode',
        help='Name of the column containing postcodes (default: postcode)'
    )
    parser.add_argument(
        '--delay',
        type=float,
        default=0.1,
        help='Delay between API requests in seconds (default: 0.1)'
    )

    args = parser.parse_args()

    process_excel(
        args.input_file,
        args.output_file,
        args.postcode_column,
        args.delay
    )


if __name__ == '__main__':
    main()
