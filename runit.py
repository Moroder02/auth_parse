#!/usr/bin/env python3
"""
Script for converting exported data from vaultwarden (JSON) to Excel format.
"""

import json
import openpyxl
from openpyxl import Workbook


def load_json_data(filepath):
    """Loads data from a JSON file."""
    with open(filepath, 'r', encoding='utf-8') as f:
        return json.load(f)


def create_collection_id_to_name_map(collections):
    """Creates a dictionary to map collection ID to its name."""
    return {coll['id']: coll['name'] for coll in collections}


def calculate_password_length(password):
    """Calculates the password length."""
    if password is None:
        return 0
    return len(str(password))


def extract_items_data(data):
    """Extracts data from items and maps them to collections."""
    collection_map = create_collection_id_to_name_map(data.get('collections', []))
    items = data.get('items', [])

    result = []
    for item in items:
        name = item.get('name', '')
        note = item.get('notes', '')
        login_info = item.get('login', {})
        username = login_info.get('username', '') if login_info else ''
        password = login_info.get('password', '') if login_info else ''

        # Calculate password length
        password_length = calculate_password_length(password)

        # Get list of collections for this item
        collection_ids = item.get('collectionIds', [])

        if collection_ids:
            # If item has collections, create a row for each collection
            for coll_id in collection_ids:
                collection_name = collection_map.get(coll_id, '')
                result.append({
                    'collection_name': collection_name,
                    'name': name,
                    'note': note,
                    'username': username,
                    'password': password,
                    'password_length': password_length
                })
        else:
            # If item has no collections, add a single row without collection name
            result.append({
                'collection_name': '',
                'name': name,
                'note': note,
                'username': username,
                'password': password,
                'password_length': password_length
            })

    return result


def write_to_excel(items_data, output_filepath):
    """Writes data to an Excel file."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Vaultwarden Export"

    # Column headers
    headers = ['Collection Name (collection_name)',
               'Item Name (name)',
               'Item Note (note)',
               'Username (login)',
               'Password (password)',
               'Password Length']

    ws.append(headers)

    # Data rows
    for item in items_data:
        ws.append([
            item['collection_name'],
            item['name'],
            item['note'],
            item['username'],
            item['password'],
            item['password_length']
        ])

    # Auto-adjust column widths
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column_letter].width = adjusted_width

    wb.save(output_filepath)
    print(f"File successfully created: {output_filepath}")


def main():
    input_file = 'example.json'
    output_file = 'vaultwarden_export.xlsx'

    # Load data
    print(f"Reading file {input_file}...")
    data = load_json_data(input_file)

    # Extract data
    print("Processing data...")
    items_data = extract_items_data(data)

    # Write to Excel
    print(f"Creating Excel file {output_file}...")
    write_to_excel(items_data, output_file)

    print(f"Done! Processed records: {len(items_data)}")


if __name__ == '__main__':
    main()