"""
    pip install openpyxl
"""

import sys
import datetime
import csv
import openpyxl
import os


WISE_COLUMN_ORDER = ['', 4, '', 10, 11, 12, 16, 9, 0, 15]
ANZ_DEBIT_COLUMN_ORDER = ['', 6, '', 5, 'NZD', 1, 2, 3, 4, 0, 8, 7]
ANZ_CREDIT_COLUMN_ORDER = ['', 4, '', 2, 'NZD', 0, '', 3, '', 6]
SPLITWISE_COLUMN_ORDER = ['', 0, '', 5, 4, 1, 7]

OUTPUT_HEADER = ['Tag', 'Date', 'Amount (NZD)', 'Amount', 'Currency', 'Description', 'D2', 'D3', 'Type', 'Exchange Rate']


def read_csv(file_path, row_start_index):
    rows = []
    if len(file_path) <= 1:
        return rows
    with open(file_path, newline='', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)
        skipped_rows = 0
        for row in reader:
            if skipped_rows < row_start_index:
                skipped_rows += 1
                continue
            rows.append(row)
    return rows


def invert_wise_amount(rows: list[list]):
    for row in rows:
        direction = row[2]
        if direction == "OUT":
            row[10] = -float(row[10])
    return rows


def invert_anz_credit_amount(rows: list[list]):
    for row in rows:
        trans_type = row[1]
        if trans_type != "C":
            row[2] = -float(row[2])
    return rows


def combine_splitwise_transactions(export_dir):
    all_files = os.listdir(export_dir)
    csv_files = [file for file in all_files if file.lower().endswith('.csv')]
    all_rows = []
    for filename in csv_files:
        file_path = os.path.join(export_dir, filename)
        rows = read_csv(file_path, 2)
        name = rows[0][6]
        rows = rows[1:]
        for row in rows:
            if len(row) < 4:
                continue
            row.append(name)
            all_rows.append(row)
    all_rows.sort(key=lambda x: x[0], reverse=True)
    return all_rows


def swap_columns(rows: list[list], order: list[int]):
    ordered_rows = []
    for row in rows:
        ordered_row = []
        for index in order:
            if type(index) == type(''):
                ordered_row.append(index)
            else:
                ordered_row.append(row[index])
        ordered_rows.append(ordered_row)
    return ordered_rows


def filter_rows_by_date(rows: list, date_index: int, date_format: str, start_date: datetime.datetime, end_date: datetime.datetime, prints=False):
    valid_rows = []
    for row in rows:
        if prints:
            print(row)
        row_date = datetime.datetime.strptime(row[date_index], date_format)
        if row_date >= start_date and row_date <= end_date:
            valid_rows.append(row)
    return valid_rows


def set_row_types(rows: list[list], transaction_type):
    for row in rows:
        if transaction_type == "WISE":
            row[1] = datetime.datetime.strptime(row[1], "%Y-%m-%d %H:%M:%S").strftime('%d/%m/%Y') # Date
        elif transaction_type == "SPLITWISE":
            row[1] = datetime.datetime.strptime(row[1], "%Y-%m-%d").strftime('%d/%m/%Y') # Date
        row[3] = float(row[3]) # Amount
    return rows


if __name__ == '__main__':
    args = sys.argv

    wise_transactions_file = args[1]
    anz_debit_transactions_file = args[2]
    anz_credit_transactions_file = args[3]
    splitwise_export_dir = args[4]

    start_date = datetime.datetime.strptime(args[5], '%Y-%m-%d')
    end_date = datetime.datetime.strptime(args[6], '%Y-%m-%d').replace(hour=23, minute=59, second=59)
    print(start_date, end_date)
    
    print('\nWISE ROWS')
    wise_rows_unordered = read_csv(wise_transactions_file, 1)
    wise_rows_inverted = invert_wise_amount(wise_rows_unordered)
    wise_rows_filtered = filter_rows_by_date(wise_rows_inverted, 4, '%Y-%m-%d %H:%M:%S', start_date, end_date)
    wise_rows_ordered = swap_columns(wise_rows_filtered, WISE_COLUMN_ORDER)
    wise_rows = set_row_types(wise_rows_ordered, "WISE")

    for row in wise_rows:
        print(row)

    print('\nANZ DEBIT ROWS')
    anz_debit_rows_unordered = read_csv(anz_debit_transactions_file, 1)
    anz_debit_rows_filtered = filter_rows_by_date(anz_debit_rows_unordered, 6, '%d/%m/%Y', start_date, end_date)
    anz_debit_rows_ordered = swap_columns(anz_debit_rows_filtered, ANZ_DEBIT_COLUMN_ORDER)
    anz_debit_rows = set_row_types(anz_debit_rows_ordered, "ANZ_DEBIT")
    
    for row in anz_debit_rows:
        print(row)

    print('\nANZ CREDIT ROWS')
    anz_credit_rows_unordered = read_csv(anz_credit_transactions_file, 1)
    anz_credit_rows_inverted = invert_anz_credit_amount(anz_credit_rows_unordered)
    anz_credit_rows_filtered = filter_rows_by_date(anz_credit_rows_inverted, 4, '%d/%m/%Y', start_date, end_date)
    anz_credit_rows_ordered = swap_columns(anz_credit_rows_filtered, ANZ_CREDIT_COLUMN_ORDER)
    anz_credit_rows = set_row_types(anz_credit_rows_ordered, "ANZ_CREDIT")
    
    for row in anz_credit_rows:
        print(row)

    print('\nSPLIT ROWS')
    splitwise_rows_all = combine_splitwise_transactions(splitwise_export_dir)
    splitwise_rows_filtered = filter_rows_by_date(splitwise_rows_all, 0, '%Y-%m-%d', start_date, end_date)
    splitwise_rows_ordered = swap_columns(splitwise_rows_filtered, SPLITWISE_COLUMN_ORDER)
    splitwise_rows = set_row_types(splitwise_rows_ordered, "SPLITWISE")

    for row in splitwise_rows:
        print(row)
    
    transactions = [OUTPUT_HEADER] + [[]] + wise_rows + [[]] + anz_debit_rows + [[]] + anz_credit_rows + [[]] + splitwise_rows

    # Write to Excel
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    for row in transactions:
        sheet.append(row)

    output_directory = os.path.join(os.path.expanduser('~'), 'Downloads')
    date_format = '%Y-%m-%d'
    output_filename = f'Spending_{start_date.strftime(date_format)}_to_{end_date.strftime(date_format)}.xlsx'
    workbook.save(os.path.join(output_directory, output_filename))
