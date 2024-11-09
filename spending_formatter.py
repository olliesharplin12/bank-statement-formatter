"""
    pip install openpyxl
"""

import sys
import datetime
import csv
import openpyxl
import os


# [
#     'Finished on',
#     'Source amount (after fees)',
#     'Source currency',
#     'Target name',
#     'Reference',
#     'Source name',
#     'ID',
#     'Exchange rate',
# ]
WISE_COLUMN_ORDER = ['', 4, '', 10, 11, 12, 16, 9, 0, 15]

# WISE_COLUMNS_TO_DROP = [  TODO: Remove
#     'Status',
#     'Direction',
#     'Created on',
#     'Source fee amount',
#     'Source fee currency',
#     'Target fee amount',
#     'Target fee currency',
#     'Target amount (after fees)',
#     'Target currency',
#     'Batch',
# ]

# [
#     'Transaction Date',
#     'Amount',
#     'Details',
#     'Particulars',
#     'Code',
#     'Reference',
#     'Type',
#     'Conversion Charge',
#     'Foreign Currency Amount',
# ]
ANZ_COLUMN_ORDER = ['', 6, '', 5, 'NZD', 1, 2, 3, 4, 0, 8, 7]

# SPLITWISE_COLUMN_ORDER = [
#     'Date',
#     'Ollie Sharplin',
#     'Currency',
#     'Description',
# ]
SPLITWISE_COLUMN_ORDER = ['', 0, '', 5, 4, 1, 7]

OUTPUT_HEADER = ['Tag', 'Date', 'GBP Amount', 'Amount', 'Currency', 'Description', 'D2', 'D3', 'Type', 'Exchange Rate']

# def format_transactions(file_path, columns_to_drop):
#     # Read the Excel file
#     df = pd.read_excel(file_path)
#     print(df)
#     print(df.columns.values)

#     # Remove unused columns
#     df.drop(columns=columns_to_drop, inplace=True)
#     print(df.columns.values)

#     return df


def read_csv(file_path, row_start_index):
    rows = []
    with open(file_path, newline='', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)
        skipped_rows = 0
        for row in reader:
            if skipped_rows < row_start_index:
                skipped_rows += 1
                continue
            rows.append(row)
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


# def write_dataframe_to_excel(output_file_path, wise_df, anz_df):
#     writer = pd.ExcelWriter(output_file_path, engine='xlsxwriter')
#     wise_df.to_excel(writer, sheet_name='Wise', index=False, columns=WISE_COLUMN_ORDER)
#     anz_df.to_excel(writer, sheet_name='ANZ', index=False, columns=ANZ_COLUMN_ORDER)
#     writer._save()


if __name__ == '__main__':
    args = sys.argv

    start_date = datetime.datetime.strptime(args[4], '%Y-%m-%d')
    end_date = datetime.datetime.strptime(args[5], '%Y-%m-%d').replace(hour=23, minute=59, second=59)
    print(start_date, end_date)

    wise_transactions_file = args[1]
    anz_transactions_file = args[2]
    splitwise_export_dir = args[3]
    
    print('\nWISE ROWS')
    wise_rows_unordered = read_csv(wise_transactions_file, 1)
    wise_rows_filtered = filter_rows_by_date(wise_rows_unordered, 4, '%Y-%m-%d %H:%M:%S', start_date, end_date)
    wise_rows = swap_columns(wise_rows_filtered, WISE_COLUMN_ORDER)

    for row in wise_rows:
        print(row)

    print('\nANZ ROWS')
    anz_rows_unordered = read_csv(anz_transactions_file, 1)
    anz_rows_filtered = filter_rows_by_date(anz_rows_unordered, 6, '%d/%m/%Y', start_date, end_date)
    anz_rows = swap_columns(anz_rows_filtered, ANZ_COLUMN_ORDER)
    
    for row in anz_rows:
        print(row)

    print('\nSPLIT ROWS')
    splitwise_rows_all = combine_splitwise_transactions(splitwise_export_dir)
    splitwise_rows_filtered = filter_rows_by_date(splitwise_rows_all, 0, '%Y-%m-%d', start_date, end_date)
    splitwise_rows = swap_columns(splitwise_rows_filtered, SPLITWISE_COLUMN_ORDER)

    for row in splitwise_rows:
        print(row)
    
    transactions = [OUTPUT_HEADER] + [[]] + wise_rows + [[]] + anz_rows + [[]] + splitwise_rows

    # Write to Excel
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    for row in transactions:
        sheet.append(row)

    output_directory = os.path.join(os.path.expanduser('~'), 'Downloads')
    date_format = '%Y-%m-%d'
    output_filename = f'Spending_{start_date.strftime(date_format)}_to_{end_date.strftime(date_format)}.xlsx'
    workbook.save(os.path.join(output_directory, output_filename))


    # wise_df = format_transactions(wise_transactions_file, WISE_COLUMNS_TO_DROP)
    # wise_rows = wise_df.values().tolist()
    # anz_df = format_transactions(anz_transactions_file, ANZ_COLUMNS_TO_DROP)

    # now = datetime.datetime.now()
    # formatted_datetime = now.strftime('%Y%m%d %H%M')
    # output_file = f'{formatted_datetime} Formatted Transactions.xlsx'

    # write_dataframe_to_excel(output_file, wise_df, anz_df)

    # xls = pd.ExcelFile(output_file)
    # df1 = pd.read_excel(xls, 'Wise')  # Read data from Sheet1
    # df2 = pd.read_excel(xls, 'ANZ')  # Read data from Sheet2

    # # Combine the rows from both dataframes
    # combined_df = pd.concat([df1, df2], ignore_index=True)

    # # Write the combined data to a new Excel file
    # with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    #     combined_df.to_excel(writer, sheet_name='All Transactions', index=False)
