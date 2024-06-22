"""
    pip install pandas openpyxl
"""

import pandas as pd
import sys
import datetime


WISE_COLUMN_ORDER = [
    'Finished on',
    'Source amount (after fees)',
    'Source currency',
    'Target name',
    'Reference',
    'Source name',
    'ID',
    'Exchange rate',
]

WISE_COLUMNS_TO_DROP = [
    'Status',
    'Direction',
    'Created on',
    'Source fee amount',
    'Source fee currency',
    'Target fee amount',
    'Target fee currency',
    'Target amount (after fees)',
    'Target currency',
    'Batch',
]

ANZ_COLUMN_ORDER = [
    'Transaction Date',
    'Amount',
    'Details',
    'Particulars',
    'Code',
    'Reference',
    'Type',
    'Conversion Charge',
    'Foreign Currency Amount',
]

ANZ_COLUMNS_TO_DROP = [
    'Processed Date',
    'Balance',
    'To/From Account Number',
]


def format_transactions(file_path, columns_to_drop):
    # Read the Excel file
    df = pd.read_excel(file_path)
    print(df)
    print(df.columns.values)

    # Remove unused columns
    df.drop(columns=columns_to_drop, inplace=True)
    print(df.columns.values)

    return df


def write_dataframe_to_excel(output_file_path, wise_df, anz_df):
    writer = pd.ExcelWriter(output_file_path, engine='xlsxwriter')
    wise_df.to_excel(writer, sheet_name='Wise', index=False, columns=WISE_COLUMN_ORDER)
    anz_df.to_excel(writer, sheet_name='ANZ', index=False, columns=ANZ_COLUMN_ORDER)
    writer._save()

if __name__ == '__main__':
    args = sys.argv
    wise_transactions_file = args[1]
    anz_transactions_file = args[2]

    wise_df = format_transactions(wise_transactions_file, WISE_COLUMNS_TO_DROP)
    anz_df = format_transactions(anz_transactions_file, ANZ_COLUMNS_TO_DROP)

    now = datetime.datetime.now()
    formatted_datetime = now.strftime('%Y%m%d %H%M')
    output_file = f'{formatted_datetime} Formatted Transactions.xlsx'

    write_dataframe_to_excel(output_file, wise_df, anz_df)

    # xls = pd.ExcelFile(output_file)
    # df1 = pd.read_excel(xls, 'Wise')  # Read data from Sheet1
    # df2 = pd.read_excel(xls, 'ANZ')  # Read data from Sheet2

    # # Combine the rows from both dataframes
    # combined_df = pd.concat([df1, df2], ignore_index=True)

    # # Write the combined data to a new Excel file
    # with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    #     combined_df.to_excel(writer, sheet_name='All Transactions', index=False)
