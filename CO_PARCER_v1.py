import xlwings as xw
import pandas as pd
import numpy as np
import re
import getpass
import os


def read_excel_data_with_xlwings():
    app = xw.apps.active
    wb = app.books.active
    ws = wb.sheets[0]

    art = []
    name = []
    count = []
    units = []
    price_final = []
    CO_date = []

    global pf_col, digits
    pf_col = 0
    max_col = max(
        (col for col in range(3, ws.range('C15').end('right').column + 5) if ws.cells(15, col).value is not None),
        default=0)
    max_row = ws.range('C' + str(ws.cells.last_cell.row)).end('up').row + 1

    # Find the column with 'Примечания'
    for col in range(3, max_col + 1):
        if ws.cells(15, col).value == 'Примечания':
            com_col = col
            pf_col = com_col - 1

    # Find the row with 'ИТОГО:'
    for col in range(3, com_col):
        for row in range(15, max_row):
            if ws.cells(row, col).value == 'ИТОГО:':
                res_row = row

    e8value = ws.range('E8').value

    # Extract digits using regular expression
    digits = int(re.search(r'№ (\d+)\s*', str(e8value)).group(1) if e8value else None)

    # Extract date using regular expression
    extracted_date = re.search(r'\b(\d{2}\.\d{2}\.\d{4})\b', e8value).group(1)

    # Read data from Excel
    for row in range(16, res_row):
        art.append(ws.cells(row, 3).value)
        name.append(ws.cells(row, 4).value)
        count.append(ws.cells(row, 5).value)
        units.append(ws.cells(row, 6).value)
        price_final.append(ws.cells(row, pf_col).value)
        CO_date.append(extracted_date)


    count = [int(value) if value is not None else 0 for value in count]
    data = np.array([art, name, count, units, price_final, CO_date]).T.tolist()
    data_f = list(filter(lambda x: x[3] is not None, data))


    # Convert data to DataFrame
    columns = ['Артикул', 'Наименование', 'Количество', 'Единица измерения', 'Финальная стоимость, руб.', 'Дата модификации']
    index = [digits for _ in range(len(data_f))]
    excel_data = pd.DataFrame(data_f, columns=columns, index=index)

    # Set display options
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)
    pd.set_option('display.max_colwidth', None)

    # Insert Presale column
    global username
    username = getpass.getuser()
    excel_data.insert(0, 'Presale', username)

    return digits, excel_data


def update_dataset(csv_file, new_data):
    try:
        if os.path.exists(csv_file):
            # Read existing dataset from CSV
            df = pd.read_csv(csv_file, index_col='Project_ID')
        else:
            df = pd.DataFrame()

        # Check if project ID exists in the dataset and delete corresponding rows
        if digits in df.index:
            df = df.drop(digits)

        # Append new data to the dataset
        df = pd.concat([df, new_data])
        df.index = df.index.rename('Project_ID')

        # Save updated dataset to CSV
        df.to_csv(csv_file)

        print(f"Dataset updated with project ID: {project_id}")
    except Exception as e:
        print(f"Error updating dataset: {e}")


# Read data from the active Excel instance using xlwings
project_id, excel_data = read_excel_data_with_xlwings()

# Define CSV file for the dataset
folder_path = rf"C:\Users\{username}\YandexDisk\C3 Presale\Прочее\CO_DATABASE\CO_PARCER"
csv_filename = f"dataset1_{username}.csv"
dataset_csv = os.path.join(folder_path, csv_filename)

# Update dataset based on project ID
update_dataset(dataset_csv, excel_data)

# Convert data to DataFrame
print(excel_data)
