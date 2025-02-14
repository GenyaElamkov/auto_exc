import pandas as pd
import numpy as np


from openpyxl import load_workbook
from auto_exc import find_xlsx_files


def find_last_row(filename, start_row=43):
    wb = load_workbook(filename=filename, read_only=True)
    sheet_ranges = wb.active
    while True:
        start_row += 1

        if sheet_ranges[f"A{start_row}"].value is None:
            wb.close()
            return (filename, start_row)

    


def main():

    # filename = r"C:\Users\Admin\Desktop\data_1\test_bif_file.xlsx"
    filenames = [r"C:\Projects\auto_exc\tets_file.xlsx"]

    dfs = []
    row_counter = 43
    counter = 2
    order = []
    # filenames = find_xlsx_files(directory=r"C:\Users\Admin\Desktop\data_1")
    for path in filenames:
        # Счет
        df_2 =  pd.read_excel(path, skiprows=30, usecols='C:W')   
        df_2.iloc[30:32]    # подумать
        for i in range(2, 22):
            order.append(int(df_2[f'Unnamed: {i}'][0]))

        # Основная таблица
        df_1 = pd.read_excel(path, skiprows=42)
        cnt = 0
        while True:
            if pd.isna(df_1['Unnamed: 0'][cnt]):
                break
            cnt += 1
            df_1.loc[cnt, 'Unnamed: 15'] = order
        df_1.iloc[42:cnt]
        
        dfs.append(df_1)
        dfs.append(df_2)


    print(''.join(map(str, order)))

    combined_df = pd.DataFrame()
    combined_df = pd.concat(dfs, ignore_index=True)


    combined_df.to_excel('report_big_file.xlsx')


if __name__ == '__main__':
    # print(find_last_row('report_big_file.xlsx'))
    
    main()