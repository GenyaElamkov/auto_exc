import pandas as pd



filename = r'C:\Users\Admin\Desktop\data_1\test_bif_file.xlsx'



df = pd.read_excel(filename)

print(df.head())