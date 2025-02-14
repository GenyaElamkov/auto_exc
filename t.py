
import pandas as pd

# Создание DataFrame
df = pd.DataFrame({'column_name': [1, 2, 3, 4, 5]})

# Добавление значения в столбец
value = 10
row_index = 2

df.loc[row_index, 'column_name'] = value

print(df)