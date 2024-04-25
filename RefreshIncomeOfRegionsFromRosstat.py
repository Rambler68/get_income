import pandas as pd
import openpyxl as pxl

# ### Предобработка
df = pd.read_excel("https://rosstat.gov.ru/storage/mediabank/urov_10subg-nm.xlsx")

# Убираем строки, в которых есть данные по округам
df = df[~df.iloc[:, 0].str.contains('округ', na=False)]

# Удаление первых 2 и последних 2 строк
df = df.iloc[2:-2]

df.iloc[0,0] = 'Год'
df.iloc[1,0] = 'Квартал'

# Транспонирование таблицы и заполнение значений вниз
df_trans = df.transpose().fillna(method='ffill')

df_trans


# ### Обработка столбца, содержащего кварталы

# Фильтрация строк
df = df_trans[~(df_trans[3].str.endswith('год') | df_trans[3].isnull())]

# Переименовываем индексы (заголовки столбцов)
df.columns = df.iloc[0]
df.drop(df.index[0])
df = df.iloc[1:]

# Замена значений в столбце с кварталами
df['Квартал'] = df['Квартал'].str.replace('IV', '4')
df['Квартал'] = df['Квартал'].str.replace('III', '3')
df['Квартал'] = df['Квартал'].str.replace('II', '2')
df['Квартал'] = df['Квартал'].str.replace('I', '1')


# Убираем слово "квартал" 
df['Квартал'] = df['Квартал'].str.replace('квартал', '', case=True)

# Переводим столбец "Квартал" в числовой формат
df['Квартал'] = pd.to_numeric(df['Квартал'], errors='coerce')


# Замена незначащих знаков (звёздочки)
df['Год'] = df['Год'].str.replace('*', "")


df['Год'] = df['Год'].str.replace('год', '', case=True)


df['Год'] = pd.to_numeric(df['Год'], errors='coerce')


# Фильтрация строк по условию
df = df[(df['Год'].isnull()) | (df['Год'] > 2018)]


df['Год'].fillna(0)


df_melt = df.melt(id_vars=['Год','Квартал'], var_name='Region', value_name='Value')



df_melt.to_excel(r'region_income.xlsx', index=False)



df_melt.to_excel(r'\\dengi-srv\report\Выгрузки\ПДН\region_income.xlsx', index=False)