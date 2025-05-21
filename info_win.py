import zipfile
from lxml import etree
import os
import pandas as pd


df=pd.read_excel('Project_all_info.xlsx', engine='openpyxl', nrows=5000)
# Получаем первый чанк
# first_chunk = next(df)
# Смотрим информацию о нем
print(df.info())
print(df.head(10))
# Сохраняем info первого чанка
with open('df_info.txt', 'w', encoding='utf-8') as f:
    df.info(buf=f, show_counts=True)
df.to_excel('first_5000.xlsx')