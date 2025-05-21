import pandas as pd
df_analiz=pd.read_excel('Аналитика_общая_17_05_2025.xlsx')
df = pd.read_excel('name_of_discount.xlsx')
dictonary=df[['Код акции' , 'РА' ,'Код товара' , 'Invent Table 2 → Item Name']]
dictonary.rename(columns={'Код акции':'AdvertActExternalCode','РА' : 'Название акции' ,
                          'Код товара' :'goodsCode', 
                          'Invent Table 2 → Item Name' :'Наименование товара'}, inplace=True )
# Шаг 2: Объединяем с основным датафреймом
keys_of_merge=['goodsCode','AdvertActExternalCode' ]
df_analiz = df_analiz.merge(dictonary, left_on=keys_of_merge, right_on=keys_of_merge, how='left')
# Шаг 3: Удаляем старые столбцы (если нужно)
df_analiz.drop(columns=['goodsCode', 'AdvertActExternalCode'], inplace=True)
# Шаг 4: Переименовываем для удобства
df_analiz.rename(columns={
    'Invent Table 2 → Item Name': 'Наименование товара'
 }, inplace=True)

print(df_analiz.head(10))