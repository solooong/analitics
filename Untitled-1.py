import pandas as pd
import xml.etree.ElementTree as ET

def parser_xml():
    # --- 1. Парсим продажи ---
    tree = ET.parse('b02_purchases_2025_04_29_return_value.xml')
    root = tree.getroot()
    sales_rows = []
    # Список ключей для чтения из purchase
    keys = ['operDay', 'shop', 'cash', 'shift', 'number', 'amount', 'discountAmount', 'fiscalDocNum', 'order', 'AdvertActExternalCode']
    # Сопоставление для переименования
    rename_map = {'amount': 'amount_itogo'}
    print('Обработка общих продаж')
    for purchase in root.findall('.//purchase'):
        purchase_data = {}
        for key in keys:
            value = purchase.attrib.get(key)
            new_key = rename_map.get(key, key)
            purchase_data[new_key] = value 
        for pos in purchase.findall('.//position'):
            pos_data = pos.attrib.copy()
            row = purchase_data.copy()
            row.update(pos_data)
            sales_rows.append(row)
    final_df = pd.DataFrame(sales_rows)
    # --- 2. Парсим скидки ---
    print('Обработка продаж включающие скидки на товары')

    tree2 = ET.parse('b02_loy_2025_04_29_return_value.xml')
    root2 = tree2.getroot()
    discount_rows = []
    for purchase in root2.findall('.//purchase'):
        purchase_data = {key: purchase.attrib.get(key) for key in [
            'shop', 'cash', 'shift', 'number', 'saletime'
        ]}
        for disc in purchase.findall('.//discount'):
            disc_data = disc.attrib.copy()
            row = purchase_data.copy()
            row.update(disc_data)
            discount_rows.append(row)
    final_df_disc = pd.DataFrame(discount_rows)

 
    final_df_disc.rename(columns={'saletime': 'operDay', 'goodCode': 'goodsCode', 'positionId':'order'}, inplace=True)
    if 'amount' in final_df_disc.columns:
        final_df_disc.drop(columns=['amount'], inplace=True)
    # --- 3. Преобразуем типы ---
    print('Корректируем данные в ячейках')

    for df in [final_df, final_df_disc]:
        # Дата
        df['operDay'] = df['operDay'].str[:10]  # Обрезаем до первых 10 символов
        df['operDay'] = pd.to_datetime(df['operDay'], errors='coerce').dt.strftime('%d-%m-%Y')
        # Числовые поля
        for col in ['shop', 'cash', 'shift', 'number', 'goodsCode', 'quantity', 'amount', 'amount_itogo', 'positionId'  , 'order'
                    ,'discountAmount','count','cost','nds','ndsSum','discountValue','costWithDiscount']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').astype('float')
    print('Объеденяем данные в один DataFrame')

    # --- 4. Объединяем по ключам ---
    key_fields = ['operDay', 'shop', 'cash', 'shift', 'number', 'goodsCode', 'order']
    merged = final_df.merge(
        final_df_disc,
        how='left',
        on=key_fields)
    # Сохраняем данные
    merged.to_excel('finality.xlsx', index=False)
    print('Итоговый файл: finality.xlsx')
    
    return merged

# Чистим базу от лишних данных. Переименовываем столбцы для корректного отображения
def clean_df(clean):
    print('Чистим базу от лишних данных. ')
    key_collum_for_drop = ['shift','number','order','departNumber','barCode','nds','ndsSum','dateCommit','insertType', 'AdvertActGUID', 'card-number', 'advertType', 'quantity', 'barCode']
    
    print('Переименовываем столбцы для корректного отображения')
    dictonary={'operDay':'Дата','shop':'Магазин','cash':'Касса',
               'shift':'Смена','number':'Номер чека','amount_itogo':'Сумма чека',
               'discountAmount':'Сумма скидки чека','fiscalDocNum':'Номер документа',
               'positionId':'Позция в чеке','goodsCode':'ID Товара:','barCode':'Штрих код',
               'count':'Количество товара','cost':'Цена товара','nds':'НДС','ndsSum':'Сумма НДС',
               'discountValue':'Размер скидки на товар','costWithDiscount':
               'Цена товара со скидкой','amount':'Сумма на товар со скидкой', 'Номер журнала МА': 'AdvertActExternalCode'}
    clean_df=pd.DataFrame(clean)
    clean_df.groupby(['operDay', 'shop', 'goodsCode'])
    key_collum_for_drop = [col for col in key_collum_for_drop if col in clean_df.columns]
    clean_df.drop(columns=key_collum_for_drop, inplace=True)
    # Rename columns
    clean_df.rename(columns=dictonary, inplace=True)
    print('Меняем местами столобцы для большего удобства')
    clean_df.to_excel('format_finality.xlsx', index=False)
    print('Итоговый файл: format_finality.xlsx')   
    return clean_df

def analitics_colums(df):
    df = pd.read_excel('test_for_analiz.xlsx')
    # Фильтруем только строки со скидкой
    df_discounted = df[df['isDiscountPurchase']==True].copy()
    # Аналитика по каждому товару в акции
    sales_of_item = (
        df_discounted.groupby(['shop', 'operDay', 'AdvertActExternalCode', 'fiscalDocNum', 'goodsCode'])
        .agg(
            sum_of_item=('amount', 'sum'),
            vsego_chevok_item=('fiscalDocNum', 'count')
        )
        .reset_index()
    )
    # Аналитика по чекам с акцией (без разделения по товарам)
    sales_of_discount = (
        df_discounted.groupby(['shop', 'operDay', 'AdvertActExternalCode', 'fiscalDocNum'])
        .agg(
            sum_of_discount=('amount', 'sum'),
            vsego_chekov_discount=('fiscalDocNum', 'count')
        )
        .reset_index()
    )
    # Общая аналитика по магазинам и датам (по чекам)
    sales_of_shop = (
        df.drop_duplicates(subset='fiscalDocNum', keep='first')
        .groupby(['shop', 'operDay'])
        .agg(
            sum_of_sale_shop=('amount_itogo', 'sum'),
            vsego_chekov_shop=('fiscalDocNum', 'count')
        )
        .reset_index()
    )
    # Merge: объединяем всё обратно в исходный DataFrame
    keys_item = ['shop', 'operDay', 'AdvertActExternalCode', 'fiscalDocNum', 'goodsCode']
    keys_discount = ['shop', 'operDay', 'AdvertActExternalCode', 'fiscalDocNum']
    df = df.merge(sales_of_item, how='left', on=keys_item)
    df = df.merge(sales_of_discount, how='left', on=keys_discount)
    df = df.merge(sales_of_shop, how='left', on=['shop', 'operDay'])
    return df
    # Указать обязательно информацию о количестве акций в данный период и количестве товара в акции


# def math_of_colum():#  Тут расчитываем показатели вхождения 

def main():
   
    original_df=parser_xml()
    analiz_df=analitics_colums(original_df)
    df_cleaned_final = clean_df(analiz_df)
    

analitics_colums()

# # Для расчёта долей в дальнейшем!!!!
# agg_extra_TO =111111111110 merged.groupby(['shop', 'cash', 'shift', 'number','operDay']).agg(
#     vsego_chekov=('fiscalDocNum', 'size'), TO_itogo_shop=('amount_itogo', 'sum')
# ).reset_index()
# # считаем ТО по товару
# agg_extra_disc=merged.groupby(['AdvertActExternalCode', 'shop','cash', 'shift', 'number', 'operDay']).agg(summa_TO_akcii=('amount', 'sum'), chekov_akcii=('fiscalDocNum', 'size')).reset_index()

# df = agg_main.merge(agg_extra_TO, how='left', on=['shop','cash', 'shift', 'number', 'operDay'])
# df = df.merge(agg_extra_disc, how='left', on=['shop','cash', 'shift', 'number', 'operDay'])
# df['share_akcii'] = df['chekov_akcii'] / df['Chekov_s_tovarom_vsego'] 



