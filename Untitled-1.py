import pandas as pd
import xml.etree.ElementTree as ET

# --- 1. Парсим продажи ---
tree = ET.parse('b02_purchases_2025_04_29_return_value.xml')
root = tree.getroot()

sales_rows = []
# Список ключей для чтения из purchase
keys = ['operDay', 'shop', 'cash', 'shift', 'number', 'amount', 'discountAmount', 'fiscalDocNum']
# Сопоставление для переименования
rename_map = {'amount': 'amount_itogo'}

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
final_df_disc.rename(columns={'saletime': 'operDay', 'goodCode': 'goodsCode'}, inplace=True)
if 'amount' in final_df_disc.columns:
    final_df_disc.drop(columns=['amount'], inplace=True)
# --- 3. Преобразуем типы ---
for df in [final_df, final_df_disc]:
    # Дата
    df['operDay'] = df['operDay'].str[:10]  # Обрезаем до первых 10 символов
    df['operDay'] = pd.to_datetime(df['operDay'], errors='coerce').dt.strftime('%d-%m-%Y')
    # Числовые поля
    for col in ['shop', 'cash', 'shift', 'number', 'goodsCode', 'quantity', 'amount', 'amount_itogo']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').astype('float')

# --- 4. Объединяем по ключам ---
key_fields = ['operDay', 'shop', 'cash', 'shift', 'number', 'goodsCode']
merged = final_df.merge(
    final_df_disc,
    how='left',
    on=key_fields)
total_amount_per_doc = merged.groupby('fiscalDocNum')['amount_itogo'].sum()
print(total_amount_per_doc)

breakpoint()
# # 1. Дополнительная агрегация по ['shop', 'operDay']


# # 2. Основная агрегация по ['AdvertActExternalCode', 'shop', 'operDay', 'goodsCode']
# agg_main = merged.groupby(['AdvertActExternalCode', 'shop', 'cash', 'shift', 'number', 'operDay', 'goodsCode']).agg(
#     Chekov_s_tovarom_vsego=('fiscalDocNum', 'size'),
#     Vsego_tovarov_shtuk=('quantity', 'sum'),
#     TO_po_tovaru=('amount', 'sum')).reset_index()

# # Для расчёта долей в дальнейшем!!!!
# agg_extra_TO = merged.groupby(['shop', 'cash', 'shift', 'number','operDay']).agg(
#     vsego_chekov=('fiscalDocNum', 'size'), TO_itogo_shop=('amount_itogo', 'sum')
# ).reset_index()
# # считаем ТО по товару
# agg_extra_disc=merged.groupby(['AdvertActExternalCode', 'shop','cash', 'shift', 'number', 'operDay']).agg(summa_TO_akcii=('amount', 'sum'), chekov_akcii=('fiscalDocNum', 'size')).reset_index()

# df = agg_main.merge(agg_extra_TO, how='left', on=['shop','cash', 'shift', 'number', 'operDay'])
# df = df.merge(agg_extra_disc, how='left', on=['shop','cash', 'shift', 'number', 'operDay'])
# df['share_akcii'] = df['chekov_akcii'] / df['Chekov_s_tovarom_vsego'] 


#  --- 5. Сохраняем результат ---
merged.to_excel('finality.xlsx', index=False)
finally_df.to_excel('agg.xlsx', index=False)
print('Итоговый файл: finality.xlsx')
