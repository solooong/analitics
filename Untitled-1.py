import pandas as pd
import xml.etree.ElementTree as ET

# --- 1. Парсим продажи ---
tree = ET.parse('b02_purchases_2025_04_29_return_value.xml')
root = tree.getroot()

sales_rows = []
for purchase in root.findall('.//purchase'):
    purchase_data = {key: purchase.attrib.get(key) for key in [
        'tabNumber', 'operDay', 'shop', 'cash', 'shift', 'number', 'amount', 'discountAmount', 'fiscalDocNum'
    ]}
    # Добавим позиции
    for pos in purchase.findall('.//position'):
        pos_data = pos.attrib.copy()
        # Добавим fiscalDocNum и purchase info к позиции
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

# --- 3. Преобразуем типы ---
for df in [final_df, final_df_disc]:
    # Дата
    df['operDay'] = pd.to_datetime(df['operDay'], errors='coerce').dt.strftime('%d:%m:%Y')
    # Числовые поля
    for col in ['shop', 'cash', 'shift', 'number', 'goodsCode']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').astype('Int64')

# --- 4. Объединяем по ключам ---
key_fields = ['operDay', 'shop', 'cash', 'shift', 'number', 'goodsCode']
merged = final_df.merge(
    final_df_disc,
    how='left',
    on=key_fields,
    suffixes=('', '_disc')
)

# --- 5. Сохраняем результат ---
merged.to_excel('finality.xlsx', index=False)
print('Done! Итоговый файл: finality.xlsx')
