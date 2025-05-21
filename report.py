import pandas as pd
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
import os
from openpyxl import load_workbook

def generate_pdf_report(df, pivot_promo, pivot_products, charts_path='charts.xlsx', output_file='Отчет_по_акциям.pdf'):
    template_str="""
        <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <title>Аналитика</title>
        <style>
            body { font-family: sans-serif; padding: 20px; }
            h1, h2 { text-align: center; color: #2c3e50; }
            table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
            th, td { border: 1px solid black; padding: 8px; text-align: left; }
            img { max-width: 100%; height: auto; display: block; margin: 20px auto; }
        </style>
    </head>
    <body>

    <h1>Аналитика продаж и акций</h1>
    <p style="text-align:center;">Дата: {{ date }}</p>

    <h2>Свод по акциям</h2>
    {{ promo_table }}

    <h2>Свод по товарам</h2>
    {{ product_table }}

    <h2>Графики</h2>
    {% for chart in charts %}
    <img src="{{ chart }}" />
    {% endfor %}

    </body>
    </html>
    """
    env = Environment(loader=FileSystemLoader('.'))
    template = env.get_template('template.html')

    # Преобразуем DataFrame в HTML
    promo_html = pivot_promo.to_html(index=False, classes='table')
    product_html = pivot_products.to_html(index=False, classes='table')

    # Сохраняем временные изображения графиков
    wb = load_workbook(charts_path)
    chart_sheet = wb['Графики']
    chart_paths = []

    temp_dir = 'temp_charts'
    os.makedirs(temp_dir, exist_ok=True)

    for idx, img in enumerate(chart_sheet._images):
        temp_img_path = os.path.join(temp_dir, f'chart_{idx}.png')
        with open(temp_img_path, 'wb') as f:
            f.write(img._data())
        chart_paths.append(temp_img_path)

    # Рендерим шаблон
    html_out = template.render(
        date=pd.Timestamp.now().strftime("%d-%m-%Y %H:%M"),
        promo_table=promo_html,
        product_table=product_html,
        charts=chart_paths
    )

    # Сохраняем в PDF
    HTML(string=html_out).write_pdf(output_file)

    # Удаляем временные файлы
    for path in chart_paths:
        os.remove(path)
    os.rmdir(temp_dir)

    print(f"✅ Отчёт сохранён как {output_file}")