# === Функция отправки файла в Google Sheets ===
def upload_to_google_sheets(file_path, sheet_name='Аналитика'):
    scope = ["https://spreadsheets.google.com/feeds ", "https://www.googleapis.com/auth/drive "]
    creds = Credentials.from_service_account_file('service_account.json', scopes=scope)
    client = gspread.authorize(creds)

    try:
        sh = client.open(sheet_name)
    except gspread.exceptions.SpreadsheetNotFound:
        sh = client.create(sheet_name)

    # Открываем первый лист и очищаем его
    worksheet = sh.sheet1
    worksheet.clear()

    # Загружаем данные из Excel
    df_all = pd.read_excel(file_path, sheet_name='Все магазины')
    data = [df_all.columns.values.tolist()] + df_all.values.tolist()
    worksheet.update(data)

    # Добавляем листы по магазинам
    xls = pd.ExcelFile(file_path)
    for sheet in xls.sheet_names:
        if sheet != "Все магазины" and sheet != "Графики":
            try:
                ws = sh.add_worksheet(title=sheet, rows="1000", cols="20")
            except gspread.exceptions.APIError:
                ws = sh.worksheet(sheet)
            df = pd.read_excel(file_path, sheet_name=sheet)
            data = [df.columns.values.tolist()] + df.values.tolist()
            ws.update(data)

    print(f"Файл загружен в Google Sheets: {sh.url}")

# # === Основная функция аналитики ===