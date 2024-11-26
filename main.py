### Этот скрипт копирует данные из Excel в google таблицу и чистит данные.

import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build

SERVICE_ACCOUNT_FILE = "password1.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
)
service = build("sheets", "v4", credentials=credentials)

SPREADSHEET_ID = "1Ji_YBR2EjBS3CgurXPecNegoRHk-EE9lGnURKuY5XEU"

excel_file_path = "22-11-2024_10-30-04/Тест для кандидата аналитика.xlsx"

revenue_df = pd.read_excel(excel_file_path, sheet_name="Выручка")
stock_df = pd.read_excel(excel_file_path, sheet_name="Остатки")
products_df = pd.read_excel(excel_file_path, sheet_name="Товары")


def remove_substrings(df):
    df = df.replace(to_replace=r"ЯЯЯ___|ЯЯЯ_|Склад ", value="", regex=True)
    return df.applymap(lambda x: x.strip() if isinstance(x, str) else x)


revenue_df = remove_substrings(revenue_df)
stock_df = remove_substrings(stock_df)
products_df = remove_substrings(products_df)

stock_df.fillna("", inplace=True)
products_df.fillna("", inplace=True)

data_revenue = [revenue_df.columns.tolist()] + revenue_df.values.tolist()
data_stock = [stock_df.columns.tolist()] + stock_df.values.tolist()
data_products = [products_df.columns.tolist()] + products_df.values.tolist()

print(data_stock)


def write_to_sheet(data, range_name):
    body = {"values": data}
    try:
        result = (
            service.spreadsheets()
            .values()
            .update(
                spreadsheetId=SPREADSHEET_ID,
                range=range_name,
                valueInputOption="RAW",
                body=body,
            )
            .execute()
        )
        print(f"{result.get('updatedCells')} ячеек обновлено в диапазоне {range_name}.")
    except Exception as e:
        print(f"Ошибка при записи в диапазон {range_name}: {e}")


write_to_sheet(data_revenue, "Выручка!A1")
write_to_sheet(data_stock, "Остатки!A1")
write_to_sheet(data_products, "Товары!A1")
