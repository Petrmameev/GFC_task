### Этот скрипт копирует данные из Excel в google таблицу, чистит данные и объединяет три таблицы

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

revenue_df.fillna("", inplace=True)
stock_df.fillna("", inplace=True)
products_df.fillna("", inplace=True)

stock_df["Конечный остаток"] = pd.to_numeric(
    stock_df["Конечный остаток"].replace("", pd.NA), errors="coerce"
).fillna(0)
revenue_df["Выручка"] = pd.to_numeric(
    revenue_df["Выручка"].replace("", pd.NA), errors="coerce"
).fillna(1)

merged_df = pd.merge(
    revenue_df,
    stock_df,
    left_on=["Номенклатура", "Подразделение"],
    right_on=["Номенклатура", "Склад"],
    how="outer",
)

merged_df["Склад"] = merged_df["Склад"].replace("", pd.NA)
merged_df["Подразделение"] = merged_df["Подразделение"].replace("", pd.NA)

merged_df["Склад"] = merged_df["Склад"].fillna(merged_df["Подразделение"])
merged_df["Подразделение"] = merged_df["Подразделение"].fillna(merged_df["Склад"])

final_merged_df = pd.merge(
    merged_df, products_df, left_on="Номенклатура", right_on="Наименование", how="left"
)

final_merged_df.fillna("", inplace=True)

final_columns = ["Подразделение", "Номенклатура", "Выручка", "Конечный остаток"]
final_columns += [col for col in products_df.columns if col != "Наименование"]

final_df = final_merged_df[final_columns]


def write_to_sheet(dataframe, sheet_name):
    body = {"values": [dataframe.columns.tolist()] + dataframe.values.tolist()}
    try:
        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{sheet_name}!A1",
            valueInputOption="RAW",
            body=body,
        ).execute()
        print(f"Данные успешно записаны в диапазон {sheet_name}!A1.")
    except Exception as e:
        print(f"Ошибка при записи в диапазон {sheet_name}: {e}")


def create_new_sheet(sheet_title):
    spreadsheet = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    sheets = spreadsheet.get("sheets", [])

    existing_titles = [sheet["properties"]["title"] for sheet in sheets]
    if sheet_title in existing_titles:
        print(f"Лист с названием '{sheet_title}' уже существует.")
        return

    requests = [
        {
            "addSheet": {
                "properties": {
                    "title": sheet_title,
                    "gridProperties": {
                        "rowCount": 100,
                        "columnCount": 10,
                    },
                }
            }
        }
    ]
    body = {"requests": requests}
    try:
        service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID, body=body
        ).execute()
        print(f"Лист '{sheet_title}' успешно создан.")
    except Exception as e:
        print(f"Ошибка при создании листа: {e}")


new_sheet_name = "Все Данные"
create_new_sheet(new_sheet_name)
write_to_sheet(final_df, new_sheet_name)

print("Данные объединены и записаны в новую вкладку.")
