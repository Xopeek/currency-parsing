import requests
import openpyxl
from xml.etree import ElementTree
from datetime import datetime

url = "https://www.cbr.ru/scripts/XML_daily.asp"


def get_exchange_rate(currency_code):
    response = requests.get(url)
    tree = ElementTree.fromstring(response.content)
    for valute in tree.findall("Valute"):
        if valute.find("CharCode").text == currency_code:
            return float(valute.find("Value").text.replace(",", "."))
    return None


def create_excel_file(usd, eur):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")

    file_name = "exchange_rates.xlsx"
    try:
        wb = openpyxl.load_workbook(file_name)
        sheet = wb.active
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.append(["USD", "EUR", "Время обновления"])

    sheet.append([usd, eur, timestamp])
    wb.save(file_name)


if __name__ == '__main__':
    usd = get_exchange_rate("USD")
    eur = get_exchange_rate("EUR")
    create_excel_file(usd, eur)
