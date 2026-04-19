import os
import argparse
import pandas
from collections import defaultdict
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
from datetime import datetime
from pathlib import Path


def load_wines_from_excel(file_path, sheet_name="Лист1"):
    file_path = Path(file_path)
    excel_data_df = pandas.read_excel(
        io=file_path, sheet_name=sheet_name, na_values="nan", keep_default_na=False)
    return excel_data_df.to_dict(orient='records')


def group_wines_by_category(wines):
    grouped_wines = defaultdict(list)
    for wine in wines:
        category = wine['Категория']
        grouped_wines[category].append(wine) 

    grouped_wines = dict(grouped_wines) 
    sorted_categories = dict(sorted(grouped_wines.items()))    
    return grouped_wines, sorted_categories



def get_year(company_age):
    if company_age % 100 in (11, 12, 13, 14):
        return 'лет'
    elif company_age % 10 == 1:
        return 'год'
    elif company_age % 10 in (2, 3, 4):
        return 'года'
    else:
        return 'лет'

def parse_arguments():
    parser = argparse.ArgumentParser(
        description="Генерация сайта о винах из данных Excel"
    )
    parser.add_argument(
        "--file-path",
        type=Path,
        default=os.getenv("WINE_DATA_PATH", "wine3.xlsx"),
        help="Путь к файлу Excel с данными о винах (по умолчанию: wine3.xlsx или переменная окружения WINE_DATA_PATH)",
    )
    parser.add_argument(
        "--sheet-name",
        default=os.getenv("WINE_SHEET_NAME", "Лист1"),
        help="Название листа в файле Excel (по умолчанию: Лист1 или переменная окружения WINE_SHEET_NAME)",
    )
    return parser.parse_args()


def main():
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    
    template = env.get_template('template.html')
    args = parse_arguments()
    wines = load_wines_from_excel(args.file_path, args.sheet_name)
    grouped_wines, sorted_categories = group_wines_by_category(wines)
    current_year = datetime.now().year
    foundation_year = 1920
    company_age = current_year - foundation_year

    years = get_year(company_age) 
    

    rendered_page = template.render(
        age_company=company_age,
        years=years,
        categories=sorted_categories
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == "__main__":
    main()
