import pandas
from collections import defaultdict
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
from datetime import datetime
from pathlib import Path


def load_wines_from_excel(file_path="wine3.xlsx", sheet_name="Лист1"):
    file_path = Path(file_path)
    exel_data_df = pandas.read_excel(
        io=file_path, sheet_name=sheet_name, na_values="nan", keep_default_na=False)
    return exel_data_df.to_dict(orient='records')


def group_wines_by_category(wines):
    grouped_wines = defaultdict(list)
    for wine in wines:
        category = wine['Категория']
        grouped_wines[category].append(wine)
    return dict(grouped_wines)


def sort_categories(grouped_wines):
    return dict(sorted(grouped_wines.items()))


def years_word(company_age):
    if company_age % 100 in (11, 12, 13, 14):
        return 'лет'
    elif company_age % 10 == 1:
        return 'год'
    elif company_age % 10 in (2, 3, 4):
        return 'года'
    else:
        return 'лет'


def main():
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')
    wines = load_wines_from_excel("wines3.xlsx")
    grouped_wines = group_wines_by_category(wines)
    sorted_categories = sort_categories(grouped_wines)
    current_year = datetime.now().year
    foundation_year = 1920
    company_age = current_year - foundation_year

    years = years_word(company_age)

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
