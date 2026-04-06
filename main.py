import pandas
import pprint 
from collections import defaultdict
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape 
from datetime import datetime

env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)

template = env.get_template('template.html')
today = datetime.now()
year_foundation = datetime(year=1920, month=1, day=1)  
age_company = today.year - year_foundation.year  
exel_data_df = pandas.read_excel(io="wine3.xlsx", sheet_name="Лист1", na_values="nan", keep_default_na=False) 
wines = exel_data_df.to_dict( orient='records') 
grouped_wines = defaultdict(list) 
for wine in wines:
    category = wine['Категория']
    grouped_wines[category].append(wine) 

grouped_wines = dict(grouped_wines) 
sorted_categories = sorted(grouped_wines.keys())

def years_form(number):
    if number % 100 in (11, 12, 13, 14):
        return 'лет'
    elif number % 10 == 1:
        return 'год'
    elif number % 10 in (2, 3, 4):
        return 'года'
    else:
        return 'лет' 

years = years_form(age_company) 


rendered_page = template.render(
    cap1_title= age_company,
    cap2_title= years, 
    wines=grouped_wines, 
    sorted_categories=sorted_categories
) 



with open('index.html', 'w', encoding="utf8") as file:
    file.write(rendered_page)  

pprint.pprint(grouped_wines, width=120)

server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler) 
server.serve_forever()
