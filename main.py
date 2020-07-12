import random
import datetime
from docxtpl import DocxTemplate
from pprint import pprint
import pandas

excel_data = pandas.read_excel('Объекты.xlsx', sheet_name='Лист1', keep_default_na=False, usecols=['Объект']).to_dict(orient='index')
# pprint(excel_data)
buildings = ["Склад гсм", "Стоянка", "Ангар"]
buildings_bypass = {}
date_year = 2020
date_month = 7
date_day = 12
star_hour = 8
stop_hour = 8

for house in excel_data.values():
    minutes = random.randrange(0, 30, 5)
    start_time = datetime.datetime(year=date_year, month=date_month, day=date_day, hour=star_hour, minute=minutes)
    to_time = datetime.datetime(year=date_year, month=date_month, day=date_day + 1, hour=stop_hour)
    time = []
    time.append(start_time)
    while time[-1] + datetime.timedelta(minutes=120) < to_time:
        minutes = random.randrange(90, 120, 5)
        time.append(time[-1] + datetime.timedelta(minutes=minutes))
    buildings_bypass[house['Объект']] = time
# pprint(buildings_bypass)


doc = DocxTemplate("template.docx")
context = {'date_time': "12.07.2020 12:47",
           'from_date': "12.07.2020",
           'to_date': "13.07.2020",
           'buildings': buildings_bypass}
doc.render(context)
doc.save("1207.docx")
