import json
from datetime import datetime, timedelta
from docxtpl import DocxTemplate
from babel.dates import format_date



context = {}

with open("data.json") as f:
    data = json.load(f)

for k,v in data.items():
    context[k] = v

doc = DocxTemplate("example_waybill.docx")


start_data = datetime.strptime(context["data_start"], '%d.%m.%Y')
stop_data = datetime.strptime(context["data_stop"], '%d.%m.%Y')
print(type((stop_data - start_data).days))
for day in range(int((stop_data - start_data).days)+1):
    current_data = start_data + timedelta(days=day)
    context["data_start_short"] = (start_data+timedelta(days=day)).strftime('%d.%m.%Y')
    context["data_start"] = format_date(start_data+timedelta(days=day), format='long', locale='ru_RU')
    context["data_stop"] = format_date(start_data+timedelta(days=day+1), format='long', locale='ru_RU')
    # подставляем контекст в шаблон
    doc.render(context)
    # сохраняем и смотрим, что получилось
    doc.save(context["data_start_short"] + ".docx")


