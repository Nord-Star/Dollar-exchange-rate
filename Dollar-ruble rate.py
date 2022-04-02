import requests
from bs4 import BeautifulSoup as bs
import xlsxwriter as xl
import matplotlib.pyplot as plt
import numpy as np

# dollar-ruble exchange rate from Yandex
url = 'https://yandex.ru/news/quotes/1.html'

# Parsing the Internet page
source = requests.get(url)
main_text = source.text
soup = bs(main_text, 'lxml')
table = soup.findAll('div', {'class' : 'news-stock-table__cell'})

# 'Cleaning' the data
ten_days_list = []
for i in table:
    ten_days_list.append(str(i.text))

date = (ten_days_list[0::3])
rate_raw = (ten_days_list[1::3])
rate = []
for i in rate_raw[1:]:
    if "," in i:
        i = i.replace(",", ".")  # changed commas to dots
    i = float(i)
    rate.append(i)
date = date[1:]  # removed the word ДАТА from the received list

# Making a graph in Matplotlib
plt.figure(figsize=(10, 6))
plt.style.use('seaborn-whitegrid')
plt.title('Курс доллара', fontsize=20, fontname='Times New Roman')
x = date[::-1]
y = rate[::-1]
plt.xticks(rotation=45)
plt.plot(x, y, marker="o", markerfacecolor='r')
for i in range(len(y)):
    plt.annotate(round(y[i], 2), (x[i], y[i]))
plt.show()

# Making a table and a graph in Excel
titles = ['Дата', 'Курс']
workbook = xl.Workbook('Курс доллара.xlsx')
worksheet = workbook.add_worksheet()
chart = workbook.add_chart({'type' : 'line'})

worksheet.write_row('B2', date)
worksheet.write_row('B3', rate)
bold = workbook.add_format({'bold': 1})
worksheet.write_column('A2', titles, bold)  # titles

chart.add_series({
    'categories': '=Sheet1!$B2:$K$2',
    'values': '=Sheet1!$B$3:$K$3',
    'marker': {'type': 'automatic'},  # markers on the graph
    'data_labels': {'value': True, 'position': 'above'}  # position of numbers on the graph
    })
chart.set_title ({'name': 'Курс доллара'})
chart.set_x_axis({'name': 'Дата'})
chart.set_y_axis({'name': 'Курс'})
chart.set_size({'width': 720, 'height': 576})  # size of the graph
chart.set_legend({'none': True})  # removed legend
chart.set_x_axis({'reverse': True})  # mirrored the graph left-right

# lines dropping to the axis Х
chart.set_drop_lines({'line': {'color': 'red', 'dash_type': 'square_dot'}})

# min and max value on axis Y
# chart.set_y_axis({'min': 80, 'max': 130})

worksheet.insert_chart('A5', chart, {'x_offset': 25, 'y_offset': 10})
workbook.close()
