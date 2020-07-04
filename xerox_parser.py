#!/usr/bin/python3

from datetime import date, datetime
from lxml import html
import requests
import xlsxwriter
from requests import exceptions as requests_exceptions



var = input("Ведите диапазон через пробел, например: 1 200000: \n")
_range =  str(var).split()
if len(_range) != 2:
    print('вы ввели что-то не то')
else:
    print('Скрипт выполняется...')

start = int(_range[0])
end = int(_range[1])
res = []

for item in range(start, end):
    print('Обрабатывается ID {}...из {}'.format(item, end))
    url = 'https://www.support.xerox.com/support/_all-products/file-download/enus.html?contentId={}'.format(item)

    try:
      response = requests.get(url, allow_redirects=False)
    except requests_exceptions.RequestException as exc:
        msg = "Connection with url {} refused.".format(response.url)
        continue
    if response.status_code == requests.codes.FOUND:
        continue

    html_string = response.text
    tree = html.fromstring(html_string)
    header_list = tree.xpath('//ul[@class ="fileInfo"]/li/strong/text()')
    data_list = tree.xpath('//ul[@class ="fileInfo"]/li/text()')

    if not header_list:
        continue

    header_list_no_whitespaces = []
    for h in header_list:
        header_list_no_whitespaces.append(h.replace(' ', ''))

    data_iter = dict(zip(header_list_no_whitespaces, data_list))
    if data_iter.get('Date'):
        data_iter['Date'] = datetime.strptime(data_iter['Date'], '%b %d, %Y').strftime('%m/%d/%Y')

    file_description = tree.xpath('//div[@class="mainBody fileDownload"]/h2[@class ="record_title"]/text()')
    file_description = file_description[0].split('File Download: ')[1] if file_description else ''

    current = [
        item,
        data_iter.get('Date', ''),
        data_iter.get('Filename', ''),
        data_iter.get('Version', ''),
        file_description
    ]
    res.append(current)

workbook = xlsxwriter.Workbook(
    'xerox_{}({}-{}).xlsx'.format(date.today(), start, end))
worksheet = workbook.add_worksheet()

headers = ['#', 'date', 'fn', 'ver', 'text']
row = 0
col = 0
for header in headers:
    worksheet.write(row, col, header)
    col = col + 1

for chunk in res:
    col = 0
    row = row + 1
    for item in chunk:
        worksheet.write(row, col, str(item))
        col = col + 1

workbook.close()
