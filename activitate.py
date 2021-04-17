import openpyxl, requests
from lxml import html
import time

wb = openpyxl.load_workbook('activitate.xlsx')
sheet = wb['Activitate']
i=2
for row in sheet.iter_cols(min_row=2, min_col=3, max_col= 3):
    for cell in row:
        sheet['E{}'.format(i)].value = sheet['F{}'.format(i)].value
        link = cell.value
        page = requests.get(link)
        tree = html.fromstring(page.content)
        checkm = tree.xpath('/html/body/div/div[7]/div[1]/div[2]/div[5]/div/text()')
        if checkm[0].strip() == 'CURRENT STATS':
            minute = tree.xpath('/html/body/div/div[7]/div[1]/div[2]/div[6]/text()')
            sheet['F{}'.format(i)].value = int(minute[4].strip())
            print('ONLINE - ', sheet['B{}'.format(i)].value, ' are:', sheet['F{}'.format(i)].value, ' ', sheet['C{}'.format(i)].value)
        else:
            minute = tree.xpath('/html/body/div/div[7]/div[1]/div[2]/div[5]/text()')
            sheet['F{}'.format(i)].value = int(minute[4].strip())
            print('OFFLINE - ', sheet['B{}'.format(i)].value, ' are:', sheet['F{}'.format(i)].value, ' ', sheet['C{}'.format(i)].value)

        i=i+1

wb.save('activitate.xlsx')
input()
