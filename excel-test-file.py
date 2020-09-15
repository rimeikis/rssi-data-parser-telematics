import random
import time
from openpyxl import Workbook
from datetime import datetime
from openpyxl.styles import Alignment

wb = Workbook()


def formattedTime():
    now = datetime.now()
    dateTimeFormatted = now.strftime("%Y-%m-%d %H_%M_%S")
    return dateTimeFormatted


wb_filename = 'parser_results_' + formattedTime() + '.xlsx'

ws1 = wb.active
ws1.title = 'test'

rssiList = []
mac = input('>> Enter MAC: ')
notes = input('>> Enter test notes: ')

ws1['A1'] = 'RSSI PARSE RESULTS'
ws1['A3'] = 'Notes'
ws1['A6'] = 'Target MAC'
ws1['B6'] = 'Start time'
ws1['C6'] = 'Duration'
ws1['D6'] = 'Data count'
ws1['E6'] = 'Average'
ws1['F6'] = 'Max'
ws1['F6'].alignment = Alignment(horizontal='center')
ws1['G6'] = 'Min'
ws1['H6'] = 'Data'


def function():
    startTime = datetime.now()
    for i in range(7, 13):
        randomRssi = random.randint(-60, -40)
        time.sleep(1)
        rssiList.append(randomRssi)
        ws1['H'+str(i)] = randomRssi
        print(randomRssi)
    ws1['A4'] = notes
    ws1['A7'] = mac
    ws1['B7'] = startTime
    ws1['C7'] = datetime.now() - startTime
    ws1['D7'] = "=COUNT(H7:H1048576)"
    ws1['E7'] = "=AVERAGE(H7:H1048576)"
    ws1['F7'] = "=MAX(H7:H1048576)"
    ws1['G7'] = "=MIN(H7:H1048576)"
    print(rssiList)
    wb.save(filename=wb_filename)


function()
