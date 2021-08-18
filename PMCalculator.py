from datetime import datetime
import pandas as pd
from datetime import timedelta
tstrt = datetime.today()

DATEROW = 3  # column with date of motor hours from source xls file
MHROW = 13  # column with motor hours from source xls file

INTERVAL_MH = 250  # PM interval in motor hours
INTERVAL_DAYS = 12  # PM interval in days

H_PER_DAY = int(round(INTERVAL_MH/INTERVAL_DAYS, 0))  # motor hours per day
stop_list = ['685_2_2_2', '684']  # list for skip inactive machines
BASE_DATE = datetime(2021, 8, 15)  # date of calculations start

xlsFile = pd.ExcelFile('tab.xls')
s_names = xlsFile.sheet_names

dataset = []
for sheet in s_names:
    if sheet not in stop_list:
        df = xlsFile.parse(sheet, header=None)
        prep = []
        for row in df.itertuples():
            if type(row[DATEROW]) is datetime and str(row[MHROW]) != 'nan':
                prep.append([row[DATEROW], row[MHROW]])
        dataset.append([sheet, prep[len(prep)-1][0], prep[len(prep)-1][1]])

step = 0
for k in dataset:
    target_mh = k[2] // INTERVAL_MH * INTERVAL_MH + INTERVAL_MH
    iter_mh = k[2]
    days_counter = 0
    while iter_mh < (target_mh-H_PER_DAY):
        iter_mh += H_PER_DAY
        days_counter += 1
    dataset[step].append(days_counter+1)
    dataset[step].append(iter_mh + H_PER_DAY)
    dataset[step].append(k[1] + timedelta(days_counter))
    step += 1

dates = []
for row in dataset:
    dates.append(row[1])
    dates.append(row[5])
min_date = min(dates)
max_date = max(dates)
date_delta = max_date-BASE_DATE


output_list = []
header = 'SN машины;ДатаПоказаний'
for q in range(0, date_delta.days+2):
    header = header + ';' + str((BASE_DATE + timedelta(q)).strftime('%d.%m.%Y'))
output_list.append(header)


for i in dataset:
    warn = ''
    if (BASE_DATE - i[1]).days > 0:
        warn = '(!)'
    str_data = str(i[0]) + warn + ';' + i[1].strftime('%d.%m.%Y') + ';' + str(i[2]) + ';'*i[3] + str(i[4])
    output_list.append(str_data)
    print(i)

for k in output_list:
    print(k)

with open(f'График ТО с {BASE_DATE.strftime("%d.%m.%Y")}.csv', 'w', encoding='cp1251') as out_file:
    for i in output_list:
        out_file.write(i+'\n')
print('Excecute time:', datetime.today()-tstrt)
