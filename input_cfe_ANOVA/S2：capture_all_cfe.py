#-*-coding:utf8-*-

import xlrd

f1 = open('all_CFE','w')

data = xlrd.open_workbook('TableS3F.xlsx')
table = data.sheet_by_index(0)

nrows = table.nrows
start_row = 0
count = 0
while table.row(start_row)[2].value != "Cancer Functional Events":
    start_row = start_row + 1
start_row = start_row + 1

for rownum in range(start_row,nrows):
    Feature = table.row(rownum)[2].value
    if Feature != "Selected Functional Events":
        f1.write(Feature+"\n")
        count = count + 1

print("Count is %d"%count)

