#-*-coding:utf8-*-

import xlrd

f1 = open('inter_CFE','w')
data = xlrd.open_workbook('TableS4C.xlsx')
table = data.sheet_by_index(0)

start_row = 0
count = 0
count_sen = 0
count_res = 0
nrows = table.nrows
while table.cell(start_row,0).value != "Domain":
    start_row = start_row + 1

start_row = start_row + 1
for rownum in range(start_row,nrows):
    row = table.row_values(rownum)
    count = count + 1
    if row:
        Cancer_type = row[0]
        Feature = row[4]
        if Feature.find(" ()")!=-1:
            Feature = Feature.strip(" ()")
        Drug_name = row[6]
        Drug_target = row[7]
        if row[13]>0:
            Response = "Resistant"
            count_res = count_res + 1
        else:
            Response = "Sensitive"
            count_sen = count_sen + 1
        f1.write(str(Cancer_type)+'@'+str(Feature)+'@'+str(Drug_name)+'@'+str(Drug_target)+'@'+Response+'\n')



print("Count_num is %d"%count)
print("Count_res is %d"%count_res)
print("Count_sen is %d"%count_sen)
f1.close()