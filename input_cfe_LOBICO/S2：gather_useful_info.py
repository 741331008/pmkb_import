#-*-coding:utf8-*-

import xlrd

if __name__ == '__main__':
    f1 = open('gather_info','w',encoding="utf8")
    data = xlrd.open_workbook('Model.xlsx')
    table = data.sheet_by_index(0)
    nrows = table.nrows
    count_line = 0
    for i in range(5,nrows):
        Cancer_type = table.row(i)[1].value
        Drug_name = table.row(i)[2].value
        if Drug_name == 681640:
            Drug_name = "681640"
        Drug_target = table.row(i)[3].value
        Drug_class = table.row(i)[4].value
        Sen_ccl = str(int(table.row(i)[5].value))
        Res_ccl = str(int(table.row(i)[6].value))
        Model = table.row(i)[7].value
        Precision = str(table.row(i)[9].value)
        Sensitivity = str(table.row(i)[10].value)
        Response = "Sensitive"
        f1.write(Cancer_type+"@"+Drug_name+"@"+Drug_target+"@"+Drug_class+"@"+Sen_ccl+"@"+Res_ccl+"@"+Model+"@"+Precision+"@"+Sensitivity+"@"+Response+"\n")
        count_line = count_line + 1

    print("Count is %d"%count_line)
    f1.close()


