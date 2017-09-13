#-*-coding:utf8-*-

import xlrd

if __name__ == '__main__':
    f1 = open('drug_imfom','w')
    data1 = xlrd.open_workbook('S1_drug target source.xlsx')
    table1 = data1.sheet_by_name("Table-S5E LOBICO models")
    data2 = xlrd.open_workbook('Model.xlsx')
    table2 = data2.sheet_by_name("best_Model")

    dict_target = {}
    dict_class = {}
    count_drug = 0
    count_num = 0
    nrows1 = table1.nrows
    for i in range(31,nrows1):
        drug_name = table1.row(i)[4].value
        if drug_name == 681640:
            drug_name = "681640"
        if drug_name not in dict_target:
            dict_target[drug_name] = table1.row(i)[5].value
            dict_class[drug_name] = table1.row(i)[6].value
            count_drug = count_drug + 1

    nrows2 = table2.nrows
    for i in range(5,nrows2):
        drug_name = str(table2.row(i)[2].value)
        #对drug_name进行格式转化，使之能够成功匹配
        if drug_name.find("_")!=-1:
            drug_name = drug_name.replace("_"," ")
        if drug_name == " 5Z -7-Oxozeaenol":
            drug_name = "(5Z)-7-Oxozeaenol"
        if drug_name == "I-BET 151":
            drug_name = "I-BET-762"
        if drug_name == "PXD101 Belinostat":
            drug_name = "PXD101, Belinostat"
        if drug_name == "VNLG124":
            drug_name = "VNLG/124"
        if drug_name == "Zibotentan ZD4054":
            drug_name = "Zibotentan, ZD4054"
        if drug_name == "681640.0":
            drug_name = "681640"
        try:
            drug_target = dict_target[drug_name]
            drug_class = dict_class[drug_name]
            count_num = count_num + 1
        except:
            f1.write("Drug not found!"+"\n")
            print(drug_name)
        else:
            f1.write(drug_name+"@"+drug_target + "@"+ drug_class + "\n")
    #print(dict_target["681640"])
    #print(dict_class)
    print("Count_drug is %d"%count_drug)
    print("Count_num is %d"%count_num)
    f1.close()

