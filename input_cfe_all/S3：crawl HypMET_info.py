#-*-coding:utf8-*-

import xlrd
import re

if __name__=='__main__':
    f1 = open("HypMET_info","w",encoding="utf8")
    data1 = xlrd.open_workbook('TableS3F.xlsx',encoding_override="utf8")
    data2 = xlrd.open_workbook('Hypermethylation_Frequency.xlsx',encoding_override="utf8")
    table1 = data1.sheet_by_index(0)
    table2 = data2.sheet_by_index(1)
    count_HypMET = 0

    nrows = table1.nrows
    nrows2 = table2.nrows
    for row in range(2,nrows):

        count_match = 0
        list_cancer = [];list_threshold=[];
        list_frequency=[];list_RefGene=[];list_Island=[]#用于记录匹配到的信息

        feature = table1.row(row)[2].value
        if feature.find("_HypMET")!=-1:
            list_feature = re.split(r'[-:)(]',feature)
            chr = list_feature[0].strip("chr")
            gene = list_feature[3]
            start_site = list_feature[1]
            stop_site = list_feature[2]
            feature_ids = feature.split("(")[0]#从feature中提取到的信息

            CFE_com = int(table1.row(row)[5].value)
            CFE_back = int(table1.row(row)[6].value)
            All_com = int(table1.row(row)[7].value)
            All_back = int(table1.row(row)[8].value)#必备的四条信息

            for i in range(7,nrows2):
                if table2.row(i)[2].value == feature_ids:
                    count_match = count_match + 1
                    list_cancer.append(table2.row(i)[1].value)
                    list_threshold.append(table2.row(i)[4].value)#float
                    list_frequency.append(table2.row(i)[12].value)#float
                    if table2.row(i)[9].value not in list_RefGene:
                        list_RefGene.append(table2.row(i)[9].value)
                    if table2.row(i)[10].value not in list_Island:
                        list_Island.append(table2.row(i)[10].value)

            f1.write(feature+"@"+str(CFE_com)+"@"+str(CFE_back)+"@"+str(All_com)+"@"+str(All_back)+"@"+chr+"@"+start_site+"@"+\
                     stop_site+"@"+gene+"@"+str(list_RefGene[0])+"@"+str(list_Island[0])+"@"+str(count_match)+"@")
            for i in range(count_match):
                f1.write(list_cancer[i]+"@"+str(list_threshold[i])+"@"+str(list_frequency[i])+"@")
            f1.write("\n")
            if count_match>1:
                print(feature_ids)
                print(count_match)
            count_HypMET = count_HypMET + 1

    print("Count_HypMET is %d"%count_HypMET)
    f1.close()