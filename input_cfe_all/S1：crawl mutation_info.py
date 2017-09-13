#-*-coding:utf8-*-

import xlrd

if __name__=='__main__':
    f1 = open("mutation_info","w")
    data = xlrd.open_workbook('TableS3F.xlsx')
    table = data.sheet_by_index(0)
    count_mutation = 0

    nrows = table.nrows
    for row in range(2,nrows):
        feature = table.row(row)[2].value
        if feature.find("_mut")!=-1:
            CFE_com = int(table.row(row)[5].value)
            CFE_back = int(table.row(row)[6].value)
            All_com = int(table.row(row)[7].value)
            All_back = int(table.row(row)[8].value)
            f1.write(feature+"@"+str(CFE_com)+"@"+str(CFE_back)+"@"+str(All_com)+"@"+str(All_back)+"\n")
            count_mutation = count_mutation+1
    print("Count_mutations is %d"%count_mutation)
    f1.close()
