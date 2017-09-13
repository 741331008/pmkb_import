#-*-coding:utf8-*-

import xlrd
import re

def feature_id(feature):
    new_feature = feature
    if feature.find("(")!=-1:
        new_feature = feature[:feature.find("(")]
    return re.sub("\D", "", new_feature)

if __name__=='__main__':
    f1 = open("RACS_info","w",encoding="utf8")
    data1 = xlrd.open_workbook('TableS3F.xlsx',encoding_override="utf8")
    data2 = xlrd.open_workbook('mmc3.xlsx',encoding_override="utf8")
    table1 = data1.sheet_by_index(0)
    table2 = data2.sheet_by_index(3)
    count_RACS = 0

    nrows = table1.nrows
    for row in range(2,nrows):
        #print(row)
        feature = table1.row(row)[2].value
        if feature.find("PANCAN")!=-1:
            feature_ids = int(feature_id(feature)); row_id = feature_ids + 871
            CFE_com = int(table1.row(row)[5].value)
            CFE_back = int(table1.row(row)[6].value)
            All_com = int(table1.row(row)[7].value)
            All_back = int(table1.row(row)[8].value)

            chr = int(table2.row(row_id)[5].value)
            start_site = int(table2.row(row_id)[6].value)
            stop_site = int(table2.row(row_id)[7].value)
            n_Genes = int(table2.row(row_id)[8].value)
            Fragile_site_name = table2.row(row_id)[9].value
            if Fragile_site_name=="NA" or table2.cell(row_id,9).ctype == 0:
                Fragile_site_name = "NA";Fragile_site_class = ""
            else:
                if table2.row(row_id)[10].value== 1:Fragile_site_class="common"
                if table2.row(row_id)[11].value== 1:Fragile_site_class="rare"
            Locus = table2.row(row_id)[13].value
            Percentage = table2.row(row_id)[14].value
            Contained_genes = table2.row(row_id)[15].value

            f1.write(feature+"@"+str(CFE_com)+"@"+str(CFE_back)+"@"+str(All_com)+"@"+str(All_back)+"@"+str(chr)+"@"+str(start_site)+"@"+\
                     str(stop_site)+"@"+str(n_Genes)+"@"+Fragile_site_name+"@"+Fragile_site_class+"@"+Locus+"@"+str(Percentage)+"@"+str(Contained_genes)+"\n")
            count_RACS = count_RACS+1
    print("Count_RACS is %d"%count_RACS)
    f1.close()