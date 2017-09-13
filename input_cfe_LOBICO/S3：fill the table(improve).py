#-*-coding:utf8-*-

import xlwt
import re

start_row = 1
#对excel表的工作表名称和首行进行初始化
def table_initialization(workbook,sheet_name,row0_list):
    table = workbook.add_sheet(sheet_name)
    for i in range(len(row0_list)):
        table.write(0,i,row0_list[i])
    return table

#5个写入的函数
def write_annotation(table,count_dir,cancer_type,response,drug,drug_target,drug_class,sen_ccl,res_ccl,precision,sensitivity):
    start_row = (count_dir - 1)*21+1
    for i in range(start_row,count_dir*21+1):
        table.write(i,0,"CD"+str(count_dir))
    list1 = ['primary_tumor_name','tumor_grade','evidence_type','evidence_direction','evidence_level','evidence_statement','evidence_link','clinical_significance', \
            'source_level','source','source_version','reference','curator','curate_time','drugs','drug target','drug class','Sensitive cell lines','Resistant cell lines',\
             'Precision','Sensitivity(Recall)']
    list2 = [cancer_type,'','Predictive','Supports','','','',response,'LOBICO','','','','jkang','2017/8/19',drug,drug_target,drug_class,sen_ccl,res_ccl,precision,sensitivity]
    for i in range(21):
        table.write(start_row+i,1,list1[i])
        table.write(start_row+i,2,list2[i])
def write_directive(table,count_dir,CD_num,CI_num,TSIC_num):
    table.write(count_dir,0,CD_num)
    table.write(count_dir,1,CI_num)
    table.write(count_dir,2,TSIC_num)
def write_ther_stra(table,count_TSID,TSID_num,TSCID_num):
    table.write(count_TSID,0,TSID_num)
    table.write(count_TSID,1,TSCID_num)
def write_ther_com(table,count_TSID,drug,TSCID_num):
    table.write(count_TSID,0,TSCID_num)
    table.write(count_TSID,2,drug)

def write_complex(table,count_CI,Model,CI_num,dict_ID):
    global start_row
    table.write(count_CI,0,CI_num)
    table.write(count_CI,2,Model)
    list_atomic = re.split(r'[&|]',Model)
    #把Model里的atomic添加到dict_ID中
    for i in list_atomic:
        i = i.strip(" ")
        #除了mutation外，其他类型的CFE都没有¬符号。
        if i.find("¬")!=-1 and (i.find("a") != -1 or i.find("d")!=-1 or i.find("-UP")!=-1 or i.find("-DOWN")!=-1):
            i = i.strip("¬")
        if i not in dict_ID:
            #对应的是第几个ID
            len_dict = len(dict_ID)
            dict_ID[i] = len_dict + 1
            write_atomic(table_atomic,i,len_dict+1,count_CI)
    len_list = len(list_atomic)
    #获取每个model的ID并添加到list_ID中
    list_ID = []
    for i in list_atomic:
        #除了mutation外，其他类型的CFE都有not
        if i.find("¬")!=-1 and (i.find("a") != -1 or i.find("d")!=-1 or i.find("-UP")!=-1 or i.find("-DOWN")!=-1):
            list_ID.append("not(ID"+str(dict_ID[i.strip(" ").strip("¬")])+")")
        else:
            list_ID.append("ID" + str(dict_ID[i.strip(" ")]))
    #分情况进行write
    if len_list == 1:
        table.write(count_CI,1,list_ID[0])
    else:
        if Model.find("&")!=-1 and Model.find("|")==-1:
            complex_indication = str(list_ID).lstrip("[").rstrip("]").replace(","," and")
            table.write(count_CI,1,complex_indication)
        if Model.find("|")!=-1 and Model.find("&")==-1:
            complex_indication = str(list_ID).lstrip("[").rstrip("]").replace(","," or")
            table.write(count_CI,1,complex_indication)
        if Model.find("|")!=-1 and Model.find("&")!=-1 and len_list==4:
            complex_indication = "("+list_ID[0]+" and "+list_ID[1]+")"+" or "+"("+list_ID[2]+" and "+list_ID[3]+")"
            table.write(count_CI,1,complex_indication)
        if Model.find("|") != -1 and Model.find("&") != -1 and len_list == 3:
            if Model.find("|")>Model.find("&"):
                complex_indication = "("+list_ID[0]+" and "+list_ID[1]+")"+" or "+list_ID[2]
                table.write(count_CI, 1, complex_indication)
            else:
                complex_indication = list_ID[0] + " or "+"("+list_ID[1]+" and "+list_ID[2]+")"
                table.write(count_CI, 1, complex_indication)


def write_atomic(table,atomic,count_ID,count_CI):
    global start_row
    #gene mutation
    if atomic.find("a")==-1 and atomic.find("d")==-1 and atomic.find("-UP")==-1 and atomic.find("-DOWN")==-1:
        table.write(start_row,0,"ID"+str(count_ID));table.write(start_row+1,0,"ID"+str(count_ID));table.write(start_row,1,"CI"+str(count_CI));table.write(start_row+1,1,"CI"+str(count_CI));
        table.write(start_row,2,"gene");table.write(start_row+1,2,"type");table.write(start_row,3,atomic.strip("¬"))
        if atomic.find("¬")!=-1:
            table.write(start_row+1,3,"no mutations")
        else:
            table.write(start_row+1,3,"gene mutations")
        start_row = start_row+2
    #RACS
    if (atomic.find("a")!=-1 or atomic.find("d")!=-1):
        for i in range(start_row,start_row+4):
            table.write(i,0,"ID"+str(count_ID));table.write(i,1,"CI"+str(count_CI))
        table.write(start_row,3,"RACS_LOBICO");table.write(start_row,2,"type");table.write(start_row+1,2,"copy number status");table.write(start_row+2,2,"location type")
        if atomic.find("(")!=-1:
            table.write(start_row+3,2,"location");table.write(start_row+2,3,"genes");table.write(start_row+3,3,atomic.lstrip("¬").lstrip("a(").lstrip("d(").rstrip(")"))
            if atomic.find("a")!=-1:
                table.write(start_row+1,3,"amplication")
            else:
                table.write(start_row+1,3,"deletion")
        else:
            table.write(start_row+3,2,"location");table.write(start_row+2,3,"locus");table.write(start_row+3,3,atomic.lstrip("a").lstrip("d"))
            if atomic.find("a")!=-1:
                table.write(start_row+1,3,"amplication")
            else:
                table.write(start_row+1,3,"deletion")
        start_row = start_row+4
    #pathways
    if atomic.find("-UP")!=-1 or atomic.find("-DOWN")!=-1:
        list = ["type","pathway","status"]
        for i in range(start_row,start_row+3):
            table.write(i,0,"ID"+str(count_ID));table.write(i,1,"CI"+str(count_CI))
        for i in range(3):
            table.write(start_row+i,2,list[i])
        table.write(start_row,3,"pathway activity");table.write(start_row+1,3,atomic.rstrip("-UP").rstrip("-DOWN"))
        if atomic.find("-UP")!=-1:
            table.write(start_row+2,3,"up")
        else:
            table.write(start_row+2,3,"down")
        start_row = start_row+3



if __name__ == '__main__':
    f1 = open("gather_info","r",encoding="utf8")
    #创建各个工作表及其首行信息
    workbook = xlwt.Workbook(encoding='utf-8')
    table_annotation = table_initialization(workbook,"annotation",[u'annotation_id',u'annotation_type',u'annotation_text'])
    table_directive = table_initialization(workbook,"collect_clinical_directive",[u'collect_clinical_directive_id',u'complexindication_id',u'therapeutic_stategy_id'])
    table_ther_stra = table_initialization(workbook,"therapeutic_stategy",[u'therapeutic_stategy_id',u'therapeutic_stategy_components_id'])
    table_ther_com = table_initialization(workbook,"therapeutic_stategy_components",[u'therapeutic_stategy_components_id',u'therapeutic_stategy_component_type',u'therapeutic_stategy_component'])
    table_complex = table_initialization(workbook,"complexindication",[u'complexindication_id',u'complexindication',u'origin_detail'])
    table_atomic = table_initialization(workbook,"atomicindication",[u'atomic_indication_id',u'complexindication_id',u'atomic_indication_type',u'atomic_indication'])

    count_dir = 0
    count_CI = 0
    count_TSID = 0
    count_ID = 0
    dict_CI = {}
    dict_TSID = {}
    dict_TSCID = {}
    dict_ID = {}
    for line in f1:
        #获取要填入表中的信息
        elements = line.split("@")
        Cancer_type = elements[0]
        Drug_name = elements[1]
        Drug_target = elements[2]
        Drug_class = elements[3]
        Sen_ccl = elements[4]
        Res_ccl = elements[5]
        Model = elements[6]
        Precision = elements[7]
        Sensitivity =elements[8]
        Response = "Sensitive"

        count_dir = count_dir + 1

        #count_CI为数值，用于计数，记录complex的个数和写入行数
        #CI_num为字符串类型，用于描述写入的complex内容，可能和count_CI一致，也可能不一致

        #仅当CFE之前未出现过时，字典中添加这个CFE，complex表添加记录
        if Model not in dict_CI:
            count_CI = count_CI + 1
            dict_CI[Model] = "CI"+str(count_CI)
            CI_num = dict_CI[Model]
            write_complex(table_complex,count_CI,Model,CI_num,dict_ID)
            #print(Model)
        else:
            CI_num = dict_CI[Model]

        #仅当这个drug未出现过时，字典中加入这个drug，且在ther_stra表和ther_com表中添加记录
        if Drug_name not in dict_TSID:
            count_TSID = count_TSID + 1
            dict_TSID[Drug_name] = "TSID"+str(count_TSID)
            dict_TSCID[Drug_name] = "TSCID"+str(count_TSID)
            TSID_num = dict_TSID[Drug_name]
            TSCID_num = dict_TSCID[Drug_name]
            write_ther_stra(table_ther_stra,count_TSID,TSID_num,TSCID_num)
            write_ther_com(table_ther_com,count_TSID,Drug_name,TSCID_num)
        else:
            TSID_num = dict_TSID[Drug_name]
            TSCID_num = dict_TSCID[Drug_name]
        CD_num = "CD"+str(count_dir)

        #写入directive
        write_directive(table_directive,count_dir,CD_num,CI_num,TSID_num)

        #写入annotation
        write_annotation(table_annotation,count_dir,Cancer_type,Response,Drug_name,Drug_target,Drug_class,Sen_ccl,Res_ccl,Precision,Sensitivity)

    workbook.save(r'pmkb_CFE_LOBICO.xls')
    f1.close()