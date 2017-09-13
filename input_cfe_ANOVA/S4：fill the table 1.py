#-*-coding:utf8-*-

import xlwt
import re
start_row = 1#全局变量
#对excel表的工作表名称和首行进行初始化
def table_initialization(workbook,sheet_name,row0_list):
    table = workbook.add_sheet(sheet_name)
    for i in range(len(row0_list)):
        table.write(0,i,row0_list[i])
    return table

#5个写入的函数
def write_annotation(table,count_dir,cancer_type,response,drug,drug_target):
    start_row = (count_dir - 1)*16+1
    for i in range(start_row,count_dir*16+1):
        table.write(i,0,"CD"+str(count_dir))
    list1 = ['primary_tumor_name','tumor_grade','evidence_type','evidence_direction','evidence_level','evidence_statement','evidence_link','clinical_significance', \
            'source_level','source','source_version','reference','curator','curate_time','drugs','drug target']
    list2 = [cancer_type,'','Predictive','Supports','','','',response,'ANOVA interactions','','','','jkang','2017/8/18',drug,drug_target]
    for i in range(16):
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
def write_complex(table,count_CI,feature,CI_num):
    table.write(count_CI,0,CI_num)
    table.write(count_CI,2,feature)
    table.write(count_CI,1,"ID"+str(count_CI))
    write_atomic(table_atomic,feature,count_CI)

def write_atomic(table,feature,count_CI):
    global start_row
    if feature.find("_fusion")!=-1:
        for i in range(start_row,start_row+3):
            table.write(i,0,"ID"+str(count_CI));table.write(i,1,"CI"+str(count_CI))
        table.write(start_row,2,"gene1");table.write(start_row+1,2,"type");table.write(start_row+2,2,"gene2");table.write(start_row+1,3,"fusion")
        atomic_list = re.split('[-_]',feature)
        table.write(start_row,3,atomic_list[0]);table.write(start_row+2,3,atomic_list[1])
        start_row = start_row + 3
    if feature.find(":")!=-1:
        for i in range(start_row,start_row+4):
            table.write(i,0,"ID"+str(count_CI));table.write(i,1,"CI"+str(count_CI));
        table.write(start_row,2,"type");table.write(start_row+1,2,"copy number status");table.write(start_row+2,2,"cna_ID");table.write(start_row+3,2,"gene");table.write(start_row,3,"RACS_ANOVA")
        status = feature.split(":")[0];cna_ID = re.split(r'[:(]',feature)[1];gene_start = feature.find("(");gene_end = feature.find(")")
        if status == "gain":
            table.write(start_row+1,3,"amplication");table.write(start_row+2,3,cna_ID)
            if gene_start!=-1:
                gene_content = feature[gene_start+1:gene_end]
                table.write(start_row+3,3,gene_content)
        if status == "loss":
            table.write(start_row+1,3,"deletion");table.write(start_row+2,3,cna_ID)
            if gene_start!=-1:
                gene_content = feature[gene_start+1:gene_end]
                table.write(start_row+3,3,gene_content)
        start_row = start_row + 4
    if feature.find("_mut")!=-1:
        gene = feature.split("_")[0]
        for i in range(start_row,start_row+2):
            table.write(i,0,"ID"+str(count_CI));table.write(i,1,"CI"+str(count_CI))
        table.write(start_row,2,"gene");table.write(start_row+1,2,"type");table.write(start_row+1,3,"gene mutations");table.write(start_row,3,gene)
        start_row = start_row + 2


if __name__ == '__main__':
    f1 = open("inter_CFE",'r',encoding='utf8')

    #创建各个工作表及其首行信息
    workbook = xlwt.Workbook(encoding='utf8')
    table_annotation = table_initialization(workbook,"annotation",[u'annotation_id',u'annotation_type',u'annotation_text'])
    table_directive = table_initialization(workbook,"collect_clinical_directive",[u'collect_clinical_directive_id',u'complexindication_id',u'therapeutic_stategy_id'])
    table_ther_stra = table_initialization(workbook,"therapeutic_stategy",[u'therapeutic_stategy_id',u'therapeutic_stategy_components_id'])
    table_ther_com = table_initialization(workbook,"therapeutic_stategy_components",[u'therapeutic_stategy_components_id',u'therapeutic_stategy_component_type',u'therapeutic_stategy_component'])
    table_complex = table_initialization(workbook,"complexindication",[u'complexindication_id',u'complexindication',u'origin_detail'])
    table_atomic = table_initialization(workbook,"atomicindication",[u'atomic_indication_id',u'complexindication_id',u'atomic_indication_type',u'atomic_indication'])

    count_dir = 0
    count_CI = 0
    count_TSID = 0
    dict_CI = {}
    dict_TSID = {}
    dict_TSCID = {}
    for line in f1:
        #获取要填入表中的信息
        elements = line.split("@")
        Cancer_type = elements[0]
        Feature = elements[1]
        Drug_name = elements[2]
        Drug_target = elements[3]
        Response = elements[4].strip("\n")

        count_dir = count_dir + 1

        #count_CI为数值，用于计数，记录complex的个数和写入行数
        #CI_num为字符串类型，用于描述写入的complex内容，可能和count_CI一致，也可能不一致

        #仅当CFE之前未出现过时，字典中添加这个CFE，complex表添加记录
        if Feature not in dict_CI:
            count_CI = count_CI + 1
            dict_CI[Feature] = "CI"+str(count_CI)
            CI_num = dict_CI[Feature]
            write_complex(table_complex,count_CI,Feature,CI_num)
        else:
            CI_num = dict_CI[Feature]

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
        write_annotation(table_annotation,count_dir,Cancer_type,Response,Drug_name,Drug_target)

    workbook.save(r'pmkb_CFE_annotation.xls')
    f1.close()