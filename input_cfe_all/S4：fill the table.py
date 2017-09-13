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
def write_annotation(table,count_dir,CFE_occu,CFE_back,All_occu,All_back):
    start_row = (count_dir - 1)*18+1
    for i in range(start_row,count_dir*18+1):
        table.write(i,0,"CD"+str(count_dir))
    list1 = ['primary_tumor_name','tumor_grade','evidence_type','evidence_direction',\
             'evidence_level','evidence_statement','evidence_link','clinical_significance', \
            'source_level','source','source_version','reference','curator','curate_time',\
             'CFE occurrence in community','CFE background occurrence','All Event occurrences in community','All Event background occurrences']
    list2 = ["PANCAN",'','Predictive','Supports','','','','','CFE','','','','jkang','2017/8/23',CFE_occu,CFE_back,All_occu,All_back]
    for i in range(18):
        table.write(start_row+i,1,list1[i])
        table.write(start_row+i,2,list2[i])
def write_directive(table,count_dir):
    table.write(count_dir,0,"CD"+str(count_dir))
    table.write(count_dir,1,"CI"+str(count_dir))
    table.write(count_dir,2,"TSID1")

def write_complex(table,count_dir,feature):
    table.write(count_dir,0,"CI"+str(count_dir))
    table.write(count_dir,1,"ID"+str(count_dir))
    table.write(count_dir,2,feature)

def write_atomic_mutation(table,count_dir,feature):
    global start_row
    list = ["gene","type"]
    for i in range(start_row,start_row+2):
        table.write(i,0,"ID"+str(count_dir))
        table.write(i,1,"CI"+str(count_dir))
        table.write(i,2,list[i-start_row])
    table.write(start_row,3,feature.strip("_mut"))
    table.write(start_row+1,3,"gene mutations")
    start_row = start_row + 2

def write_atomic_RACS(table,count_dir,feature,chr,start_site,stop_site,n_Genes,\
                      Fragile_name,Fragile_class,locus,Percentage,Contained_genes):
    global start_row
    if feature.find("loss")!=-1:status = "deletion"
    else:status = "amplification"
    cna_id = re.split(r'[:(]',feature)[1].strip(" ")
    mentioned_genes = ""
    if feature.find("(")!=-1:
        mentioned_genes = re.split(r'[()]',feature)[1]
    list1=["type","copy number status","cna ID","CFE mentioned genes","chr","start","stop","n_Genes",\
          "FragileSiteName","Fragile Site Class","locus","Percentage of tumor samples","Contained genes"]
    list2=["RACS_CFE",status,cna_id,mentioned_genes,chr,start_site,stop_site,\
           n_Genes,Fragile_name,Fragile_class,locus,Percentage,Contained_genes]
    for i in range(start_row,start_row+13):
        table.write(i,0,"ID"+str(count_dir))
        table.write(i,1,"CI"+str(count_dir))
        table.write(i,2,list1[i-start_row])
        table.write(i,3,list2[i-start_row])
    start_row = start_row + 13

def write_atomic_HypMET(table,count_dir,count_cancer,chr,start,stop,gene,\
                        RefGene,Island,list_cancer,list_threshold,list_frequency):
    global start_row
    cancer_type = ",".join(list_cancer).strip(",")
    list_frequency_new = ["NA" if x=="" else x for x in list_frequency]
    threshold = ",".join(list_threshold)
    frequency = ",".join(list_frequency_new)
    list1 = ["type","chr","start","stop","gene","UCSC_RefGene_Group",'Relation_to_UCSC_CpG_Island',\
             'cancer type','beta value threshold','frequency']
    list2 = ["hypermethylation",chr,start,stop,gene,RefGene,Island,cancer_type,threshold,frequency]

    for i in range(start_row,start_row+10):
        table.write(i,0,"ID"+str(count_dir))
        table.write(i,1,"CI"+str(count_dir))
        table.write(i,2,list1[i-start_row])
        table.write(i,3,list2[i-start_row])
    start_row = start_row + 10


if __name__ == '__main__':
    f1 = open("mutation_info","r",encoding="utf8")
    f2 = open("RACS_info","r",encoding="utf8")
    f3 = open("HypMET_info","r",encoding="utf8")
    #创建各个工作表及其首行信息
    workbook = xlwt.Workbook(encoding='utf-8')
    table_annotation = table_initialization(workbook,"annotation",[u'annotation_id',u'annotation_type',u'annotation_text'])
    table_directive = table_initialization(workbook,"collect_clinical_directive",[u'collect_clinical_directive_id',u'complexindication_id',u'therapeutic_stategy_id'])
    table_ther_stra = table_initialization(workbook,"therapeutic_stategy",[u'therapeutic_stategy_id',u'therapeutic_stategy_components_id'])
    table_ther_com = table_initialization(workbook,"therapeutic_stategy_components",[u'therapeutic_stategy_components_id',u'therapeutic_stategy_component_type',u'therapeutic_stategy_component'])
    table_complex = table_initialization(workbook,"complexindication",[u'complexindication_id',u'complexindication',u'origin_detail'])
    table_atomic = table_initialization(workbook,"atomicindication",[u'atomic_indication_id',u'complexindication_id',u'atomic_indication_type',u'atomic_indication'])
    table_ther_stra.write(1,0,"TSID1");table_ther_stra.write(1,1,"TSCID1")
    table_ther_com.write(1,0,"TSCID");table_ther_com.write(1,2,"No therapy.")

    count_dir = 0
    for line in f1:
        count_dir = count_dir + 1
        list_feature = line.split("@")
        feature = list_feature[0]
        CFE_occu = list_feature[1]
        CFE_back = list_feature[2]
        All_occu = list_feature[3]
        All_back = list_feature[4].strip("\n")
        write_annotation(table_annotation, count_dir, CFE_occu, CFE_back, All_occu, All_back)
        write_directive(table_directive,count_dir)
        write_complex(table_complex,count_dir,feature)
        write_atomic_mutation(table_atomic,count_dir,feature)
    for line in f2:
        count_dir = count_dir + 1
        line = line.strip("\n")
        list_feature = line.split("@")

        feature = list_feature[0];CFE_occu = list_feature[1];CFE_back = list_feature[2]
        All_occu = list_feature[3];All_back = list_feature[4];chr = list_feature[5]
        start = list_feature[6];stop = list_feature[7];n_Genes = list_feature[8]
        Fragile_name = list_feature[9];Fragile_class = list_feature[10];locus = list_feature[11]
        Percentage = list_feature[12];Contained_genes = list_feature[13]

        write_annotation(table_annotation, count_dir, CFE_occu, CFE_back, All_occu, All_back)
        write_directive(table_directive,count_dir)
        write_complex(table_complex,count_dir,feature)
        write_atomic_RACS(table_atomic, count_dir, feature, chr, start, stop, n_Genes, \
                          Fragile_name, Fragile_class, locus, Percentage, Contained_genes)
    for line in f3:
        count_dir = count_dir + 1
        line = line.strip("\n")
        list_feature = line.split("@")
        count_cancer = int((len(list_feature)-13)/3)
        list_cancer = []; list_threshold = []; list_frequency = []

        feature = list_feature[0];CFE_occu = list_feature[1];CFE_back = list_feature[2]
        All_occu = list_feature[3];All_back = list_feature[4];chr = list_feature[5]
        start = list_feature[6];stop = list_feature[7];gene = list_feature[8]
        RefGene = list_feature[9];Island = list_feature[10];
        for i in range(12,len(list_feature)):
            if (i-12)%3 == 0:
                list_cancer.append(list_feature[i])
            if (i-12)%3 == 1:
                list_threshold.append(list_feature[i])
            if (i-12)%3 == 2:
                list_frequency.append(list_feature[i])
        write_annotation(table_annotation, count_dir, CFE_occu, CFE_back, All_occu, All_back)
        write_directive(table_directive,count_dir)
        write_complex(table_complex,count_dir,feature)
        write_atomic_HypMET(table_atomic, count_dir, count_cancer, chr, start, stop, gene, \
                            RefGene, Island, list_cancer, list_threshold, list_frequency)

    print("Count_directive is %d"%count_dir)
    workbook.save(r'pmkb_CFE_All.xls')
    f1.close()
    f2.close()
    f3.close()