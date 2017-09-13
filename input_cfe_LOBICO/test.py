#-*-coding:utf8-*-
import re

if __name__ == '__main__':
    f1 = open('gather_info','r')
    for line in f1:
        model = line.split("@")[6]
        list_atomic = re.split(r'[&|]',model)
        #print(list_atomic)
        if model.find("&")!=-1 and model.find("|")!=-1 and len(list_atomic)!=4:
            print(model)

    f1.close()

