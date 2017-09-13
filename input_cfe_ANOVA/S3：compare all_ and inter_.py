#-*-coding:utf8-*-

f1 = open("all_CFE","r")
f2 = open("inter_CFE","r")

List_CFE = []
count_not_pair = 0
for line in f1:
    feature = line.strip("\n")
    List_CFE.append(feature)

for line in f2:
    feature = line.split("@")[1]
    if feature not in List_CFE:
        count_not_pair = count_not_pair + 1
        print(feature)
        print

print
print("count is %d"%count_not_pair)
f1.close()
f2.close()
