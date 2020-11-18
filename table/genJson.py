import json
import os

fileInfo = json.load(open("./fileInfo.json", "r", encoding="utf-8"))

#判斷是為了避免覆蓋到已經寫好描述的資料
if os.path.isfile(fileInfo["descriptionFileName"]):
    tableDes = json.load(open(fileInfo["descriptionFileName"], "r", encoding="utf-8"))
else:
    tableDes = {}

with open(fileInfo["sqlFileName"], "r", encoding = "utf-8") as inputContent:
        lines = inputContent.readlines() #讀進全部
        for line in lines: 
            if "CREATE TABLE" in line.upper():
                tableName = line.split()[2].strip("`")
                #若已經有描述的，不需要再加入節點
                if tableName in tableDes:
                    continue
                tableDes[tableName] = ""

#排序，由A到Z排序
tableDesSorted = {key:tableDes[key] for key in sorted(tableDes, key = lambda i: (i))}


with open(fileInfo["descriptionFileName"], 'w', encoding="utf-8") as outfile:
    #要有indent，輸出才會漂亮；ensure_ascii=False才可輸出中文，否則會輸出unicode
    json.dump(tableDesSorted, outfile, indent=4, ensure_ascii=False)