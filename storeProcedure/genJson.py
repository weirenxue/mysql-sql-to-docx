import json
import os

fileInfo = json.load(open("./fileInfo.json", "r", encoding="utf-8"))

#判斷是為了避免覆蓋到已經寫好描述的資料
if os.path.isfile(fileInfo["descriptionFileName"]):
    targetDes = json.load(open(fileInfo["descriptionFileName"], "r", encoding="utf-8"))
else:
    targetDes = {}

with open(fileInfo["sqlFileName"], "r", encoding = "utf-8") as inputContent:
        lines = inputContent.readlines() #讀進全部
        inTarget = False #是否在Procedure的範圍
        target = targetDes
        for line in lines: 
            lineU = line.upper() # 都轉成大寫，統一比對
            lineUSplit = lineU.split()
            #是否即將脫離Procedure的範圍
            if "$$" in line: 
                inTarget = False
            elif "PROCEDURE" in lineU:
                targetName = line.split()[3].strip("`") #前後拿掉`符號
                #若SP不在
                if targetName not in target:
                    target[targetName] = {}
                #若SP的description不在
                if 'description' not in target[targetName]:
                    target[targetName]['description'] = ""
                #若SP的inputParameter不在
                if 'inputParameter' not in target[targetName]:
                    target[targetName]['inputParameter'] = {}
                if 'outputResult' not in target[targetName]:
                    target[targetName]['outputResult'] = ""
                indices = [i for i, element in enumerate(lineUSplit) if element == "(IN" or element == "IN"]
                #抓出輸入參數
                for index in indices:
                    x = {}
                    #若參數不在列表中
                    if line.split()[index + 1].strip('`,)') not in targetDes[targetName]['inputParameter']:
                        targetDes[targetName]['inputParameter'][line.split()[index + 1].strip('`,)')] = ''
                
                inTarget = True
#排序，由A到Z排序
targetDesSorted = {key:targetDes[key] for key in sorted(targetDes, key = lambda i: (i))}

with open(fileInfo["descriptionFileName"], 'w', encoding="utf-8") as outfile:
    #要有indent，輸出才會漂亮；ensure_ascii=False才可輸出中文，否則會輸出unicode
    json.dump(targetDesSorted, outfile, indent=4, ensure_ascii=False)