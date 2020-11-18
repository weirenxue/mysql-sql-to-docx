1. 將`sql`檔案`(*.sql)`放到此目錄下，並修改`fileInfo.json`  
    1. `sqlFileName`為`sql`檔案的檔名
    1. `docxFileName`為要輸出的`docx`檔名
    1. `pdfFileName`則是格式從`docx`轉為`pdf`後的檔名
    1. `tableDescriptionFileName`是資料表描述檔的檔名，必須為`json`
    1. 以上檔名皆需要包含副檔名。
1. 在此目錄中執行`genJson.py`，會產出上一步第四點所設定`tableDescriptionFileName`之檔案，裡面包含各資料表名稱，將描述填入資料表名稱後的雙引號(`""`)內即可。
1. 在此目錄中執行`main.py`，會產出`docx`與`pdf`檔，檔名皆為第一步所設定之檔名，若此時打開檔案，就會發現是美麗的資料表文件呈現在眼前。