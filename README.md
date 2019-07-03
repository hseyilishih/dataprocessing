# os.walk.py

1.	使用Spyder IDE, 建立base以外的獨立工作環境,import所需套件
2.  input 是 xxx.sql 的檔案
3.	使用 os 套件, 歷遍工作環境內一個資料夾內的 .sql 檔案
4.	盤點出所有檔案名稱, 排除不需要納入盤點的檔案(eg. .py)
5.	使用pandas + list + string + if else + for loop + find 處理, 
    將每個 .sql 程式內文, 所需要的訊息摘取出來 (eg. program name, create table name, primary key, column name, datatype)
6.	使用 xlwings 套件, 處理excel column heading 背景顏色/font color/自動欄寬
7.	輸出成excel, 檔名帶著當天日期, 內有3個worksheet, 每個row會編流水號, 會count total row/column
