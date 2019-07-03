# -*- coding: utf-8 -*-
"""
Created on Wed Jul  3 11:37:32 2019

@author: 000147A
https://docs.xlwings.org/en/stable/api.html
https://docs.xlwings.org/en/stable/converters.html
https://www.itread01.com/content/1535218689.html
https://www.jianshu.com/p/b534e0d465f7
"""

import os
import xlwings as xw
from xlwings.utils import rgb_to_int
import xlwings.constants
import pandas as pd

from datetime import date
today = date.today()
# dd/mm/YY
str_today = today.strftime("%Y-%m-%d")
#print("d1 =", d1)

#df = pd.read_csv('效能調教 INDEX .sql', sep=r'\n', engine='python', header=None, encoding='utf8', names=['statement'])

output_list =[]
output_nametable =[]
output_namecol =[]
 


for dirPath, dirNames, fileNames in os.walk("Z:\Project EDW_Aaron_2019_晉升考核\晉升考核Table_Schema"):
    #print (dirPath, dirNames, fileNames)
    for f in fileNames:
        if "sql" in f: 
            output_list.append(str(f).replace('開表','').replace('.sql',''))
            #print('========================================', f)
            #print (os.path.join(dirPath, f))
            df = pd.read_csv(f, sep=r'\n', engine='python', header=None, encoding='utf8', names=['statement'])      
            #print(df)
            name=''
            comment= ''
            pk = ''
            for x in df['statement']:
                #print( x)
                
                if 'CREATE INDEX' in x.upper():
                    continue                
                elif 'CREATE TABLE' in x and '--' not in x:
                    name = (x.replace('CREATE TABLE','').replace('(','').strip())
                    #print(name)                           
                elif 'COMMENT ON TABLE' in x:
                    comment = x.replace('COMMENT ON TABLE','').replace(name,'').replace('IS','').replace(';','').replace("'",'').strip()
                    #print(comment)                    
                elif 'PRIMARY KEY' in x.upper() and '--' not in x:
                    position=x.upper().find('PRIMARY KEY')
                    pk = (x[position:].replace('PRIMARY KEY','').strip('( );') )   
                    
                if ' NUMBER'in x.upper() and '--' not in x:
                    position=x.upper().find('NUMBER')
                    name_col = (x[:position].replace('NUMBER','').strip('( );') )  
                    output_namecol.append((f,name,name_col,'NUMBER'))                    
                elif ' VARCHAR2'in x.upper() and '--' not in x:
                    position=x.upper().find('VARCHAR2')
                    name_col = (x[:position].replace('VARCHAR2','').strip('( );') )  
                    output_namecol.append((f,name,name_col,'VARCHAR2'))                     
                elif ' CHAR'in x.upper() and '--' not in x:
                    position=x.upper().find('CHAR')
                    name_col = (x[:position].replace('char','').strip('( );') )  
                    output_namecol.append((f,name,name_col,'CHAR'))                       
                elif ' DATE NULL'in x.upper() and '--' not in x and 'COMMENT' not in x:
                    position=x.upper().find('DATE NULL')
                    name_col = (x[:position].strip('( );') )  
                    output_namecol.append((f,name,name_col,'DATE'))   

                                           
            #if name > '' or comment > '':
            output_nametable.append((f,name, comment, pk))
                                        
                    

#print(output_list)
print(len(output_list))
print(len(output_nametable))

print(*output_nametable, sep = "\n")  
print(*output_namecol , sep = "\n")  

#df = pd.DataFrame(output_list)
#df.columns = ['name_object']
#df


wb = xw.Book()

sht  = wb.sheets.add('list_SQL')
sht.clear()
sht.range('b1').options(transpose=True).value = output_list #直接從list輸出

sht2 = wb.sheets.add('list_table')
sht2.clear()
sht2.range('b2').value = output_nametable #直接從list輸出

sht3 = wb.sheets.add('list_column')
sht3.clear()
sht3.range('b2').value = output_namecol #直接從list輸出

#sht.range("A1").value = df
sht.api.rows('1:1').insert  #新增第一ROW當作column heading
sht.range('a1').value = ['seq','name_object']

#sht2.api.rows('1:1').insert  #新增第一ROW當作column heading
sht2.range('a1').value = ['seq','name_sql', 'name_createtable', 'label_table', 'primary_key']
sht3.range('a1').value = ['seq','name_sql', 'name_createtable', 'name_column', 'datatype']


rng = sht.range('a1:z1')
rng.color=(180,198,230)
rng.api.Font.Color = rgb_to_int((20, 20, 255))

rng = sht2.range('a1:z1')
rng.color=(180,198,230)
rng.api.Font.Color = rgb_to_int((20, 20, 255))

rng = sht3.range('a1:z1')
rng.color=(180,198,230)
rng.api.Font.Color = rgb_to_int((20, 20, 255))

# sht.range('A1:A2').api.merge #合并单元格
# 返回和设置当前格子的高度和宽度
#print(rng.width)
#print(rng.height)
#rng.row_height=40

#sht.range('c1').add_hyperlink(r'www.transglobe.com.tw','全球人壽','鏈接到')


ttl_row = sht.api.Cells.Find(What="*",
                       After=sht.api.Cells(1, 1),
                       LookAt=xlwings.constants.LookAt.xlPart,
                       LookIn=xlwings.constants.FindLookIn.xlFormulas,
                       SearchOrder=xlwings.constants.SearchOrder.xlByRows,
                       SearchDirection=xlwings.constants.SearchDirection.xlPrevious,
                       MatchCase=False)

ttl_col = sht.api.Cells.Find(What="*",
                      After=sht.api.Cells(1, 1),
                      LookAt=xlwings.constants.LookAt.xlPart,
                      LookIn=xlwings.constants.FindLookIn.xlFormulas,
                      SearchOrder=xlwings.constants.SearchOrder.xlByColumns,
                      SearchDirection=xlwings.constants.SearchDirection.xlPrevious,
                      MatchCase=False)

#print((ttl_row.Row, ttl_col.Column))

for i in range(2,int(ttl_row.Row)+1):
    sht.range('a'+str(i)).value = i-1


ttl_row = sht2.api.Cells.Find(What="*",
                       After=sht.api.Cells(1, 1),
                       LookAt=xlwings.constants.LookAt.xlPart,
                       LookIn=xlwings.constants.FindLookIn.xlFormulas,
                       SearchOrder=xlwings.constants.SearchOrder.xlByRows,
                       SearchDirection=xlwings.constants.SearchDirection.xlPrevious,
                       MatchCase=False)

ttl_col = sht2.api.Cells.Find(What="*",
                      After=sht.api.Cells(1, 1),
                      LookAt=xlwings.constants.LookAt.xlPart,
                      LookIn=xlwings.constants.FindLookIn.xlFormulas,
                      SearchOrder=xlwings.constants.SearchOrder.xlByColumns,
                      SearchDirection=xlwings.constants.SearchDirection.xlPrevious,
                      MatchCase=False)
for i in range(2,int(ttl_row.Row)+1):
    sht2.range('a'+str(i)).value = i-1
    

ttl_row = sht3.api.Cells.Find(What="*",
                       After=sht.api.Cells(1, 1),
                       LookAt=xlwings.constants.LookAt.xlPart,
                       LookIn=xlwings.constants.FindLookIn.xlFormulas,
                       SearchOrder=xlwings.constants.SearchOrder.xlByRows,
                       SearchDirection=xlwings.constants.SearchDirection.xlPrevious,
                       MatchCase=False)

ttl_col = sht3.api.Cells.Find(What="*",
                      After=sht.api.Cells(1, 1),
                      LookAt=xlwings.constants.LookAt.xlPart,
                      LookIn=xlwings.constants.FindLookIn.xlFormulas,
                      SearchOrder=xlwings.constants.SearchOrder.xlByColumns,
                      SearchDirection=xlwings.constants.SearchDirection.xlPrevious,
                      MatchCase=False)
for i in range(2,int(ttl_row.Row)+1):
    sht3.range('a'+str(i)).value = i-1
    
sht.autofit()
sht2.autofit()
sht3.autofit()
#sht3.range('d1:d20').column_width=30

sht4=wb.sheets['工作表1']
sht4.delete()


wb.save(str('0 output_list_'+str_today + '.xlsx'))
