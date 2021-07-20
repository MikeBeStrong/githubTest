import xlwings as xw
import sqlite3

def check_frdm(str):
    if int(str[-1])%2==1:
        return '法人代码'+str+'校验错'
    else:
        return '法人代码'+str+'校验正确'

workbook = xw.books.open(r'D:\njc\Python学习\withSqlite\输出：审核结果.xlsx')

workbook = xw.books.open(r'D:\njc\Python学习\withSqlite\input-zzy.xlsx')
'''
worksheet = workbook.sheets[0]
value = worksheet.range('A2').expand('table').value

#print(value)
#for row in value:
#    print(row)

con=sqlite3.connect("test.db")
con.create_function('check_frdm',1,check_frdm)
cur=con.cursor()
cur.execute('delete from zzy')
print('delete from zzy')

cur.execute('select * from zzy')
print('select * from zzy -- result:')
for row in cur:
    print(row)

sql = 'insert into zzy values (?, ?, ?);'
print('executemany... ',sql)

cur.executemany(sql,value)
con.commit()

cur.execute('select * from zzy')
print('select * from zzy -- result:')
for row in cur:
    print(row)

cur.execute('select check_frdm(frdm),* from zzy')
print('select check_frdm(frdm),* from zzy  --  result:')
for row in cur:
    print(row)

print('select check_frdm(frdm),* from zzy  --  result to xlsx:')
cur.execute('select check_frdm(frdm),* from zzy')
rows=cur.fetchall()
print(rows)

new_workbook = xw.books.add()
new_worksheet = new_workbook.sheets[0]
new_worksheet.name='审核结果'
new_worksheet['A1'].value=['frdm校验结果','id','frdm','frmc']
new_worksheet['A2'].value=rows
new_worksheet.autofit()
new_workbook.save(r'输出：审核结果.xlsx')
new_workbook.close()

cur.close()
con.close()

input('Press <Enter> to quit: ')
'''