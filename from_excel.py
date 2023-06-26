import pymysql
from openpyxl import Workbook

conn = mysql.connector.connect(user='root', host='localhost', database='abcd_dress-500')
cursor = conn.cursor()
sql = "select * from dresses"
cursor.execute(sql)
conn.commit()

wb=Workbook()
ws=wb.active

ws["A1"].value="id"
ws["B1"].value="name"
ws["C1"].value="description"
ws["D1"].value="did_you_know"
ws["E1"].value="category"
ws["F1"].value="type"
ws["G1"].value="state_name"
ws["H1"].value="key_words"
ws["I1"].value="image_url"
ws["J1"].value="status"
ws["K1"].value="notes"

export=cursor.fetchall()
for x in export:
    ws.append(x)
    wb.save('dresses.xlsx')
conn.close()