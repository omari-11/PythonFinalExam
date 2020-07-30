# Read! => .Readme File
import sqlite3
import random
import xlsxwriter
import statistics 
from random import randint, uniform
conn = sqlite3.connect('condDB.db')
mycursor = conn.cursor()

records = []
stat = []

personWorkingTimes = []
alp = ["A","B","C","D","E","F","G","H","I","K","L","M","N","O","P","Q","R","S","T","V","X","Y","Z"]

def getModel():
    chars = []
    for j in range(0, 5):
        index = randint(0, 22)
        chars.append(alp[index])
    return("".join(chars))

def getPrice():
    index = randint(500, 5000)
    return(str(index))

def getCondIdNumber():
    chars = []
    for j in range(0, 11):
        index = randint(0, 9)
        chars.append(str(index))
    return("".join(chars))

def getWorkTime():
    sum = 0
    stat = []
    for j in range(0, 30):
        index = uniform(0, 12)
        stat.append(index)
        sum += index
    personWorkingTimes.append(stat)
    return("{:.1f}".format(sum))

for i in range(0, 100): 
    cond = {"model": getModel(), "price": getPrice(), "condID": getCondIdNumber(), "workTime": getWorkTime()}
    records.append(cond)
    sql = "INSERT INTO cond (model,  price, condID, workTime)  VALUES ('"+cond["model"]+"', '"+cond["price"]+"', '"+cond["condID"]+"', '"+cond["workTime"]+"')"
    mycursor.execute(sql)
conn.commit()

# Step2

def getSoldCond():
    sum = 0
    for i in records:
        sum += int(i['price'])
    return(sum / len(records))
sodlCond = getSoldCond()

workbook = xlsxwriter.Workbook('cond.xlsx')
worksheet = workbook.add_worksheet()
worksheet2 = workbook.add_worksheet()
row = 0
row2 = 0
col = 0
worksheet.write(row, col,     "Model")
worksheet.write(row, col + 1, "Price")
worksheet.write(row, col + 2, "Cond ID Number")
worksheet.write(row, col + 3, "Work Time")
worksheet.write(row, col + 4, "Sold Cond")
worksheet2.write(row, col, "Statistics")

row += 1
row2 += 1

for i in records:
    if int(i['price']) > sodlCond :
        worksheet.write(row, col,     str(i["model"]))
        worksheet.write(row, col + 1, str(i["price"]))
        worksheet.write(row, col + 2, str(i["condID"]))
        worksheet.write(row, col + 3, str(i["workTime"]))
        worksheet.write(row, col + 4, round(float(i["workTime"])))
        row += 1

for j in personWorkingTimes:
    worksheet2.write(row2, col , (statistics.stdev(j)))
    row2 += 1

chart = workbook.add_chart({'type': 'line'})
chart.add_series({'values': '=Sheet1!$E$1:$E$'+str(row-1)})
worksheet.insert_chart('H5', chart)
workbook.close()