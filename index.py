# Read! => .Readme File
import sqlite3
import random
import xlsxwriter
import statistics 
from random import randint, uniform
conn = sqlite3.connect('PersonalDB.db')
mycursor = conn.cursor()
records = []
stat = []
personWorkingTimes = []
alp = ["ა","ბ","გ","დ","ე","ვ","ზ","თ","ი","კ","ლ","მ","ნ","ო","პ","ჟ","რ","ს","ტ","უ","ფ","ქ","ღ","ყ","შ","ჩ","ც","ძ","წ","ჭ","ხ","ჯ","ჰ"]

def getRandomName():
    chars = []
    for j in range(0, 5):
        index = randint(0, 32)
        chars.append(alp[index])
    return("".join(chars))

def getRandomLastName():
    chars = []
    for j in range(0, 10):
        index = randint(0, 32)
        chars.append(alp[index])
    return("".join(chars))

def getRandomAge():
    index = randint(20, 60)
    return(str(index))

def getRandomPersonalN():
    chars = []
    for j in range(0, 11):
        index = randint(0, 9)
        chars.append(str(index))
    return("".join(chars))

def getRandomWorkTime():
    sum = 0
    stat = []
    for j in range(0, 30):
        index = uniform(0, 12)
        stat.append(index)
        sum += index
    personWorkingTimes.append(stat)
    return("{:.1f}".format(sum))

for i in range(0, 100): 
    person = {"name": getRandomName(), "lastName": getRandomLastName(), "age": getRandomAge(), "personalN": getRandomPersonalN(), "workTime": getRandomWorkTime()}
    records.append(person)
    sql = "INSERT INTO person (name, lastName, age, personalN, workTime)  VALUES ('"+person["name"]+"', '"+person["lastName"]+"', '"+person["age"]+"', '"+person["personalN"]+"', '"+person["workTime"]+"')"
    mycursor.execute(sql)
conn.commit()

# Step2

def getAverage():
    sum = 0
    for i in records:
        sum += int(i['age'])
    return(sum / len(records))
average = getAverage()

workbook = xlsxwriter.Workbook('personal.xlsx')
worksheet = workbook.add_worksheet()
worksheet2 = workbook.add_worksheet()
row = 0
row2 = 0
col = 0
worksheet.write(row, col,     "Name")
worksheet.write(row, col + 1, "Last Name")
worksheet.write(row, col + 2, "Age")
worksheet.write(row, col + 3, "Personal Number")
worksheet.write(row, col + 4, "Work Time")
worksheet.write(row, col + 5, "Average Work Time")
worksheet2.write(row, col, "Statistics")

row += 1
row2 += 1

for i in records:
    if int(i['age']) > average :
        worksheet.write(row, col,     str(i["name"]))
        worksheet.write(row, col + 1, str(i["lastName"]))
        worksheet.write(row, col + 2, str(i["age"]))
        worksheet.write(row, col + 3, str(i["personalN"]))
        worksheet.write(row, col + 4, str(i["workTime"]))
        worksheet.write(row, col + 5, round(float(i["workTime"])))
        row += 1

for j in personWorkingTimes:
    worksheet2.write(row2, col , (statistics.stdev(j)))
    row2 += 1

chart = workbook.add_chart({'type': 'line'})
chart.add_series({'values': '=Sheet1!$F$1:$F$'+str(row-1)})
worksheet.insert_chart('H5', chart)
workbook.close()