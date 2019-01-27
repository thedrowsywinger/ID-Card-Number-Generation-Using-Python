'''
In this version, i was asked to generate id in the 
following format: CITY - DESIGNATION - YEAR THEY JOINED - SERIAL(based on the year they joined) 

'''

import xlrd, xlsxwriter

loc = ("./memberpractise.xlsx")

workbook = xlsxwriter.Workbook("output.xlsx")
worksheet = workbook.add_worksheet()
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)


# sheet.cell_value(0,0)
year_serial = []
eleven = 1
twelve = 1
thirteen = 1
forteen = 1
fifteen = 1
sixteen = 1
seventeen = 1
eighteen = 1

for i in range(sheet.nrows):
    year_serial.append(sheet.cell_value(i,1))

# a = 0
# b = 0 
# c = 0
# d = 0
# e = 0
# f = 0
# g = 0
# h = 0

#i know i can make the following process dynamic, i will implement soon as i get some free time

for j,a in enumerate(year_serial):
    if( year_serial[j] == 2011 ):
        year_serial[j] = '11'+str(eleven)
        eleven+=1
    elif( year_serial[j] == 2012 ):
        year_serial[j] = '12'+str(twelve)
        twelve+=1
    elif( year_serial[j] == 2013 ):
        year_serial[j] = '13'+str(thirteen)
        thirteen+=1
    elif( year_serial[j] == 2014 ):
        year_serial[j] = '14'+str(forteen)
        forteen+=1
    elif( year_serial[j] == 2015 ):
        year_serial[j] = '15'+str(fifteen)
        fifteen+=1
    elif( year_serial[j] == 2016 ):
        year_serial[j] = '16'+str(sixteen)
        sixteen+=1
    elif( year_serial[j] == 2017 ):
        year_serial[j] = '17'+str(seventeen)
        seventeen+=1
    elif( year_serial[j] == 2018 ):
        year_serial[j] = '18'+str(eighteen)
        eighteen+=1
    else:
        print("Invalid")
    
    # year_serial[j] = '161'+str(a) 
# print (year_serial)

designation = []

for i in range(sheet.nrows):
    designation.append(sheet.cell_value(i,2))

for i,a in enumerate(designation):
    if (designation[i] == 'Senior Member'):
        designation[i] = 'SM'
    elif (designation[i] == 'Executive Body'):
        designation[i] = 'EB'
    elif (designation[i] == 'General Member'):
        designation[i] = 'GM'
    elif (designation[i] == 'Probationary Member'):
        designation[i] = 'PM'
    elif (designation[i] == 'Campus Compass'):
        designation[i] = 'CC'
    else:
        print("Invalid")
# print(designation)

city = []

for i in range(sheet.nrows):
    city.append(sheet.cell_value(i,3))

for i,a in enumerate(city):
    if(city[i] == 'Dhaka'):
        city[i] = 'D'
    elif(city[i] == 'Chittagong'):
        city[i] = 'C'
    else:
        print("Invalid")
# print(city)

output = []

for i in range(len(year_serial)):
    listadd = (city[i]+designation[i]+year_serial[i])
    output.append(listadd)

# print(output)







row = 0
column = 0

for c in output:
     worksheet.write(row, column, c)
     row+=1

workbook.close()