'''
In this version, they decided to change the id format to [YEAR THEY JOINED - SERIAL (according to the year the joined) - BIRTHDATE - BIRTHMONTH) . Now this required me to 
to work with the given birthdays list. This is pretty simple thanks to pandas. I easily extracted the dates and months 

'''
import xlrd, xlsxwriter, pandas

#input file
loc = ("./practice.xlsx")

#output file
workbook = xlsxwriter.Workbook("output.xlsx")
worksheet = workbook.add_worksheet()

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

dates = []

months = []

for i in range(sheet.nrows):
    dates.append(sheet.cell_value(i,2))

days = []

for i in dates:
    date = pandas.to_datetime(i).date() #removes any unnecessary data and extracts the date
    day = date.day #extracts the day
    days.append(str(day).zfill(2)) #zfill(2) here adds a zero before any single digit so the number of digits in the id card are the same
    

for i in dates:
     date = pandas.to_datetime(i).date()
     month = date.month
     months.append(str(month).zfill(2))

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

#i know i can make the following process dynamic, i will implement soon as i get some free time

for j,a in enumerate(year_serial):
    if( year_serial[j] == 2011 ):
        year_serial[j] = '11'+str(eleven).zfill(2)
        eleven+=1
    elif( year_serial[j] == 2012 ):
        year_serial[j] = '12'+str(twelve).zfill(2)
        twelve+=1
    elif( year_serial[j] == 2013 ):
        year_serial[j] = '13'+str(thirteen).zfill(2)
        thirteen+=1
    elif( year_serial[j] == 2014 ):
        year_serial[j] = '14'+str(forteen).zfill(2)
        forteen+=1
    elif( year_serial[j] == 2015 ):
        year_serial[j] = '15'+str(fifteen).zfill(2)
        fifteen+=1
    elif( year_serial[j] == 2016 ):
        year_serial[j] = '16'+str(sixteen).zfill(2)
        sixteen+=1
    elif( year_serial[j] == 2017 ):
        year_serial[j] = '17'+str(seventeen).zfill(2)
        seventeen+=1
    elif( year_serial[j] == 2018 ):
        year_serial[j] = '18'+str(eighteen).zfill(2)
        eighteen+=1
    else:
        print("Invalid")

output = []

for i in range(len(year_serial)):
    listadd = (year_serial[i]+days[i]+months[i])
    output.append(listadd)
    

row = 0
column = 0

for c in output:
     worksheet.write(row, column, c)
     row+=1

workbook.close()
