import csv
import glob
import xlsxwriter

#install XlsxWriter https://xlsxwriter.readthedocs.io/getting_started.html
# "pip install XlsxWriter"

#print(glob.glob("csv/*.csv"))

#fileName = 'test.csv'
array_data = []
#array_fileNames = ['test.csv','test2.csv']
array_fileNames = glob.glob("csv/*.csv")


#declare excel properties
workbookName = "Out.xlsx"

headers = ["header1","header2","header3","header4"]
number_of_columns = len(headers)


for fileName in array_fileNames:
    with open(fileName, 'r') as csvfile:
        reader = csv.reader(csvfile, delimiter = ';')
        temp_row = []
        index = 0
        for rows in reader:
            for numberRows in range(0,number_of_columns):
                if(index == 1):
                    temp_row.append(rows[numberRows])
            #print(rows)
            index += 1
        #print(reader)

        for element in range(len(temp_row)):
            temp_row[element] = temp_row[element].replace(".",",")

        temp_row.append(fileName.replace("csv\\",""))
        array_data.append(temp_row)

outWorkbook = xlsxwriter.Workbook(workbookName)
outSheet = outWorkbook.add_worksheet()

headers.append("File Name")
#write data to file
for item in range(len(headers)):
    outSheet.write(0,item, headers[item])

for row in range(len(array_data)):
    for column in range(len(array_data[0])):
        outSheet.write(row+1,column,array_data[row][column])

outWorkbook.close()
