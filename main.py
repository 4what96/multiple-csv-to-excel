import csv
import glob
import xlsxwriter

#install XlsxWriter: https://xlsxwriter.readthedocs.io/getting_started.html


array_data = []
array_fileNames = glob.glob("csv/*.csv")


#declare excel properties
workbookName = "Out.xlsx"

headers = ["header1","header2","header3","header4"]
number_of_columns = len(headers)


for fileName in array_fileNames:
    with open(fileName, newline='') as csvfile:
        raw_data = list(csv.reader(csvfile, delimiter='\t'))
        temp_row = []

        for i in range(1, number_of_columns + 1):
            temp_row.append(raw_data[i][1])

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
