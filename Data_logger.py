import pandas as pd
from openpyxl import load_workbook
import Methods as Mf
import sys

print ("Starting....")
filesRead=0
# Machine Menu and Select File to read---------------------------------------------------
choice = (input("Chose a device to log readings from: \n1. CT Analyzer.\n2. Votano.\n"))
if (int(choice=="exit")|(int(choice)!=1|int(choice)!=2|int(choice)!=3)):
    print ("Incorect selection, please try again!")
    sys.exit()
    
filePath=input("Enter file Prefix: ")
if (filePath=="exit"):
    sys.exit()
for time in ["M","L","A"]:
    for mes in range(1, 4):
        dfr = Mf.FiletoRead(Mf.switch(int(choice)), filePath, time, mes, filesRead)
        filePath=dfr[1]
        if (dfr[0].empty):
            print ("No "+time+str(mes)+" file")
        else:
# Select Values to be read and import them-----------------------------------------------
            DataVals = Mf.ValuestoWrite(int(choice), dfr[0])
            df2w = pd.DataFrame(DataVals[0])

# Load document to write to and calculate current row number
            book = load_workbook(DataVals[1])
            count = Mf.filelLength(book)

# Write values to new document
            writer = pd.ExcelWriter(DataVals[1], engine='openpyxl') 
            writer.book = book
            writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
            df2w.transpose().to_excel(writer, "Data",  index=False, header=False, startcol=4, startrow=count)
            writer.save()
            filesRead=filesRead+1

print("Complete!!!")
#To do:
#Automatic counting
#Readme file