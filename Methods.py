# -*- coding: utf-8 -*-
"""
Created on Mon Mar 11 16:09:28 2019

@author: Bjorn Dierks
"""
import pandas as pd
import sys

def switch(choice):
    if (choice==1):
        switcher = ["C:\\Users\Bjorn\Desktop\Bjorns work at Alectrix\CT Analyser\\","IEC_60044_1_M_Single"]
    elif (choice==2):
        switcher = ["C:\\Users\Bjorn\Desktop\Bjorns work at Alectrix\VOTANO\\","Data"]
    elif (choice==3):
        switcher = ["C:\\Users\Bjorn\Desktop\Bjorns work at Alectrix\Data logger\\","IEC_60044_1_M_Single"]
    return switcher
    

def FiletoRead(FL, inpt, time, mes, runNum):
    fp=False
    FR=True
    while (fp==False):
        try:    
            fileRead = FL[0]+"Saved Measurements\\"+inpt+time+str(mes)+".xlsx"
            dfr = pd.read_excel(fileRead, sheet_name=FL[1])
            fp=True
            FR=True
        except:
            if (runNum==0):
                print ("File could not be opened, please try again or type \"exit\"")
                inpt=input("Enter file Prefix: ")
                if (inpt=="exit"):
                    sys.exit()
                fp=False
                FR=False
            else:
                dfr=pd.DataFrame()
                fp=True
                FR=False
        if (FR==True):
            print ("Reading:",fileRead) 
    return [dfr,inpt]

def ValueLoc(choice, dfr):
    if (choice==1):
        vals=[dfr.iloc[67,25], dfr.iloc[69,25], dfr.iloc[78,25], dfr.iloc[80,25]]
    elif (choice==2):
        vals=[dfr.iloc[204,9]*100, dfr.iloc[214,9]]
    elif (choice==3):
        vals=[dfr.iloc[67,25], dfr.iloc[69,25], dfr.iloc[78,25], dfr.iloc[80,25]]
    return vals


def ValuestoWrite(choice, dfr):
    path=switch(choice)[0]
    vals=ValueLoc(choice, dfr)
    if (choice==1):
        switcher = [vals, path + "CT Analyser Data.xlsx"]
    elif (choice==2):
        switcher = [vals, path + "VOTANO Data.xlsx"]
    elif (choice==3):
        switcher = [vals, path + "CT Analyser Data.xlsx"]
    return switcher


def filelLength(book):
    count=0
    for row in book['Data']:
        if (row[4].value):
            count=count+1
    return count