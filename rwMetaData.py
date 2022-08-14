#! /usr/bin/python3

import os
import re
import sys
import warnings

warnings.filterwarnings("ignore")
sys.path.append('/Users/ronperez/Desktop/Stuff/Programming/python/modules')

import openpyxl


#get input file from user
def getInputFile():

    while True:
        inf = input("Give me your input file: ")
        if os.path.exists(inf):
            break
        else:
            print("   Does not exist")

    return inf

def readWS(ws):

    #set the start of where to read from input xlSheet
    startRow = 5
    startCol = 2

    endRow = 5
    endCol = 34

    for value in ws.iter_rows(min_row=startRow,min_col=startCol,max_row=endRow,max_col=endCol,values_only=True):
        print(value)

xlinf = getInputFile()
workbook  = openpyxl.load_workbook(filename=xlinf) #set the load_workbook
worksheet = workbook["1. Master Metadata"]         #set the name of the worksheet to read (later have user enter)

readWS(worksheet)
