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

    endRow = 100
    endCol = 34


    rowCount = startRow

    for value in ws.iter_rows(min_row=startRow,min_col=startCol,max_row=endRow,max_col=endCol,values_only=True):
        #left off here
        #
        Filename         = value[0]
        Title            = value[1]
        Description      = value[2]
        WazeeCategory    = value[3]
        Format           = value[4]
        SubtitleLanguage = value[5]
        AudioLanguage    = value[6]
        ReleaseDate      = value[7]
        Season           = value[8]
        Episode          = value[9]
        EpisodeName      = value[10]
        DevelopmentType  = value[11]
        Genre            = value[12]
        SubGenre         = value[13]
        Host             = value[14]
        Panel            = value[15]
        Actors           = value[16]
        Guests           = value[17]
        Models           = value[18]
        Contestants      = value[19]
        Announcer        = value[20]
        Narrator         = value[21]
        ProgramVersion   = value[22]
        Color            = value[23]
        Keywords         = value[24]
        SCCFilename      = value[25]
        BoxcomURL        = value[26]
        PosterFrameTimeStart = value[27]
        SccOffset        = value[28]
        HouseNumber      = value[29]
        BuzzrID          = value[30]
        DMHScope         = value[31]
        VTR              = value[32]
        print(rowCount,Keywords)
        rowCount += 1
        #######################

xlinf = getInputFile()
workbook  = openpyxl.load_workbook(filename=xlinf) #set the load_workbook
worksheet = workbook["1. Master Metadata"]         #set the name of the worksheet to read (later have user enter)

readWS(worksheet)
