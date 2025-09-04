#!/usr/bin/env python3
import os
import sys
import numpy as np
import pandas as pd
from typing import Union
from typing import List

#This program is used to get and return all sheets within a given excel spreadsheet

#Define the data structure used to hold each row within each sheet
class inputSheetRow:
    def __init__(self, val: list):
        self.val = val

#Define the data structure used for handling each sheet within the input spreadsheet
class inputSheet:
    def __init__(self, name: str, label: List[str], row: List[inputSheetRow]):
        self.name = name
        self.label = label
        self.row = row





#Define a function that opens the specified excel sheet and returns a list of inputSheet objects, one for each sheet.
def xlsxReader(inputPath: str):
    
    #Open the excel file and read it into RAM
    inputFile = pd.ExcelFile(inputPath)

    #Find the names of the sheets we want to read
    sheetNames = inputFile.sheet_names #Create a list of all the sheet names present in the spreadsheet

    #If there's nothing in sheetNames, we have a problem we can't fix.
    if(0==len(sheetNames)):
        class nothingToOperateOn(Exception):
            pass
        raise nothingToOperateOn("No sheet names found!!!")
    

    #Create a list of inputSheet objects, one object for each element in sheetNames
    sheetsToReturn = [] #Create a list to hold our sheets
    for i in range(len(sheetNames)):
        currSheet = inputFile.parse(sheetNames[i]) #Read the current sheet
        
        #Turn each row into a inputSheetRow object and append that object to a list.
        currSheetRows = []
        for j in range(len(currSheet)):
            currSheetRows.append(inputSheetRow(currSheet.values[j]))#Read the row and append it to the list of rows

        sheetsToReturn.append(inputSheet(sheetNames[i],currSheet.columns.tolist(),currSheetRows))


        pass
    return sheetsToReturn

if __name__ == "__main__":
    if(1 == len(sys.argv)):
        readSheet = xlsxReader('Register Maps.xlsx')
    else:
        readSheet = xlsxReader(sys.argv[1])
    print("Read "+str(len(readSheet))+" sheets!")