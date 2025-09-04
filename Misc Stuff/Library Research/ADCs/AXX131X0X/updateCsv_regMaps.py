#!/usr/bin/env python3
import os
import pandas as pd
from typing import Union
from typing import List
import math
import csv
import shutil
from lib_getExcelSheets import xlsxReader

#Define the file path for the input (.xlsx) file and the output (.sqlite) file
spreadsheetPath = 'Register Maps.xlsx'
outputPath = './regMapCSV/'


#read the input sheet
xlsxSheets = xlsxReader(spreadsheetPath)

#Trim the sheet by collumns to get rid of the parts there for user validation
numbOfCollumns = 6 #Keep the first _ collumns within each sheet
for i in range(len(xlsxSheets)): #Within each sheet...
    for j in range(len(xlsxSheets[i].row)): #Within each row...
        xlsxSheets[i].row[j].val = xlsxSheets[i].row[j].val[:numbOfCollumns] #Keep only the first numbOfCollumns collums
    xlsxSheets[i].label = xlsxSheets[i].label[:numbOfCollumns] #Do the same for the labels


#The collumns are in the following format:
#Address,Register,NAME,Reset_Value,Start_Bit,Length
# (hex) , (text),(text),(bin/hex) , (dec) ,   (dec)

#Create a data structure to hold each register entry
class regEntry:
    def __init__(self,address,register,name,reset_value,start_bit,length):
        self.address = address
        self.register = register
        self.name = name
        self.reset_value = reset_value
        self.start_bit = start_bit
        self.length = length

#Create a data structure to hold the registers in each sheet
class regEntries:
    def __init__(self,entry: List[regEntry]):
        self.entry = entry

#Create a function to calculate the length of either a binary (0bX...) or hex (0xX...) string in bits
def binHexlength(inStr: str):
    #Determine if the string is a hex string or a binary string
    if(inStr[:2]=='0b'):
        return len(inStr)-2 #It's a binary string. The length in bits is the length - 2
    else:
        return (len(inStr)-2)*4 #It's a hex string. The length in bits is the length-2 times 4 bits per hex char


#Validate the contents for each sheet
for sheet in xlsxSheets:
    
    #Take the sheet and break it into registers
    sheetFields = []
    for row in sheet.row:#Take each row, convert it to a regEntry object, and append it to sheetFields
        sheetFields.append(regEntry(row.val[0],row.val[1],row.val[2],row.val[3],row.val[4],row.val[5]))
    sheetRegs = []
    for field in sheetFields: #For each field, put it into it's register
        #Extract the address for the current entry and convert it to a uint so we can use it as the index for sheetRegs
        currAddr = int(field.address,16)
        try:
            sheetRegs[currAddr].entry.append(field) #Add the field to the needed regEntries object if the needed regEntries object exists
        except IndexError:
            #The regEntries object for the needed register doesn't exist
            sheetRegs.insert(currAddr,regEntries([field]))
    
    del sheetFields #Free up this memory now that the fields are sorted by register

    
    #Validate the contents of each register
    for register in sheetRegs:
        #First, check if the total number of bits is correct
        bitsPresent = 0
        for i in range(len(register.entry)):
            bitsPresent = bitsPresent + register.entry[i].length
        if(bitsPresent != 16):
            class improperNumbOfBits(Exception):
                pass
            raise improperNumbOfBits("Found that the number of bits in the current register was NOT 16 bits!!!")
        del bitsPresent

        #Next, check if the number of bits in the entry.length matches the number of bits in the reset value
        for i in range(len(register.entry)):
            if(register.entry[i].length != binHexlength(register.entry[i].reset_value)):
                class improperNumbOfBits(Exception):
                    pass
                raise improperNumbOfBits("Found that the number of bits in the current field did NOT match the specified field length!!!")
        
        #Next, check to make sure that the start index of each entry is correct
        runningSum = 0
        for i in range(len(register.entry)):
            if(register.entry[i].start_bit != runningSum):
                class improperStartIndex(Exception):
                    pass
                raise improperStartIndex("Found that the start index was incorrect!!!")
            else:
                runningSum = runningSum + register.entry[i].length


        

        




#Clear the contents of the output folder
try:
    shutil.rmtree(outputPath)
except OSError as e:
    print(f"Error: {e}")
os.makedirs(outputPath)



#Now we can write the output files!
#Create one .csv for each table
for i in range(len(xlsxSheets)):
    currOutFile = outputPath + xlsxSheets[i].name + ".csv"
    with open(currOutFile, 'w', newline='', encoding='utf-8') as file:
        csv_writer = csv.writer(file)
        csv_writer.writerow(xlsxSheets[i].label)
        for j in range(len(xlsxSheets[i].row)):
            csv_writer.writerow(xlsxSheets[i].row[j].val)
    print("Wrote file "+currOutFile+" successfully")