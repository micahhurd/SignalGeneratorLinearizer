# Signal Generator Linearizer
# By Micah Hurd
programName = "Signal Generator Linearizer"
version = 1

exec(open("C:\\Users\\Micah\\PycharmProjects\\Libraries\\readConfigFile.py").read())
exec(open("C:\\Users\\Micah\\PycharmProjects\\Libraries\\timerFunction.py").read())
exec(open("C:\\Users\\Micah\\PycharmProjects\\Libraries\\spinTimerFunction.py").read())
exec(open("C:\\Users\\Micah\\PycharmProjects\\Libraries\\userPrompt.py").read())
exec(open("C:\\Users\\Micah\\PycharmProjects\\Libraries\\guiTools.py").read())

from os import system, name
import time
import statistics
import math
import pyvisa as visa
import datetime
import csv
import os
from pathlib import Path
import shutil
import ctypes

# from win32com import client
# from PyPDF2 import PdfFileMerger
# import xlwings as xw
# from distutils.dir_util import copy_tree
# import ast
# from time import sleep
# import numpy as np
# import pandas as pd
# import random
# from pathlib import Path
# from pandas import ExcelWriter
# from pandas import ExcelFile
# import decimal

def file_check(file):
    if not os.path.exists(file):
        return False
    else:
        return True

def clear():                                                # Clears the console
    # for windows
    if name == 'nt':
        _ = system('cls')

        # for mac and linux(here, os.name is 'posix')
    else:
        _ = system('clear')

def GPIB_resource(deviceName):
    global rm
    global list
    global resources
    global inst

    rm = visa.ResourceManager()
    list = rm.list_resources()  # Place the resources into tuple "List"
    resources = []  # Prep for the tuple to be converted "resources" from "list"

    for i, a in enumerate(list):  # Enumerate through the tuple "list"
        resources.append(a)  # Place the current enumeration into the indexed spot of "resources"

    print("Detected VISA resources")
    print("=========================================================")

    x = 0
    for i in resources:  # Enumerate through list "resources"
        print('Resource ' + str(x) + ': ' + i)  # Print each resource contained in list "resources"
        x += 1  # X is used to itemize and index the list that gets printed
    print("Resource " + str(x) + ': The GPIB address of my instrument was not auto-detected')
    print("=========================================================")

    address = -1
    while (address < 0) or (address > (x)):
        try:
            address = int(input(
                " > Enter the resource number of the " + str(deviceName) + " (0 through " + str(x) + "): "))
        except ValueError:
            print("No valid integer entered! Please try again ...")
        if address < 0 or address > (x):
            print("Available resources are 0 through " + str(x) + ". Please try again.")

    if address == x:
        address = GPIB_User_Entry(deviceName)
        inst = rm.open_resource(address)
        del inst.timeout
        return inst
    else:
        inst = rm.open_resource(resources[address])
        del inst.timeout
        return inst

def GPIB_User_Entry(deviceName):
    global rm
    global list
    global resources
    global inst

    address = -1
    while (address < 1) or (address > 32):
        try:
            address = int(input(
                " > Enter the GPIB address of the " + str(deviceName) + " (1 through 32): "))
        except ValueError:
            print("No valid integer entered! Please try again ...")
        if address < 1 or address > 32:
            print("Available GPIB addresses are 1 through 32. Please try again.")


    address = str(address)
    address = "GPIB0::" + address + "::INSTR"

    return address


def queryVisa(instrument, command, sFunc=""):
    command = str(command)
    try:
        msmt = (instrument.query(command))
    except:
        pass
    if msmt.startswith('\"') and msmt.endswith('\"'):
        msmt = msmt[1:-1]
    # msmt = float(msmt)
    msmt = msmt.strip()

    if sFunc != "":
        sFunc = sFunc.lower()

        if sFunc == "float":
            msmt = float(msmt)

    return msmt

def writeVisa(instrument, command):
    inst = instrument
    inst.write(command)

def readTxtFile(filename,searchTag,index, sFunc=""):
    searchTag = searchTag.lower()
    # print("Search Tag: ",searchTag)

    # Open the file
    with open(filename, "r") as filestream:
        # Loop through each line in the file
        for line in filestream:
            # Split each line into separated elements based upon comma delimiter
            currentLine = line.split(",")
            # select the first comma seperated value
            search = str(currentLine[0])
            # Set the value to all lowercase for comparison
            search = search.lower()

            # Remove the newline symbol from the list, if present
            lineLength = len(currentLine)
            lastElement = lineLength - 1
            if currentLine[lastElement] == "\n":
                currentLine.remove("\n")
            lineLength = len(currentLine)
            lastElement = lineLength - 1

            # Output the line which matches the search
            # If index is populated then output the indexed field
            if search == searchTag:
                if index == "":
                    output = currentLine
                    # break
                    # return currentLine

                if type(index) is int:
                    if index > lineLength:
                        output = "Index out of range"
                        # break
                    else:
                        output = currentLine[index]
                        # break
                        # return currentLine[index]

                if type(index) is str and index.find(":"):
                    index = index.split(":")
                    index[0] = int(index[0])

                    if index[0] > lineLength:
                        output = "Index out of range"
                        # break
                        # return "Index out of range"

                    if index[1] != "" and index[1] != " ":
                        index[1] = int(index[1])
                        if index[1] > lineLength:
                            output = "Index out of range"
                            # break
                            # return "Index out of range"

                    if index[1] == "" or index[1] == " ":
                        index[1] = lastElement

                    parsedLine = []
                    while index[0] <= index[1]:
                        x = currentLine[index[0]]
                        parsedLine.append(x)
                        index[0] += 1
                    output = parsedLine
                    # break
                    #return parsedLine.

                # Apply string manipulation functions, if requested (optional argument)
                if sFunc != "":
                    sFunc = sFunc.lower()

                    if sFunc == "strip":
                        output = output.strip()
                filestream.close()
                return output
        filestream.close()
        return "Searched term could not be found"


    filestream.close()
    return "Searched term could not be found"

# def userPrompt(message,inType="",range=""):
#     # requires import os
#     def stringInput(prompt):
#         a = False
#         b = False
#         while a == False:
#             string = input(" > " + str(prompt) + " ")
#             while b == False:
#                 check = input(" > Please verify that \"{}\" is correct. Enter (y)es or (n)o: ".format(string))
#                 # print("entered: ",check)
#                 check = check.lower()
#                 # print("lowered: ", check)
#                 if check != "y" and check != "n":
#                     print("!! You must enter \"y\" for Yes or \"n\" for No !!")
#                 elif check == "y" or check == "n":
#                     b = True
#             if check == "y":
#                 a = True
#             else:
#                 b = False
#         string = string.strip()
#         return string
#
#     def yesNoPrompt(prompt):
#         b = False
#         while b == False:
#             check = input(" > "+ str(prompt) + " - Enter (y)es or (n)o: ")
#             # print("entered: ",check)
#             check = check.lower()
#             # print("lowered: ", check)
#             if check != "y" and check != "n":
#                 print("!! You must enter \"y\" for Yes or \"n\" for No !!")
#             elif check == "y" or check == "n":
#                 b = True
#         if check == "y":
#             return True
#         else:
#             return False
#
#     def checkFileExists(prompt):
#         string = input(" > " + str(prompt) + " ")
#         b = os.path.isfile(string)
#         while b == False:
#             string = input(" > File \"{}\" does not exist. Please re-enter: ".format(string))
#             b = os.path.isfile(string)
#
#         string = string.strip()
#         return string
#
#     def numberInput(prompt,range):
#         range = str(range)
#         range = range.split(":")
#         numType = str(range[2])
#         if numType == "whole":
#             lower = int(range[0])
#             upper = int(range[1])
#         else:
#             lower = float(range[0])
#             upper = float(range[1])
#
#         number = lower - 1
#         while (number < lower) or (number > upper):
#             number = input(
#                 " > " + prompt + " (" + str(lower) + " through " + str(upper) + "): ")
#             try:
#                 if numType == "float":
#                     number = float(number)
#                 elif numType == "whole":
#                     number = int(number)
#             except:
#                 print("No valid number entered! Please try again...")
#                 number = lower - 1
#             else:
#                 if numType == "float" and type(number) is float:
#                     if number < lower or number > upper:
#                         print("Allowed values are " + float(lower) + " through " + float(upper) + ". Please try again.")
#                     else:
#                         return number
#                 elif numType == "whole" and type(number) is int:
#                     if number < lower or number > upper:
#                         print("Allowed values are " + str(lower) + " through " + str(upper) + ". Please try again.")
#                     else:
#                         return number
#
#     if inType != "":
#         inType = inType.lower()
#
#     if inType == "":
#         input(" > " + str(message) + "\n > Press Enter to continue... ")
#         return 0
#     elif inType == "yn":
#         return yesNoPrompt(message)
#     elif inType == "string":
#         return stringInput(message)
#     elif inType == "file":
#         return checkFileExists(message)
#     elif inType == "num":
#         return numberInput(message,range)

def create_log(log_file):

    f= open(log_file,"w+")

    f.close()
    return 0

def writeLog(analysis,logFile):
    write_mode = "a"

    currentDT = datetime.datetime.now()

    date_time = currentDT.strftime("%Y-%m-%d %H:%M:%S")

    with open(logFile, mode=write_mode, newline='') as result_file:
        result_writer = csv.writer(result_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)

        result_writer.writerow([date_time, analysis])

    result_file.close()

    return 0

def dbmToPow(dbm):
    dbm = float(dbm)
    result = (10**(dbm / 10)) / 1000
    return result

def powToDbm(pow):
    # requires import math
    pow = float(pow)
    pow = 10 * math.log10(pow * 1000)
    return pow

def percentOfDbm(percent,dbm):
    # requires import math
    dbm = float(dbm)
    percent = float(percent)
    # Convert the log power to lin power (dBm)
    linPow = (10**(dbm / 10)) / 1000

    # Calculate the percent of linear power
    percent = linPow / 100 * percent

    # Find the linear power high and low values
    high = linPow + percent
    low = linPow - percent

    # Convert the linear high and low values to log high and low values
    logHigh = 10 * math.log10(high*1000)
    logLow = 10 * math.log10(low * 1000)

    # Find the difference between the log power and the log high and low powers
    logHigh = logHigh - dbm
    logLow = dbm - logLow

    # return the greatest difference
    if logHigh > logLow:
        return logHigh
    else:
        return logLow

def write_results_excel(filename, sheet_name, result, writeRow):
    dirname = os.path.split(os.path.abspath(__file__))        # Find the current working directory

    wb = xw.Book(filename)                                          # Open the xlsm workbook
    sheet = wb.sheets[sheet_name]                                   # Open the excel input sheet for the program

    # program_output = [     0,   1,        2,   3,    4]
    # program_output = [logPow, msd, linLimit, unc, eval]

    colA = "A" + str(writeRow)
    colB = "B" + str(writeRow)
    colC = "C" + str(writeRow)
    colD = "D" + str(writeRow)
    colE = "E" + str(writeRow)

    # Insert values into the excel datasheet
    sheet.range(colA).value = result[0]
    sheet.range(colB).value = result[1]
    sheet.range(colC).value = result[2]
    sheet.range(colD).value = result[3]
    sheet.range(colE).value = result[4]

    #filename = dirname + "\\" + filename                            # Create a new filename w/full directory
                                                                    # This keeps xlwings from asking if you want to save
    # a1 = str(dirname)
    # a2 = "\\"
    # a3 = str(filename)
    # filename2 = a1 + a2 + a3
    #
    # wb.save(filename2)
    wb.save()

    return 0

def write_excel_header(filename, sheet_name, write):
    dirname = os.path.split(os.path.abspath(__file__))        # Find the current working directory

    wb = xw.Book(filename)                                          # Open the xlsm workbook
    sheet = wb.sheets[sheet_name]                                   # Open the excel input sheet for the program

    # program_output = [    0,        1,            2,         3]
    # program_output = [Model, Cal Date, Asset Number, Procedure]

    # Insert values into the excel datasheet
    sheet.range("B2").value = write[0]
    sheet.range("B3").value = write[1]
    sheet.range("B4").value = write[2]
    sheet.range("B5").value = write[3]

    #filename2 = str(dirname) + "\\" + str(filename)                            # Create a new filename w/full directory
                                                                    # This keeps xlwings from asking if you want to save
    # a1 = str(dirname)
    # a2 = "\\"
    # a3 = str(filename)
    # filename2 = a1 + a2 + a3
    # filename2 = filename2.strip()
    # wb.save(filename2)

    return 0

def create_excel_result_file(model, asset, sourceFile):

    currentDT = datetime.datetime.now()
    now = str(currentDT.strftime("%Y-%m-%d-%H%M%S"))
    now = now.strip()
    # print(now)
    # input()

    src_file = Path(sourceFile)
    src_file_wo_ext = src_file.with_suffix('')
    extension_type = src_file.suffix
    dst_file = str(src_file_wo_ext) + " - " + str(model) + " - " + str(asset) + " - " + str(now) + extension_type

    # Delete the measurement file if it already exists
    if os.path.isfile(dst_file):
        userPrompt("File \"" + str(
            dst_file) + "\" already exists; backup the original, if you wish to save it, before proceeding. Press Enter to Continue...")
        os.remove(dst_file)

    shutil.copy(src_file, dst_file)

    resultFile = dst_file

    return resultFile

def exceltopdf(inputFile,outFile,sheetName):
    success = 1
    excel = client.DispatchEx("Excel.Application")
    excel.Visible = 0

    wb = excel.Workbooks.Open(inputFile)
    ws = wb.Worksheets[sheetName]

    try:
        wb.SaveAs(outFile, FileFormat=57)
    except Exception:
        print("Failed to convert output file to PDF")
        success = 0
        #print(str(e))
    finally:
        wb.Close()
        excel.Quit()
    return success

def mergePDF(inputFiles,outFile):
    pdfs = inputFiles
    merger = PdfFileMerger()
    for pdf in pdfs:
        merger.append(pdf)

    merger.write(outFile)
    merger.close()

def fileExistCheck(file):
    # requires from pathlib import Path
    # requires import os

    dirCheck = False
    dirCheck = os.path.isfile(file)
    if dirCheck == False:
        return 0  # indicating that the file did not exist already

    n = 1
    while dirCheck == True:
        # write_analysis_log("Backup destination directory already exists! :" + toDirectory, logFile)
        n += 1

        newFile = Path(file)
        newFile = newFile.with_suffix('')
        newFile = str(newFile) + " (" + str(n) + ").pdf"
        dirCheck = os.path.isfile(newFile)

    return newFile  # return new recommended file name because the file already exists

def userInterfaceHeader(program,version,cwd,logFile,msg=""):
    print(program + ", Version " + str(version))
    print("Current Working Directory: " + str(cwd))
    print("Log file located at working directory: " + str(logFile))
    print("=======================================================================")
    if msg != "":
        print(msg)
        print("_______________________________________________________________________")
    return 0

def createValuesImage(dutGenName,stdGenName,asset,reCharacterizationInterval,resX,resY,startX,startY,lineSpacing):
    from PIL import Image, ImageDraw, ImageFont
    # import datetime
    from datetime import datetime
    from datetime import timedelta
    import os
    cwd = os.getcwd()


    # Work out current date and re-characterization date
    currentDT = datetime.now()
    date_time = currentDT.strftime("%Y-%m-%d, %H%M")
    charDate = date_time

    currentDT2 = datetime.now() + timedelta(days=reCharacterizationInterval)
    date_time2 = currentDT2.strftime("%Y-%m-%d, %H%M")
    reCharDate = date_time2
    print(str(reCharDate))

    # Setup Image Parameters
    img = Image.new('RGB', (resX, resY), color=(73, 109, 137))
    fnt = ImageFont.truetype('C:/Windows/Fonts/Arial.ttf', 50)
    d = ImageDraw.Draw(img)
    x = startX
    y = startY
    lineSpacing = lineSpacing

    d.text((x, y), "{} was characterized by {},".format(dutGenName, stdGenName), font=fnt, fill=(255, 255, 0))
    y+=lineSpacing
    d.text((x, y), "Asset Number {}.".format(asset), font=fnt, fill=(255, 255, 0))
    y += lineSpacing
    d.text((x, y), "----------------------------------------------------------".format(asset), font=fnt, fill=(255, 255, 0))
    y += lineSpacing
    d.text((x, y), "Characterization Date: {}".format(charDate), font=fnt, fill=(255, 255, 0))
    y += lineSpacing
    d.text((x, y), "Re-Characterization Due: {}".format(reCharDate), font=fnt, fill=(255, 255, 0))


    fileName = "{} - {} - {} - {}.bmp".format(dutGenName,stdGenName,asset,charDate)
    img.save(fileName)
    return fileName

logFile = "sgLinearizerlog.txt"
cwd = os.getcwd()
useTKINTER = False

if os.path.isfile(logFile):
    # print ("Log File Exists")
    logFile = logFile
else:
    create_log(logFile)

writeLog("Starting New Program Instance ---------------------------",logFile)
writeLog("Program Name: "+str(programName),logFile)
writeLog("Program Version: "+str(version),logFile)
writeLog("Current Working Directory: "+str(cwd),logFile)

clear()
userInterfaceHeader(programName,version,cwd,logFile)


# Check installed version of microsoft office
# if useTKINTER == True:
#     check = guiTools.yesNoGUI("Does this computer have Microsoft Office 2012 or newer installed?")
# else:
#     check = userPrompt("Does this computer have Microsoft Office 2012 or newer installed?", "yn")
# if check == True:
#     office2012orNewer = True
# else:
#     xcelSavePDF = False
#     office2012orNewer = False

xcelSavePDF = False
office2012orNewer = False

# if office2012orNewer == False:
#     writeLog("User selected that installed Microsoft Office version is older than 2012", logFile)
#     if useTKINTER == True:
#         guiTools.popupMsg("This program will only be able to write measurement results to Excel\nThe user must print the Excel data to PDF and merge manually.")
#     else:
#         userPrompt("This program will only be able to write measurement results to Excel\nThe user must print the Excel data to PDF and merge manually.")
# else:
#     writeLog("User selected that installed Microsoft Office version is 2012 or newer", logFile)


# check = False
# while check == False:
#     if useTKINTER == True:
#         guiTools.popupMsg("At the following file dialogue prompt, select the measurement template file.")
#         templateFile = guiTools.getFilePath("*.lin", templateDirectory, "Lin Template")
#         dutModel = readTxtFile(templateFile, "dutModel", 1, sFunc="strip")
#         writeLog("Template Selected: " + str(templateFile), logFile)
#         check = guiTools.yesNoGUI("Is {} the correct model of the DUT?".format(dutModel))
#
#     else:
#         # Load and verify the correct template file
#         print("Templates Folder: {}".format(templateDirectory))
#         print("===")
#         print("Scanning for measurement templates...")
#         print("===========================================")
#
#         x = 1
#         templateList = []
#         for file in os.listdir(templateDirectory):
#             if file.endswith(".lin"):
#                 print("Template " + str(x) + ": " + str(file))
#                 templateList.append(file)
#
#             x += 1
#
#         qtyTemplates = len(templateList)
#         qtyTemplatesStr = "1:" + str(qtyTemplates) + ":whole"
#
#         if qtyTemplates == 0:
#             print("No templates found in the template folder.")
#         print("===========================================")
#
#         if qtyTemplates == 0:
#             check = False
#             while check == False:
#                 templateFile = userPrompt("Enter the full path-filename of the measurement template:", "file")
#                 # templateFile = templateDirectory + "\\" + templateFile
#                 print("Scanning template for model number...")
#                 time.sleep(0.5)
#                 dutModel = readTxtFile(templateFile, "dutModel", 1, sFunc="strip")
#                 writeLog("Template Selected: " + str(templateFile), logFile)
#                 check = userPrompt("Is \"" + str(dutModel) + "\" the correct model of the DUT?", "yn")
#         else:
#             check = False
#             while check == False:
#                 templateFileNumber = userPrompt("Select the desired template", "num", qtyTemplatesStr)
#                 templateFileNumber -= 1
#                 templateFile = templateDirectory + "\\" + str(templateList[templateFileNumber])
#                 print("Scanning template for model number...")
#                 time.sleep(0.5)
#                 dutModel = readTxtFile(templateFile, "dutModel", 1, sFunc="strip")
#                 writeLog("Template Selected: " + str(templateFile), logFile)
#                 check = userPrompt("Is \"" + str(dutModel) + "\" the correct model of the DUT?", "yn")


configFile = "sglconfig.cfg"
# dutModel = readConfigFile(templateFile, "dutModel")




# DUT Information Routine

# # print("===========================================")
# # model = input("Enter the model number of the DUT: ")
# asset = userPrompt("Enter the asset number of the DUT:","string")
# writeLog("User entered >{}< as DUT asset number".format(str(asset)), logFile)

stdGenName = readConfigFile(configFile, "stdGen")
writeLog("stdGen: "+str(stdGenName),logFile)
dutGenName = readConfigFile(configFile, "dutGen")
writeLog("dutGen: "+str(dutGenName),logFile)

currentDT = datetime.datetime.now()
currentDT = currentDT.strftime("%Y-%m-%d %H%M")
measLog = dutGenName + " Char By " + stdGenName + " - " + currentDT + ".csv"
# Delete the CSV measurement file if it already exists
if os.path.isfile(measLog):
    if useTKINTER == True:
        guiTools.popupMsg("Measurement log file \"" + str(measLog) + "\" already exists; backup the original, if you wish to save it, before proceeding.")
    else:
        userPrompt("Measurement log file \"" + str(measLog) + "\" already exists; backup the original, if you wish to save it, before proceeding. Press Enter to Continue...")
    os.remove(measLog)
    writeLog("Measurement log file > {} < already existed - deleted old log".format(str(measLog)), logFile)
else:
    create_log(measLog)
writeLog("Created measurement log file: " + str(measLog), logFile)


dutCorFile = dutGenName + " Correction Data.csv"
# Delete the CSV measurement file if it already exists
if os.path.isfile(dutCorFile):
    if useTKINTER == True:
        guiTools.popupMsg("DUT Correction Data file \"" + str(dutCorFile) + "\" already exists; backup the original, if you wish to save it, before proceeding.")
    else:
        userPrompt("DUT Correction Data file \"" + str(dutCorFile) + "\" already exists; backup the original, if you wish to save it, before proceeding. Press Enter to Continue...")
    os.remove(dutCorFile)
    writeLog("DUT Correction Data file > {} < already existed - deleted old correction file".format(str(dutCorFile)), logFile)
else:
    create_log(dutCorFile)
writeLog("Created DUT Correction Data file: " + str(dutCorFile), logFile)

# Import values from configuration file ------------------------------
print("Importing parameters from the configuration template...")
writeLog("Importing parameters from configuration file: {}".format(str(configFile)), logFile)
time.sleep(0.5)
charFreq = readConfigFile(configFile,"characterizationFrequency")
charFreq = float(charFreq)
pMeterName = readConfigFile(configFile,"pMeter")
writeLog("pMeterName: "+str(pMeterName),logFile)
sensorMaxPow = readConfigFile(configFile,"sensorMaxPow")
sensorMaxPow = float(sensorMaxPow)
reCharacterizationInterval = readConfigFile(configFile,"reCharacterizationInterval")
reCharacterizationInterval = int(reCharacterizationInterval)
uom = readConfigFile(configFile,"uom")
settlingTime = readConfigFile(configFile,"settlingTime")
settlingTime = int(settlingTime)
samplingQuantity = readConfigFile(configFile,"samplingQuantity")
samplingQuantity = int(samplingQuantity)
dutTargetAcc = readConfigFile(configFile,"dutTargetAcc")
dutTargetAcc = float(dutTargetAcc)

resX = readConfigFile(configFile,"resX")
resX = int(resX)
resY = readConfigFile(configFile,"resY")
resY = int(resY)
startX = readConfigFile(configFile,"startX")
startX = int(startX)
startY = readConfigFile(configFile,"startY")
startY = int(startY)
lineSpacing = readConfigFile(configFile,"lineSpacing")
lineSpacing = int(lineSpacing)

stdGenID = readConfigFile(configFile,"stdGenID")
stdGenRst = readConfigFile(configFile,"stdGenRst")
stdGenFreq = readConfigFile(configFile,"stdGenFreq")
stdGenPow = readConfigFile(configFile,"stdGenPow")
stdGenConfig = readConfigFile(configFile,"stdGenConfig")
stdGenOn = readConfigFile(configFile,"stdGenOn")
stdGenOff = readConfigFile(configFile,"stdGenOff")

dutGenID = readConfigFile(configFile,"dutGenID")
dutGenRst = readConfigFile(configFile,"dutGenRst")
dutGenFreqSet = readConfigFile(configFile,"dutGenFreqSet")
dutGenPowSet = readConfigFile(configFile,"dutGenPowSet")
dutGenFreqRead = readConfigFile(configFile,"dutGenFreqRead")
dutGenPowRead = readConfigFile(configFile,"dutGenPowRead")
dutGenConfig = readConfigFile(configFile,"dutGenConfig")
dutGenOn = readConfigFile(configFile,"dutGenOn")
dutGenOff = readConfigFile(configFile,"dutGenOff")

pmID = readConfigFile(configFile,"pmID")
pmRst = readConfigFile(configFile,"pmRst")
pmZero = readConfigFile(configFile,"pmZero")
pmCal = readConfigFile(configFile,"pmCal")
pmZeroCal = readConfigFile(configFile,"pmZeroCal")
pmFreq = readConfigFile(configFile,"pmFreq")
pmConfig = readConfigFile(configFile,"pmConfig")
pmOpc = readConfigFile(configFile,"pmOpc")
pmRead = readConfigFile(configFile,"pmRead")

# Get the steps list and convert from string to float
steps = readConfigFile(configFile,"logSteps","listOut")
steps = [x.strip() for x in steps]
steps = [float(x) for x in steps]

# End Import values from configuration file ------------------------------



# if pdfMerge == "yes":
#     writeLog("Template requires a PDF merge...", logFile)
#     if office2012orNewer == True:
#         if useTKINTER == True:
#             guiTools.popupMsg(
#                 "The template file settings require that the results of this calibration be merged\nto an existing (parent) PDF File\n\nAt the following file dialogue prompt, choose the PDF file.")
#             pdf1 = guiTools.getFilePath("*.pdf", cwd, "PDF File")
#         else:
#             print(
#                 "The template file settings require that the results of this calibration be merged to an existing (parent) PDF file.")
#             pdf1 = userPrompt("Enter the full path and filename of the parent PDF file:", "file")
#         writeLog("User entered value for PDF parent file: {}".format(str(pdf1)), logFile)
#     else:
#         if useTKINTER == True:
#             guiTools.popupMsg(
#                 "The template file settings require that the results of this calibration be merged\nto an existing (parent) PDF File,\nhowever this computer does not have Office 2012 or great.\n\nThe user will have to create and merge the pdf manually")
#         else:
#             userPrompt("The template file settings require that the results of this calibration be merged\nto an existing (parent) PDF File,\nhowever this computer does not have offce 2012 or great.\n\nThe user will have to create and merge the pdf manually")
#         writeLog("Template requires PDF merge, but office version installed is less than 2012", logFile)


# Power on and connect instruments --------------------------------------------------
clear()
userInterfaceHeader(programName,version,cwd,logFile,"Required Instruments")
print("Obtain, power on, and \"GPIB-in\" the following:")
print("- Fluke 96270A")
print("- R&S SMA100B")
print("- EPM Compatable Power Meter")
print("- AT/E4412A Power Sensor")
print("")
print("Ensure the power sensor is connected to channel A of the power meter")
userPrompt("Press enter to continue...")

clear()
userInterfaceHeader(programName, version, cwd, logFile)
asset = userPrompt("Enter the asset number of the {}:".format(stdGenName),"string")
writeLog("User entered >{}< as standard generator asset number".format(str(asset)), logFile)

tempBool = userPrompt("Have all instruments stabilized with power on for at least 1 hour?","yn")
if tempBool == False:
    writeLog("User selected NO instrument are not stabilized", logFile)
    print("The instruments will now be allowed to stabilize for 1 hour")
    userPrompt("Press enter to continue...")

    clear()
    userInterfaceHeader(programName, version, cwd, logFile, "Stabilizing Instruments")
    timerFunction(3600, debug=True)
else:
    writeLog("User selected that instruments were stabilized for at least an hour", logFile)


# Load variable names with VISA resources for each instrument --------
clear()
userInterfaceHeader(programName,version,cwd,logFile)
stdGenVisa = GPIB_resource(stdGenName)
writeLog("Leveled Generator Visa Resource: {}".format(str(stdGenVisa)), logFile)

clear()
userInterfaceHeader(programName,version,cwd,logFile)
dutGenVisa = GPIB_resource(dutGenName)
writeLog("DUT Generator Visa Resource: {}".format(str(dutGenVisa)), logFile)

clear()
userInterfaceHeader(programName,version,cwd,logFile)
pMeterVisa = GPIB_resource(pMeterName)
writeLog("Power Meter Visa Resource: {}".format(str(pMeterVisa)), logFile)


clear()
userInterfaceHeader(programName,version,cwd,logFile,"VISA device response")
response = queryVisa(stdGenVisa,stdGenID)
print("Device " + str(stdGenName) + ": "+str(response))
writeLog("Leveled Generator IDN? Response: {}".format(str(response)), logFile)
response = queryVisa(dutGenVisa,dutGenID)
print("Device " + str(dutGenName) + ": "+str(response))
writeLog("DUT Generator IDN? Response: {}".format(str(response)), logFile)
response = queryVisa(pMeterVisa,pmID)
print("Device " + str(pMeterName) + ": "+str(response))
writeLog("Power Meter IDN? Response: {}".format(str(response)), logFile)
userPrompt("Press enter to continue...")
# Done loading resources ---------------------------------------------


clear()
userInterfaceHeader(programName,version,cwd,logFile,"VISA device configuration")
# Leveled Generator Config Routine --------------------------------------
print("Resetting " + str(stdGenName) + "...")
writeVisa(stdGenVisa,stdGenRst)
time.sleep(1)
print("Configuring " + str(stdGenName) + "...")
writeVisa(stdGenVisa,stdGenConfig)
writeLog("Reset and configured visa device {}".format(str(stdGenName)), logFile)
time.sleep(1)
# End Leveled Generator Config Routine --------------------------------------

# DUT Generator Config Routine --------------------------------------
print("Resetting " + str(dutGenName) + "...")
writeVisa(stdGenVisa,dutGenRst)
time.sleep(1)
print("Configuring " + str(dutGenName) + "...")
writeVisa(stdGenVisa,dutGenConfig)
writeLog("Reset and configured visa device {}".format(str(dutGenName)), logFile)
time.sleep(1)
# End DUT Generator Config Routine --------------------------------------

# PM Config Routine --------------------------------------
time.sleep(1)

print("Resetting " + str(pMeterName) + "...")
writeVisa(pMeterVisa,pmRst)
time.sleep(1)
print("Configuring " + str(pMeterName) + "...")
writeVisa(pMeterVisa,pmConfig)
writeLog("Reset and configured visa device {}".format(str(pMeterName)), logFile)

time.sleep(1)

clear()
userInterfaceHeader(programName,version,cwd,logFile,"Power Meter Initial Zero")

tempBool = userPrompt("Has the power sensor already been zeroed and calibrated?","yn")
if tempBool == False:
    writeLog("User selected to zero & cal power meter", logFile)
    print("Connect the power sensor to the power meter reference port.")
    userPrompt("Press enter to continue...")
    print("Zero and Calibrating Power Meter...")
    writeVisa(pMeterVisa,pmZeroCal)
    time.sleep(0.25)
    response = queryVisa(pMeterVisa,pmOpc)
    print("Meter Zero Cal Response: ",response)
    time.sleep(1)
else:
    writeLog("User chose not to zero and cal the power meter", logFile)

# End PM Config Routine --------------------------------------

clear()
userInterfaceHeader(programName,version,cwd,logFile,"Linearity Test Setup")

# Zeroing Routine -------------------------------------------
print("Connect the power sensor to the output of the {}.".format(stdGenName))
userPrompt("Press enter to continue...")
# writeVisa(stdGenVisa,stdGenOff)
# response = queryVisa(pMeterVisa, pmRead, "float")
# writeLog("Power meter measured power with no input power at start of linearity test setup: {}".format(str(response)), logFile)
#
# print("Setting leveled generator...")
#
# time.sleep(1)
# writeVisa(stdGenVisa,stdGenFreq.replace("<val>","50000000"))
# time.sleep(0.1)
# writeVisa(stdGenVisa,stdGenPow.replace("<val>","0"))
# time.sleep(0.1)
# writeVisa(stdGenVisa,stdGenOn)
# time.sleep(0.1)
# print("Calibrating the power meter at 0 dBm...")
# writeVisa(pMeterVisa,pmCal)
# response = queryVisa(pMeterVisa,pmOpc)
# print("Power meter responded: ",str(response))
# response = queryVisa(pMeterVisa,pmRead,"float")
# print("Power cal value: ",str(response))
# writeLog("Power meter cal value: {}".format(str(response)), logFile)
#
# time.sleep(1)
# writeVisa(stdGenVisa,stdGenFreq.replace("<val>","50000000"))
# time.sleep(0.1)
# writeVisa(stdGenVisa,stdGenPow.replace("<val>",refPow))
# time.sleep(0.1)
# writeVisa(stdGenVisa,stdGenOn)
# time.sleep(3)
# print("Calibrating the power meter to the reference power level ({} dBm)...".format(refPow))
# response = queryVisa(pMeterVisa,pmRead,"float")
# print("Power measured at reference level: ",str(response))
# writeLog("Power meter measured power at reference level: {}".format(str(response)), logFile)
#
# print("Measuring reference bias...")
# # Calculate an offset based upon the amount of residual zero power
# msmts = []
# x = 30
# while x != 0:
#     time.sleep(1)
#     print("{} seconds remaining...\r".format(x), end="")
#     x -= 1
#
# for i in range(30):
#     time.sleep(0.1)
#     response = queryVisa(pMeterVisa, pmRead, "float")
#     msmts.append(response)
#
# # print(msmts)
# average = statistics.mean(msmts)
# # print("resulting average value: " + str(average))
# offset = average - float(refPow)
# offsetSdev = statistics.stdev(msmts)
# print("Bias calculated to be: " + str(offset))
# writeLog("Power meter calculated bias: {}".format(str(offset)), logFile)
# print("Bias SDEV calculated to be: " + str(offsetSdev))
# # End Zeroing Routine -------------------------------------------
#
# # Create the excel results file from the template source file
# writeLog("Excel source file: {}".format(str(excelSource)), logFile)
# excelMsmtFile = create_excel_result_file(dutModel, asset, excelSource)
# writeLog("Excel measurement file: {}".format(str(excelMsmtFile)), logFile)
#
#
# # write the header info to the excel results file
# toExcel = [dutModel, currentDT, asset, str(programName) + " Version " + str(version)]
# write_excel_header(excelMsmtFile, msmtFileSheet, toExcel)
# writeLog("Wrote header info to the excel measurement file", logFile)


# Measurement routine
clear()
userInterfaceHeader(programName,version,cwd,logFile,"Performing Linearity Test")
writeLog("Started linearity test", logFile)

time.sleep(1)
writeVisa(stdGenVisa,stdGenFreq.replace("<val>", str(charFreq)))
writeVisa(stdGenVisa,stdGenOn)
writeVisa(pMeterVisa,pmFreq.replace("<val>", str(charFreq)))

# Loop to use standard generator to characterize the sensor
stdPowVals = []
stdLogVals = []
for index, i in enumerate(steps):
    # Break out test points and tolerances
    # linLimit = float(tol[x])                                # receive the limit as a percentage
    logPow = float(i)                                       # receive the test point power as log (dBm)
    linPow = dbmToPow(logPow)                               # obtain the test point power in linear units (Watts)
    # logLimit = percentOfDbm(linLimit,logPow)                # convert the linear limits to log (+/-) dB limits
    # These are the logarithmic upper and lower bounds in dBm
    # high = logPow + logLimit
    # low = logPow - logLimit

    # Stabilize the generator and sensor...
    writeVisa(stdGenVisa, stdGenPow.replace("<val>", str(i)))
    clear()
    userInterfaceHeader(programName, version, cwd, logFile, "Transfer Characterization In Process")
    spinTimerFunction(settlingTime, "Stabilizing Output Level at {} dBm...".format(i))

    clear()
    userInterfaceHeader(programName, version, cwd, logFile, "Transfer Characterization In Process")
    spinner = 0
    msmts = []
    for j in range(samplingQuantity):
        if spinner == 4:
            spinner = 0
        if spinner == 0:
            progress = "|"
        elif spinner == 1:
            progress = "/"
        elif spinner == 2:
            progress = "-"
        elif spinner == 3:
            progress = "\\"
        print('Measuring... {}\r'.format(progress), end="")
        time.sleep(0.25)
        response = queryVisa(pMeterVisa, pmRead, "float")
        msmts.append(response)
        spinner += 1
    msd = statistics.mean(msmts)
    stdLogVals.append(msd)                  # Save log value measured to list
    linMsd = dbmToPow(msd)
    stdPowVals.append(linMsd)               # Save linear value measured to list

    powAcc = (linMsd / linPow * 100) - 100

    absPowAcc = abs(powAcc)

    if samplingQuantity > 1:
        msdSdev = statistics.stdev(msmts)

    unc = "- -"

    toLog = "{}- ,{}, dBm Step,{}, dBm- Measured ,{}, dBm (,{}, Watt Step ,{}, Watt)".format(stdGenName,logPow,logPow,msd,linPow,linMsd)

    writeLog(toLog, measLog)

writeVisa(stdGenVisa,stdGenRst)
writeLog("Completed standard linearity characterization loop", logFile)


print("Perform the following:")
print("- Disconnect the power sensor from the output of the {}.".format(stdGenName))
print("- Connect the power sensor to the output of the {}.".format(dutGenName))
userPrompt("Press enter to continue...")


writeVisa(dutGenVisa,dutGenFreqSet.replace("<val>", str(charFreq)))
writeVisa(dutGenVisa,dutGenOn)

# Loop to use sensor to characterize the DUT generator
dutPowVals = []
dutLogVals = []
dutInitialLogVals = []
dutInitialPowVals = []
for index, i in enumerate(steps):
    # Break out test points and tolerances
    # linLimit = float(tol[x])                                # receive the limit as a percentage
    logPow = float(i)                                       # receive the test point power as log (dBm)
    linPow = dbmToPow(logPow)                               # obtain the test point power in linear units (Watts)
    # logLimit = percentOfDbm(linLimit,logPow)                # convert the linear limits to log (+/-) dB limits
    # These are the logarithmic upper and lower bounds in dBm
    # high = logPow + logLimit
    # low = logPow - logLimit
    clear()
    # Stabilize the generator and sensor...
    writeVisa(dutGenVisa,dutGenPowSet.replace("<val>", str(logPow)))
    userInterfaceHeader(programName, version, cwd, logFile, "DUT Characterization In Process")
    spinTimerFunction(settlingTime,"Stabilizing Output Level at {} dBm...".format(i))

    clear()
    userInterfaceHeader(programName, version, cwd, logFile, "DUT Characterization In Process")
    spinner = 0
    msmts = []
    for j in range(samplingQuantity):
        if spinner == 4:
            spinner = 0
        if spinner == 0:
            progress = "|"
        elif spinner == 1:
            progress = "/"
        elif spinner == 2:
            progress = "-"
        elif spinner == 3:
            progress = "\\"
        print('Measuring Output Level at {} dBm... {}\r'.format(i,progress), end="")
        time.sleep(0.25)
        response = queryVisa(pMeterVisa, pmRead, "float")
        msmts.append(response)
        spinner += 1
    msd = statistics.mean(msmts)
    dutInitialLogVals.append(msd)                  # Save log value measured to list
    linMsd = dbmToPow(msd)
    dutInitialPowVals.append(linMsd)               # Save linear value measured to list
    powAcc = (linMsd / linPow * 100) - 100
    absPowAcc = abs(powAcc)

    # Loop three time to get the dut generator output as close to the sensor indicated nominal as possible
    for j in range(5):

        tempDiff = stdLogVals[index] - msd

        # Only adjust power if the difference was greater than dutTargetAcc
        if abs(tempDiff) > dutTargetAcc:
            logPow = logPow + tempDiff

            writeVisa(dutGenVisa, dutGenPowSet.replace("<val>", str(logPow)))
            spinTimerFunction(settlingTime,"Stabilizing Output Level at {} dBm...".format(i))
            clear()

            spinner = 0
            msmts = []
            for j in range(samplingQuantity):
                if spinner == 4:
                    spinner = 0
                if spinner == 0:
                    progress = "|"
                elif spinner == 1:
                    progress = "/"
                elif spinner == 2:
                    progress = "-"
                elif spinner == 3:
                    progress = "\\"
                print('Measuring Output Level at {} dBm... {}\r'.format(i,progress), end="")
                time.sleep(0.25)
                response = queryVisa(pMeterVisa, pmRead, "float")
                msmts.append(response)
                spinner += 1
            msd = statistics.mean(msmts)

    dutLogVals.append(msd)  # Save log value measured to list
    linMsd = dbmToPow(msd)
    dutPowVals.append(linMsd)  # Save linear value measured to list
    powAcc = (linMsd / linPow * 100) - 100
    absPowAcc = abs(powAcc)
    dutGenSetLogVal = logPow
    logPow = i


    if samplingQuantity > 1:
        msdSdev = statistics.stdev(msmts)

    unc = "- -"

    toLog = "{}- ,{}, dBm Step- DUT Set ,{}, dBm- Measured ,{}, dBm".format(dutGenName,logPow,dutGenSetLogVal,msd)
    writeLog(toLog, measLog)

    toFile = "{},{},{},{},{},{}".format(dutGenName, charFreq, logPow, dutGenSetLogVal, msd, asset)
    writeLog(toFile, dutCorFile)


writeVisa(dutGenVisa,dutGenRst)
writeLog("Completed DUT linearity characterization loop", logFile)

writeLog("Creating image file...", logFile)
fileName = createValuesImage(dutGenName,stdGenName,asset,reCharacterizationInterval,resX,resY,startX,startY,lineSpacing)
imagePath = cwd + "\\" + fileName
imagePath = os.path.normpath(imagePath)
writeLog("Image created at: {}".format(imagePath), logFile)

try:
    SPI_SET_WALLPAPER = 20
    ctypes.windll.user32.SystemParametersInfoW(SPI_SET_WALLPAPER, 0, imagePath, 0)
    writeLog("Set desktop successfully", logFile)
except:
    writeLog("Could not set desktop background", logFile)
    print("Could not set desktop background...")
    print("Background image at: {}".format(imagePath))
    print("Contact QA for assistance")
    input("Press any key to exit program...")
    sys.exit()

print("Characterization completed; this window will auto-close in 5 seconds")
time.sleep(5)

writeLog("Program Terminated --------------------------------------", logFile)
sys.exit()


# Data Combination Routine



# Update Ideas
# Write code to limit the output power of the generators to protect the sensor
# Add checks to verify the power sensor is properly connected
