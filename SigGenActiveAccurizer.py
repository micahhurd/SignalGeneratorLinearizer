# Power Sensor Linearity Program
# By Micah Hurd
programName = "Signal Generator Active Accurizer"
version = 1

from os import system, name
import time
import datetime
import csv
import os

def readConfigFile(filename, searchTag, sFunc=""):
    searchTag = searchTag.lower()
    # print("Search Tag: ",searchTag)

    # Open the file
    with open(filename, "r") as filestream:
        # Loop through each line in the file

        for line in filestream:

            if line[0] != "#":

                currentLine = line
                equalIndex = currentLine.find('=')
                if equalIndex != -1:

                    tempLength = len(currentLine)
                    # print("{} {}".format(equalIndex,tempLength))
                    tempIndex = equalIndex
                    configTag = currentLine[0:(equalIndex)]
                    configTag = configTag.lower()
                    configTag = configTag.strip()
                    # print(configTag)

                    configField = currentLine[(equalIndex + 1):]
                    configField = configField.strip()
                    # print(configField)

                    # print("{} {}".format(configTag,searchTag))
                    if configTag == searchTag:

                        # Split each line into separated elements based upon comma delimiter
                        # configField = configField.split(",")

                        # Remove the newline symbol from the list, if present
                        lineLength = len(configField)
                        lastElement = lineLength - 1
                        if configField[lastElement] == "\n":
                            configField.remove("\n")
                        # Remove the final comma in the list, if present
                        lineLength = len(configField)
                        lastElement = lineLength - 1

                        if configField[lastElement] == ",":
                            configField = configField[0:lastElement]

                        lineLength = len(configField)
                        lastElement = lineLength - 1

                        # Apply string manipulation functions, if requested (optional argument)
                        if sFunc != "":
                            sFunc = sFunc.lower()

                            if sFunc == "listout":
                                configField = configField.split(",")

                        filestream.close()
                        return configField

        filestream.close()
        return "Searched term could not be found"

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

def GPIB_resource(deviceName,searchString=""):
    import pyvisa as visa
    global rm
    global list
    global resources
    global inst

    searchString = searchString.lower()

    rm = visa.ResourceManager()
    list = rm.list_resources()  # Place the resources into tuple "List"
    resources = []  # Prep for the tuple to be converted "resources" from "list"

    for i, a in enumerate(list):  # Enumerate through the tuple "list"
        resources.append(a)  # Place the current enumeration into the indexed spot of "resources"

    if searchString != "":
        inst = ""
        x = 0
        for index, i in enumerate(resources):  # Enumerate through list "resources"
            currentString = i.lower()
            tempSearch = currentString.find(searchString)
            # print("{} - {} - {}".format(i, searchString, tempSearch))

            if tempSearch == 0:
                inst = rm.open_resource(resources[index])
        if inst == "":
            print("Could not find GPIB resource containing: {}".format(searchString))
            return "Could not find GPIB resource containing: {}".format(searchString)
    else:
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

def userInterfaceHeader(program,version,cwd,logFile,msg=""):
    print(program + ", Version " + str(version))
    print("Current Working Directory: " + str(cwd))
    print("Log file located at working directory: " + str(logFile))
    print("=======================================================================")
    if msg != "":
        print(msg)
        print("_______________________________________________________________________")
    return 0

def readCorrectionFile(filename):
    charFreqList = []
    setValList = []
    correctedValList = []
    # Open the file
    with open(filename, "r") as filestream:
        # Loop through each line in the file

        for line in filestream:
            lineList = []
            if line[0] != "#":

                currentLine = line
                currentLine = currentLine.strip()
                lineList = currentLine.split(",")
                # print(lineList)
                charFreq = float(lineList[2])
                setValue = float(lineList[3])
                correctedValue = float(lineList[4])
                charFreqList.append(charFreq)
                setValList.append(setValue)
                correctedValList.append(correctedValue)

        filestream.close()

    return (charFreqList, setValList, correctedValList)

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

def isProcessActive(searchTerm):
    import os
    from os import system, name
    import sys

    procList = "wmic process get description,executablepath"

    processList = os.popen(procList).read()
    searchTerm = searchTerm.lower()
    # print("search searchTerm: ",searchTerm)
    processList = processList.split(" ")
    r = []
    for i in processList:
        if "\n" in i:
            a = 1
        else:
            if i != "":
                # print(i)
                r.append(i)

    a = False
    # print(r)
    for i in r:
        j = i.lower()
        if searchTerm in j:
            # print("True")
            a = True

    return a

logFile = "aalog.txt"
cwd = os.getcwd()
useTKINTER = False

if os.path.isfile(logFile):
    # print ("Log File Exists")
    logFile = logFile
else:
    create_log(logFile)

writeLog("================== New Program Iterationn ====================",logFile)

configFile = "aaconfig.cfg"
# Import values from configuration file ------------------------------
print("Importing parameters from the configuration template...")
writeLog("Importing parameters from configuration file: {}".format(str(configFile)), logFile)
time.sleep(0.5)
stdGenName = readConfigFile(configFile, "stdGen")
writeLog("stdGen: "+str(stdGenName),logFile)
corrFileName = readConfigFile(configFile, "correctionFileName")
writeLog("corrFileName: "+str(corrFileName),logFile)
queryInterval = readConfigFile(configFile, "queryInterval")
queryInterval = float(queryInterval)
writeLog("queryInterval: "+str(queryInterval),logFile)
processName = readConfigFile(configFile, "processName")
writeLog("processName: "+str(processName),logFile)
visaResourceIdent = readConfigFile(configFile, "visaResource")
writeLog("visaResourceIdent: "+str(visaResourceIdent),logFile)

stdGenID = readConfigFile(configFile,"stdGenID")
writeLog("stdGenID: "+str(stdGenID),logFile)
stdGenRst = readConfigFile(configFile,"stdGenRst")
writeLog("stdGenRst: "+str(stdGenRst),logFile)
stdGenFreqSet = readConfigFile(configFile,"stdGenFreqSet")
writeLog("stdGenFreqSet: "+str(stdGenFreqSet),logFile)
stdGenPowSet = readConfigFile(configFile,"stdGenPowSet")
writeLog("stdGenPowSet: "+str(stdGenPowSet),logFile)
stdGenFreqRead = readConfigFile(configFile,"stdGenFreqRead")
writeLog("stdGenFreqRead: "+str(stdGenFreqRead),logFile)
stdGenPowRead = readConfigFile(configFile,"stdGenPowRead")
writeLog("stdGenPowRead: "+str(stdGenPowRead),logFile)
stdGenConfig = readConfigFile(configFile,"stdGenConfig")
writeLog("stdGenConfig: "+str(stdGenConfig),logFile)
stdGenOn = readConfigFile(configFile,"stdGenOn")
writeLog("stdGenOn: "+str(stdGenOn),logFile)
stdGenOff = readConfigFile(configFile,"stdGenOff")
writeLog("stdGenOff: "+str(stdGenOff),logFile)
# End Import values from configuration file ------------------------------


# Load correction values from file ---------------------------------------
cwd = os.getcwd()
corrFileName = cwd + "\\" + corrFileName
charFreqList, setValList, correctedValList = readCorrectionFile(corrFileName)
writeLog("charFreqList: "+str(charFreqList),logFile)
writeLog("setValList: "+str(setValList),logFile)
writeLog("correctedValList: "+str(correctedValList),logFile)
# End Loading Correction Values From File --------------------------------


# Load variable names with VISA resources for each instrument --------
clear()
userInterfaceHeader(programName,version,cwd,logFile)
genVisa = GPIB_resource(stdGenName, visaResourceIdent)
writeLog("Leveled Generator Visa Resource: {}".format(str(genVisa)), logFile)

clear()
userInterfaceHeader(programName,version,cwd,logFile,"VISA device response")
response = queryVisa(genVisa,stdGenID)
print("Device " + str(stdGenName) + ": "+str(response))
writeLog("*IDN? Response: {}".format(response),logFile)

# Loop to keep running until the program process closes
clear()
userInterfaceHeader(programName, version, cwd, logFile, "Correction Service Is Active")
processLive = True
correctedValue = "Null"
while processLive == True:
    # Code to check if the process is still active
    processLive = isProcessActive(processName)

    # Get the current power and frequency from the generator
    response = queryVisa(genVisa, stdGenPowRead)
    setPow = float(response)
    response = queryVisa(genVisa, stdGenFreqRead)
    setFreq = float(response)
    
    # Check to see if the current power and frequency are within the characterization list
    for index,i in enumerate(setValList):
        # print("set pow {} - {}, Freq {} - {}".format(setPow,i,setFreq,charFreqList[index])
        if (setPow == i) and (setFreq == charFreqList[index]):
            correctedValue = correctedValList[index]
            # Set the generator to the specified correction
            writeVisa(genVisa, stdGenPowSet .replace("<val>", str(correctedValue)))
            # writeLog("Generator set to {} dBm at {} Hz, corrected to {} dBm".format(setPow,setFreq,correctedValue),logFile)
            clear()
            setFreq = round(setFreq / 1000000)
            correctedValue = round(correctedValue, 3)
            userInterfaceHeader(programName, version, cwd, logFile, "Displaying Most Recent Correction")
            print("Freq: {} MHz, Set Power: {} dBm, Corrected Power: {} dBm".format(setFreq, setPow, correctedValue))        

    
    time.sleep(1)

# Make the GPIB address get pulled from the config file


