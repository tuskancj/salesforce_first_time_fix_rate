import tkinter as tk
import webbrowser as wb
import tkinter.font as tkFont
import csv
from tkinter import Canvas, Frame, filedialog, Text, getdouble, messagebox
import os
from tkinter.constants import W
from typing import List, final
from datetime import date, datetime, timedelta
import enum
from genericpath import exists

#export csv files with UTF-8 encoding

# TO DO:
#clean up/examine logs
#flag timesheets greater than 15 hours?

#Asset Mean Time Between Failure (in calendar days):
#   If Install date pre-dates earliest pulled Timestamp of data from SFDC, 
#   MTBF = [Most Recent TS-Oldest TS]/# Repairs 
#   -Otherwise- 
#   MTBF = [Most Recent TS-Install Date]/# Repairs

#Contract Mean Time Between Failure (in calendar days):
#   MTBF = [Contract End Date - Contract Start Date]/# Repairs

# FLAGS and USER-INPUT variables (change these for different reporting results)
ftf_threshold = 45 #number of days for first time fix rate threshold
flag_quickLoad = True #if true, do not require files to be loaded one by one.  This requires files:  "assets.csv", "contracts.csv", "cases.csv", "WOs.csv", "timesheets.csv"
flag_quickload_initDir = "/" #only needs to be set if flag_quickLoad = False.  Will set initial starting folder to select files.
flag_onsiteLaborOnly = True #if true, only look at onsite labor.  if false, onsite and depot labor will be observed.  travel is excluded

#global variables
IncidentDateTime = "12:00" #military time given to incident date of case
OldestTS:datetime = datetime.today() #placeholder for oldest final onsite WO timestamp.  Will be updated while cycling through data to show oldest timestamp
MostRecentTS:datetime = datetime(2000, 1, 1, 12, 0, 0, 0) #placeholder and will be updated while cycling through data to know the most recent timestamp
log = list() #log list - add logs here.  These will be printed at the end line by line to review
log.append(["method", "description", "metadata"])

#GUI layout
root = tk.Tk()
root.title("MTBF and FTF Rate Algorithm")
root.geometry('600x450')

#global classes
class Asset:
    def __init__(self, productType, serialNumber, currentCoverage, primaryFSE, installationDate, account, status):
        self.ProductType = productType
        self.serialNumber = serialNumber
        self.currentCoverage = currentCoverage
        self.primaryFSE = primaryFSE
        self.installationDate = installationDate
        self.account = account
        self.status = status
        self.workOrders = []    # creates a new empty list for each item.  
        self.contracts = []

    def addWorkOrder(self, WorkOrder):
        self.workOrders.append(WorkOrder)

    def addContract(self, Contract):
        self.contracts.append(Contract)
class WorkOrder:
    def __init__(self, number, type, serialNumber, incidentDate, closeDate, engineer, case, parentCase, account, territory):
        self.number = number #WO number
        self.type = type #WO type
        self.serialNumber = serialNumber #serial number associated with WO
        self.incidentDate = incidentDate #Case incident date
        self.closeDate = closeDate #WO closure date
        self.engineer = engineer #assigned engineer to WO
        self.case = case #case number
        self.parentCase = parentCase #potential parent case of WO's case
        self.account = account
        self.territory = territory #service territory
        self.onsiteTimesheets = [] #list of Timesheets with type onsite labor
        self.travelTimesheets = []
        self.depotTimesheets = []
        self.assetType = self.getAssetType()

    def addOnsiteTimesheet(self, Timesheet):
        self.onsiteTimesheets.append(Timesheet)
    def addTravelTimesheet(self, Timesheet):
        self.travelTimesheets.append(Timesheet)
    def addDepotTimesheet(self, Timesheet):
        self.depotTimesheets.append(Timesheet)
    def getAssetType(self)->str:
        findProduct(self.serialNumber)
class Engineer:
    def __init__(self, name, territory):
        self.name = name
        self.territory = territory
        self.touches = []
        self.touches_Repairs = []
        self.touches_PMIs = []
        self.touches_Installations = []
        self.territoryDictionary = dict() #key = territory name, value = WO/touch instance counter.  The most counts sets Engineer territory every time a touch is added.
        self.dict_RepairsByQtr = dict()
        self.dict_PMIsByQtr = dict()
        self.dict_InstallsByQtr = dict()

    def addTouch(self, _touch):
        #add touch to various lists
        touch:Touch = _touch
        self.touches.append(touch)

        #calculate territory based on WO/Touch territory information
        if touch.territory in self.territoryDictionary:
            count:int = self.territoryDictionary[touch.territory]
            self.territoryDictionary[touch.territory] = count+1
        else:
            self.territoryDictionary[touch.territory] = int(1)

        if self.territory and not self.territory.isspace():
            topTerritory:str = self.territory
        else:
            topTerritory = ""
        for key in self.territoryDictionary:
            count_key = self.territoryDictionary[key]
            count_currentTop = self.territoryDictionary[topTerritory]
            if count_key>count_currentTop:
                topTerritory = key

        self.territory = topTerritory
class Touch:
    def __init__(self, assetType, serialNumber, WO_number, WO_type, TBR, next_WO_number, containsMiddleWO:bool, openEnded:bool, contractTypeAtTime, finalTimestamp, onsiteWOdays, territory, onsiteLaborDuration, assetCurrentStatus, travelLaborDuration, depotLaborDuration):
        self.assetType = assetType
        self.serialNumber = serialNumber
        self.WO_number = WO_number
        self.WO_type = WO_type
        self.TBR = TBR
        self.next_WO_number = next_WO_number
        self.containsMiddleWO = containsMiddleWO
        self.openEnded = openEnded
        self.contractTypeAtTime = contractTypeAtTime
        self.finalTimestamp = finalTimestamp
        self.onsiteWOdays = onsiteWOdays
        self.territory = territory
        self.onsiteLaborDuration = onsiteLaborDuration
        self.assetCurrentStatus = assetCurrentStatus
        self.quarter = self.findQtr(self.finalTimestamp)
        self.travelLaborDuration = travelLaborDuration
        self.depotLaborDuration = depotLaborDuration
    
    def findQtr(self, _finalTimestamp)->str:
        return getQTR_nomenclature(getDateAndTime(_finalTimestamp))
class Timesheet:
    def __init__(self, WO_number, initialTimestamp, finalTimestamp, type, engineer, assetSerialNumber, duration):
        self.WO_number = WO_number
        self.initialTimestamp = initialTimestamp
        self.finalTimestamp = finalTimestamp
        self.type = type
        self.engineer = engineer
        self.assetSerialNumber = assetSerialNumber
        self.duration = duration
class Contract:
    def __init__(self, number, account, serialNumber, type, startDate, endDate, primaryFSE, installationDate):
        self.number = number #WO number
        self.account = account
        self.serialNumber = serialNumber
        self.type = type #WO type
        self.startDate = startDate #start date of contract
        self.endDate = endDate #end date of contract
        self.primaryFSE = primaryFSE
        self.installationDate = installationDate
class Territory:
    def __init__(self, name) -> None:
        self.name = name
        self.touches = []

    def addTouch(self, _touch):
        touch:Touch = _touch

        #add to full touches list
        self.touches.append(touch)
    
#links
link_AT = "https://cytekdevelopment.lightning.force.com/lightning/r/Report/00OUe0000002yT7MAI/view?queryScope=userFoldersCreatedByMe"
link_WO = "https://cytekdevelopment.lightning.force.com/lightning/r/Report/00OUe0000002yZZMAY/view?queryScope=userFoldersCreatedByMe"
link_TS = "https://cytekdevelopment.lightning.force.com/lightning/r/Report/00OUe0000002yXxMAI/view?queryScope=userFoldersCreatedByMe"
link_CT = "https://cytekdevelopment.lightning.force.com/lightning/r/Report/00OUe0000002yUjMAI/view?queryScope=userFoldersCreatedByMe"
link_CS = "https://cytekdevelopment.lightning.force.com/lightning/r/Report/00OUe0000002yWLMAY/view?queryScope=userFoldersCreatedByMe"

#global variables
dict_cases = dict() #dictionary, key = case, value = list of WorkOrder Objects
dict_finalTimeSheets = dict() #dictionary, key = WO, value = Timesheet object
dict_initialTimeSheets = dict() #dictionary, key = WO, value = Timesheet object
dict_assets = dict() #dictionary, key = serial number, value = Asset object
dict_engineers = dict() #dictionary, key = engineer name, value = Engineer object
dict_territories = dict() #dictionary, key = territory name, value = Territory object
dict_productLines = dict() #dictionary, key = Product Line, value ProductLine Object
list_MTBR_sameCase = list()
text_loaded = "Loaded!"
text_loaded_AT = tk.StringVar()
text_loaded_AT.set("")
text_loaded_CT = tk.StringVar()
text_loaded_CT.set("")
text_loaded_CS = tk.StringVar()
text_loaded_CS.set("")
text_loaded_WO = tk.StringVar()
text_loaded_WO.set("")
text_loaded_TS = tk.StringVar()
text_loaded_TS.set("")
filenameBase = ""

#methods
def getDateAndTime(s)->datetime:
    #returns datetime with year, month, day, hour, minute ints
    if len(s) > 11:
        spl = s.split(" ")
        dateStr = spl[0].split("/")
        # timeStr = spl[1].split(":")
        timeSpl = spl[1].split(":")
        new = timeSpl[0].zfill(2)
        iTime = new+":"+timeSpl[1]+" "+spl[2]
        military_time = datetime.strptime(iTime, '%I:%M %p').strftime('%H:%M')
        timeStr = military_time.split(":")
        returnDateTime = datetime(int(dateStr[2]), int(dateStr[0]), int(dateStr[1]), int(timeStr[0]), int(timeStr[1]))
        return returnDateTime
    elif len(s) > 0:
        
        dateStr = s.split("/")
        timeStr = IncidentDateTime.split(":")
        returnDateTime = datetime(int(dateStr[2]), int(dateStr[0]), int(dateStr[1]), int(timeStr[0]), int(timeStr[1]))
        return returnDateTime
    else:
        timeStr = IncidentDateTime.split(":")
        return datetime(int('1900'), int('01'), int('01'), int(timeStr[0]), int(timeStr[1]))
def checkWithOldestTimestamp(s):
    global OldestTS
    ts = getDateAndTime(s)
    if ts < OldestTS:
        OldestTS = ts
def checkWithMostRecentTimestamp(s):
    global MostRecentTS
    ts = getDateAndTime(s)
    if ts > MostRecentTS:
        MostRecentTS = ts
def findProduct(serialNumber)->str:
    #flow sight = 'FS'
    #imagestream = 'ISX'
    #cellstream = 'CS'
    #guava = starts with 'EC1, EC2, 543, 662, 673, 847, 857, GTI'
    #muse = starts with 'MU, 720, WXS'

    switch={
        'N': "Northern Lights",
        'NL': "Northern Lights",
        'NM': "Northern Lights",
        'NE': "Northern Lights",
        'R': "Aurora 3L",
        'RM': "Aurora 3L",
        'Y': "Aurora 4L - YG",
        'V': "Aurora 4L - UV",
        'U': "Aurora 5L",
        'S': "Aurora CS",
        'FS': 'FlowSight',
        'CS': 'CellStream',
        'MU': 'Muse',
        'ISX': 'ImageStream',
        'EC1': 'Guava',
        'EC2': 'Guava',
        '543': 'Guava',
        '662': 'Guava',
        '673': 'Guava',
        '847': 'Guava',
        '857': 'Guava',
        'GTI': 'Guava',
        '720': 'Muse',
        'WXS': 'Muse'
    }
    type:str

    type = 'Other'

    try:
        if len(serialNumber)>=1:
            #known cytek serial numbers:
            if(serialNumber[0] in ['N', 'R', 'Y', 'V', 'U', 'S']):
                if len(serialNumber) == 5 or 'SP' in serialNumber:
                    type = switch.get(serialNumber[0],"Other")
            if(type == 'Other'):
                type = switch.get(serialNumber[0:2], "Other")
            if(type == 'Other'):
                type = switch.get(serialNumber[0:3], 'Other')
    except:
        print('exception' + serialNumber)
            
    return type
def contractIsCurrent(c:Contract)->bool:
    cStart = getDateAndTime(c.startDate)
    cEnd = getDateAndTime(c.endDate)
    todayy = datetime.today()
    # print(todayy)
    # print(cStart)
    # print(cEnd)
    # if cStart < cEnd:
    #     print("works")
    # if todayy < cEnd:
    #     print("less than end")
    if cStart<todayy and todayy<cEnd:
        return True
    else:
        return False
def findContractAtTimeOfWO(t:datetime, a:Asset)->str:
    cType = "Billable"
    for contract in a.contracts:
        c:Contract = contract
        startDate = getDateAndTime(c.startDate)
        endDate = getDateAndTime(c.endDate)
        # if t>=startDate:
        #     if a.serialNumber == "U0867":
        #         print("greater than")
        # if t<=endDate:
        #     if a.serialNumber == "U0867":
        #         print("less than")
        if t>=startDate and t<=endDate:
            cType = c.type
            break
    return cType
def findNumberOfOnsiteTimestamps(w:WorkOrder)->str:
    list_days = []
    for ts in w.onsiteTimesheets:
        ts:Timesheet
        d = getDateAndTime(ts.initialTimestamp).date()
        if str(d) not in list_days:
            list_days.append(str(d))
    return str(len(list_days))
def findQTR(month:str)->str:
    # switch={
    #    "January": "Q1",
    #    "February": "Q1",
    #    "March": "Q1",
    #    "April": "Q2",
    #    "May": "Q2",
    #    "June": "Q2",
    #    "July": "Q3",
    #    "August": "Q3",
    #    "September": "Q3",
    #    "October": "Q4",
    #    "November": "Q4",
    #    "December": "Q4",
    #    }
    switch={
       "1": "Q1",
       "2": "Q1",
       "3": "Q1",
       "4": "Q2",
       "5": "Q2",
       "6": "Q2",
       "7": "Q3",
       "8": "Q3",
       "9": "Q3",
       "10": "Q4",
       "11": "Q4",
       "12": "Q4",
       }
    type:str
    q = switch.get(month,"Q1")
    return q
def getQTR_buckets()->list:
    l = list()
    

    oldest_month = OldestTS.month
    oldest_year = OldestTS.year
    newest_month = MostRecentTS.month
    newest_year = MostRecentTS.year
    oldest_qtr = findQTR(str(oldest_month))
    newest_qtr = findQTR(str(newest_month))
    current_q = int(oldest_qtr[1])
    final_q = 4
    #nomenclature = 2021_Q1, 2021_Q2, etc
    for y in range(oldest_year, newest_year+1, 1): #oldest year remove +1, newest year add +1
        if y == newest_year:
            final_q = int(newest_qtr[1])
        for q in range(current_q, final_q+1, 1): #final q +1
            s = str(y)+"_Q"+str(q)
            l.append(s)
        current_q = 1
    return l
def getQTR_nomenclature(dt:datetime)->str:
    # dt = getDateAndTime(finalTimestamp)
    y = str(dt.year)
    q = findQTR(str(dt.month))
    s = y+"_"+q
    return s
def findAssetTerritory(asset:Asset, methodForLogging:str)->Territory:
    territory:Territory
    assetFSE:Engineer
    if asset.primaryFSE in dict_engineers:
        assetFSE = dict_engineers[asset.primaryFSE]
        #get Territory obj
        if assetFSE.territory in dict_territories:
            territory = dict_territories[assetFSE.territory]
        else:
            log.append([methodForLogging, "territory doesn't exist for the primary FSE on asset "+asset.serialNumber+" at account "+asset.account+".  Assigning asset territory to ''.  MTBF data will be lost"])
            territory = Territory("")
    else:
        #there are assets with incorrect primary objects assigned.  E.g. middle initial "L" vs no middle initial.  Need to find primary FSE based on list of WOs in this scenario
        dict_possibleFSEs = dict() #key=engineer name str, value = count
        for _tempWO, index in enumerate(asset.workOrders):
            tempWO:WorkOrder = asset.workOrders[_tempWO]
            if tempWO.engineer in dict_possibleFSEs:
                tempInt:int = dict_possibleFSEs[tempWO.engineer]
                newCount = tempInt+1
                dict_possibleFSEs[tempWO.engineer] = newCount
            else:
                dict_possibleFSEs[tempWO.engineer] = 1
        sorted(dict_possibleFSEs, key=dict_possibleFSEs.get, reverse=True)
        # next(iter(my_dict)) #this will get first item in list
        found = False
        for key in dict_possibleFSEs:
            if key in dict_engineers:
                assetFSE = dict_engineers[key]
                #get Territory obj
                if assetFSE.territory in dict_territories:
                    territory = dict_territories[assetFSE.territory]
                    found = True
                    break
                # else:
        if found:
            log.append([methodForLogging, "primary FSE '" + asset.primaryFSE + "' doesn't exist for asset "+asset.serialNumber+" at account "+asset.account+".  Found that '" + assetFSE.name + "' had the most WOs assigned.  Assigning to '" + assetFSE.name + "'s' territory of "+assetFSE.territory])
        else:
            log.append([methodForLogging, "territory doesn't exist for the primary FSE '" + asset.primaryFSE + "' on asset "+asset.serialNumber+" at account "+asset.account+".  Could not find a suitable replacement by scanning the rest of the WOs for this asset.  Assigning asset territory to ''."])
            territory = Territory("")
        
    return territory

# asks user to select csv file, then stores Asset info in dictionary (key, value)
# key --> Asset Serial Number
# value --> Asset Object
def parseAssets():
    #clear variables
    text_loaded_AT.set("")
    dict_assets.clear()

    #get the file
    if flag_quickLoad:
        filename = "assets.csv"
    else:
        filename = filedialog.askopenfilename(parent = root, initialdir="/Users/ctuskan/Desktop", title="Select Assets CSV File", filetypes=(("CSV", "*.csv"), ("All Files", "*.*")))
    with open(filename, encoding='utf-8') as csvfile:
        readCSV = csv.reader(csvfile, delimiter=',')
        for row in readCSV:
            if(readCSV.line_num!=1): # skip the header
                a = Asset(findProduct(row[1]), row[1], "Billable", row[6], row[2], row[3], row[4])
                p = a.ProductType
                # if p=="Other":
                #     continue

                if a.serialNumber in dict_assets:
                    log.append(["parse Assets", a.serialNumber + " is listed twice in Assets report.  Ignoring anything other than the first entry"])
                else:
                    dict_assets[a.serialNumber] = a
                    
        #update user
        if flag_quickLoad:
            parseContracts()
        else:
            text_loaded_AT.set(text_loaded)

# asks user to select csv file, then stores Asset Contract info in dictionary (key, value)
# key --> Asset Serial Number
# value --> Asset Object which holds a list of contracts
def parseContracts():
    #clear variables
    text_loaded_CT.set("")

    #get the file
    if flag_quickLoad:
        filename = "contracts.csv"
    else:
        filename = filedialog.askopenfilename(parent = root, initialdir=flag_quickload_initDir, title="Select Contracts CSV File", filetypes=(("CSV", "*.csv"), ("All Files", "*.*")))
    with open(filename, encoding='utf-8') as csvfile:
        readCSV = csv.reader(csvfile, delimiter=',')
        for row in readCSV:
            if(readCSV.line_num!=1): # skip the header
                ct = Contract(row[0], row[1], row[2], row[5], row[6], row[7], row[4], row[8])
                # if ct.serialNumber == "U0867":
                #     print("test")
                p = findProduct(ct.serialNumber)
                # if p=="Other":
                #     continue

                if ct.serialNumber in dict_assets:
                    a:Asset = dict_assets[ct.serialNumber]
                    a.addContract(ct)
                    if contractIsCurrent(ct):
                        a.currentCoverage = ct.type
                else:
                    log.append(["parse Contracts", ct.serialNumber + " at " + ct.account +" is not in Assets dictionary.  Ignoring."])
                    
        #update user
        if flag_quickLoad:
            parseParentCases()
        else:
            text_loaded_CT.set(text_loaded)

# asks user to select csv file, then stores case, parent case in dictionary (key, value)
# key --> case as 8 digit number
# value --> parent case
def parseParentCases():
    #clear variables
    text_loaded_CS.set("")
    dict_cases.clear()

    #get the file
    if flag_quickLoad:
        filename = "cases.csv"
    else:
        filename = filedialog.askopenfilename(parent = root, initialdir=flag_quickload_initDir, title="Select Cases CSV File", filetypes=(("CSV", "*.csv"), ("All Files", "*.*")))
    with open(filename, encoding='utf-8') as csvfile:
        readCSV = csv.reader(csvfile, delimiter=',')
        for row in readCSV:
            if(readCSV.line_num!=1): # skip the header
                case = row[0]
                parent_case = row[1]

                # cases duplicate if multiple wo are within the case.
                if case in dict_cases:
                    desc = case +' already in cases dictionary with parent case of ' + dict_cases[case] + '. Skipping for parent case '+ parent_case +'.'
                    log.append(['parseParentCases', desc])
                else:
                    dict_cases[case] = parent_case
        #update user
        if flag_quickLoad:
            parseWorkOrders()
        else:
            text_loaded_CS.set(text_loaded)

# asks user to select csv file, then stores WO, closure date, and owner in dictionary (key, value)
# key --> WO as 8 digit number
# value --> Asset Object which holds a list of WOs
def parseWorkOrders():
    #clear variables
    text_loaded_WO.set("")

    #get the file
    if flag_quickLoad:
        filename = "WOs.csv"
    else:
        filename = filedialog.askopenfilename(parent = root, initialdir=flag_quickload_initDir, title="Select Work Orders CSV File", filetypes=(("CSV", "*.csv"), ("All Files", "*.*")))
    with open(filename, encoding='utf-8') as csvfile:
        readCSV = csv.reader(csvfile, delimiter=',')
        for row in readCSV:
            if(readCSV.line_num!=1): # skip the header
                wo = WorkOrder(row[0], row[7], row[6], row[1], row[2], row[3], row[9], "NA", row[4], row[10]) #set WO object

                if wo.case in dict_cases:
                    wo.parentCase = dict_cases[wo.case]

                p = findProduct(wo.serialNumber)
                # if p=="Other":
                #     continue

                if wo.serialNumber in dict_assets:
                    a:Asset = dict_assets[wo.serialNumber]
                    a.addWorkOrder(wo)
                else:
                    # #assigning the WO FSE here which isn't necessarily account primary, but most of these assets with no documented contract are 'other'
                    log.append(["parse WOs","no serial number found in Assets dictionary for WO " + wo.number + ", s/n "+ wo.serialNumber +" at " + wo.account + " when parsing Work Orders.  Ignoring.  This is likely an asset that was returned to Cytek for one reason or another."])
                    
        #update user
        if flag_quickLoad:
            parseTimeSheets()
        else:
            text_loaded_WO.set(text_loaded)

# asks user to select csv file, then stores WO & latest timesheet in one of two dictionaries (key, value)
# one dictionary for initial timesheets, the other for final timesheets
# key --> WO as 8 digit number
# value --> Timesheet object
# ALL timesheets are stored as a list within the WO object that is contained in the Asset list of WOs
def parseTimeSheets():
    #clear variables
    text_loaded_TS.set("")
    dict_finalTimeSheets.clear()

    #ask user to load file
    if flag_quickLoad:
        fileName = "timesheets.csv"
    else:
        fileName = filedialog.askopenfilename(parent = root, initialdir=flag_quickload_initDir, title="Select Timesheets CSV File", filetypes=(("CSV", "*.csv"), ("All Files", "*.*")))
    with open(fileName, encoding='utf-8') as csvfile:
        readCSV = csv.reader(csvfile, delimiter=',')
        replaceFinal = False
        replaceInitial = False
        for row in readCSV:
            if(readCSV.line_num!=1): # skip the header
                ts:Timesheet = Timesheet(row[0], row[6], row[7], row[9], row[4], row[11], row[5]) #get timesheet object
                p = findProduct(ts.assetSerialNumber)
                # if p=="Other":
                #     continue

                if "Onsite" in ts.type: #ensure timesheet is onsite labour type, then store in Asset WO for counting number of onsite timesheets per WO later
                    checkWithOldestTimestamp(ts.initialTimestamp) #these functions are used to tell the program the timespan of all dates pulled from SFDC report
                    checkWithMostRecentTimestamp(ts.finalTimestamp)

                    if ts.assetSerialNumber in dict_assets:
                        a:Asset = dict_assets[ts.assetSerialNumber]
                        found = False
                        for wo in a.workOrders:
                            if wo.number == ts.WO_number:
                                wo.addOnsiteTimesheet(ts)
                                found = True
                                break
                        if not found:
                            log.append(["parse Timesheets","WO "+ ts.WO_number + " not found in Asset " + a.serialNumber +" WO list when trying to add timesheet with initial labor:  " + ts.initialTimestamp + " | final labor:  "+ ts.finalTimestamp])
                    else:
                        log.append(["parse Timesheets", "Timesheet serial number "+ ts.assetSerialNumber + " not found in Asset dictionary.  Unable to log onsite timesheets to associated Asset->Work Order"])

                if "Travel" in ts.type: #capture travel time
                    if ts.assetSerialNumber in dict_assets:
                        a:Asset = dict_assets[ts.assetSerialNumber]
                        found = False
                        for wo in a.workOrders:
                            if wo.number == ts.WO_number:
                                wo.addTravelTimesheet(ts)
                                found = True
                                break
                        if not found:
                            log.append(["parse Timesheets","WO "+ ts.WO_number + " not found in Asset " + a.serialNumber +" WO list when trying to add TRAVEL timesheet with initial labor:  " + ts.initialTimestamp + " | final labor:  "+ ts.finalTimestamp])
                    else:
                        log.append(["parse Timesheets", "TRAVEL Timesheet with serial number "+ ts.assetSerialNumber + " not found in Asset dictionary.  Unable to log onsite timesheets to associated Asset->Work Order"])

                if "Depot" in ts.type: #capture depot time
                    if ts.assetSerialNumber in dict_assets:
                        a:Asset = dict_assets[ts.assetSerialNumber]
                        found = False
                        for wo in a.workOrders:
                            if wo.number == ts.WO_number:
                                wo.addDepotTimesheet(ts)
                                found = True
                                break
                        if not found:
                            log.append(["parse Timesheets","WO "+ ts.WO_number + " not found in Asset " + a.serialNumber +" WO list when trying to add DEPOT timesheet with initial labor:  " + ts.initialTimestamp + " | final labor:  "+ ts.finalTimestamp])
                    else:
                        log.append(["parse Timesheets", "DEPOT Timesheet with serial number "+ ts.assetSerialNumber + " not found in Asset dictionary.  Unable to log onsite timesheets to associated Asset->Work Order"])

                #FINAL TIMESHEET CHECK
                if ts.WO_number in dict_finalTimeSheets: # check if WO exists in final timesheets dict
                    prevTS:Timesheet = dict_finalTimeSheets[ts.WO_number] #get previous timesheet
                    prevDateTime = getDateAndTime(prevTS.finalTimestamp) #get previous finalTimestamp saved in dictionary
                    currentDateTime = getDateAndTime(ts.finalTimestamp) #get current finalTimestamp
                    if currentDateTime>prevDateTime:
                        replaceFinal = True
                        #also compare timecard owner before existing?
                else: #doesn't exist, add to final timesheet dictionary
                    replaceFinal = True
                if replaceFinal:
                    if flag_onsiteLaborOnly:
                        if "Onsite" in ts.type: #ensure timesheet is labor/labour type, not travel or blank
                            dict_finalTimeSheets[ts.WO_number] = ts #add object to dictionary under WO number key
                    else:
                        if "Lab" in ts.type:
                            dict_finalTimeSheets[ts.WO_number] = ts #add object to dictionary under WO number key
                    replaceFinal = False

                #INITIAL TIMESHEET CHECK
                if ts.WO_number in dict_initialTimeSheets: # check if WO exists in initial timesheets dict
                    prevTS:Timesheet = dict_initialTimeSheets[ts.WO_number] #get previous timesheet
                    prevDateTime = getDateAndTime(prevTS.initialTimestamp) #get previous initialTimestamp saved in dictionary
                    currentDateTime = getDateAndTime(ts.initialTimestamp) #get current initialTimestamp
                    if prevDateTime>currentDateTime:
                        replaceInitial = True
                        #also compare timecard owner before existing?
                else: #doesn't exist, add to initial timesheet dictionary
                    replaceInitial = True
                if replaceInitial:
                    if flag_onsiteLaborOnly:
                        if "Onsite" in ts.type: #ensure timesheet is labor/labour type, not travel or blank
                            dict_initialTimeSheets[ts.WO_number] = ts #add object to dictionary under WO number key
                    else:
                        if "Lab" in ts.type:
                            dict_initialTimeSheets[ts.WO_number] = ts #add object to dictionary under WO number key
                    replaceInitial = False
                    
    if not flag_quickLoad:
        text_loaded_TS.set(text_loaded)
    findTBRs()


def findTBRs():
    #clear the same-case list and set column headers
    list_MTBR_sameCase.clear()
    list_MTBR_sameCase.append(["Serial Number", "Case", "Current WO", "WO Owner", "Current WO Type", "Current WO Final Timestamp", "Next WO Case", "Next WO Case Parent", "Next WO", "Next WO Owner", "Next WO Type", "Next WO Initial Timestamp"])
    
    for key in dict_assets:
        asset:Asset = dict_assets[key]
        dict_assetWOs = dict() #create dictionary of WOs for each asset.  Key = wo number, value = WO object
        for obj in asset.workOrders:
            wo:WorkOrder = obj
            dict_assetWOs[wo.number] = wo       

        #sort the WOs based on final timesheet end datetime
        tempWoList = list(dict_assetWOs.keys())
        dict_WO_And_finalTimesheet = dict()
        list_WO_withNoTime_index = []
        for index, wo in enumerate(tempWoList):
            if wo in dict_finalTimeSheets:
                ts:Timesheet = dict_finalTimeSheets[wo]
                dict_WO_And_finalTimesheet[wo] = getDateAndTime(ts.finalTimestamp)
            else:
                wOrder = dict_assetWOs[wo] #get the WorkOrder obj
                log.append(["FSE TBR","timesheet doesn't exist in WO " + wo + ". removing from tempWOlist.  FSE, serialNumber, Account->", wOrder.engineer, wOrder.serialNumber, wOrder.account])
                list_WO_withNoTime_index.insert(0,index)
        for i in list_WO_withNoTime_index:
            del tempWoList[i]
        orderedWoList = sorted(tempWoList, key = dict_WO_And_finalTimesheet.__getitem__)

        #iterate through ordered WO list to set MTBRs
        for index, value in enumerate(orderedWoList):
            currentWO:WorkOrder = dict_assetWOs[orderedWoList[index]] #get current WO
            nextRepairWO:WorkOrder
            ts_final:Timesheet = dict_finalTimeSheets[currentWO.number]
            currentCompletionDate = getDateAndTime(ts_final.finalTimestamp)
            mtbr:str
            openEnded = True
            containsMiddleWO = False

            if (index+1) <= (len(orderedWoList)-1): #ensure there is another WO
                r = range(index+1, len(orderedWoList))
                
                for n in r:
                    sn = orderedWoList[n]
                    nextRepairWO = dict_assetWOs[sn]
                    if nextRepairWO.type == "Repair" or nextRepairWO.type == "Service":
                        
                        #find time difference between current WO completion date and repair WO initial Labor date
                        ts_initial:Timesheet = dict_initialTimeSheets[nextRepairWO.number]
                        nextInitialDate = getDateAndTime(ts_initial.initialTimestamp)
                        # if nextInitialDate>=currentCompletionDate: #ensure dates align with WO number in ascending order
                        if currentWO.case != "" and currentWO.case != nextRepairWO.case and nextRepairWO.parentCase != currentWO.case: #same case repairs shouldn't go against mtbr.  Skip 
                            openEnded = False
                            delta = nextInitialDate-currentCompletionDate
                            hours = float(delta.days)*float(24) + float(delta.seconds)/float(60*60) #datetime subtraction only generates days and seconds variables
                            days = hours/float(24)
                            mtbr = str(format(days, '.2f'))
                            break
                        elif currentWO.case == "":
                            log.append(["find MTBRs", "WO "+currentWO.number+" does not have a case associated."])
                        elif currentWO.case == nextRepairWO.case or nextRepairWO.parentCase == currentWO.case:
                            list_MTBR_sameCase.append([currentWO.serialNumber, currentWO.case, currentWO.number, currentWO.engineer, currentWO.type, str(currentCompletionDate), nextRepairWO.case, nextRepairWO.parentCase, nextRepairWO.number, nextRepairWO.engineer, nextRepairWO.type, str(nextInitialDate)])
                    else:
                        containsMiddleWO = True

                        #continue to add same-case WOs even if they're not repair
                        if currentWO.case == "":
                            log.append(["find MTBRs", "WO "+currentWO.number+" does not have a case associated."])
                        elif currentWO.case == nextRepairWO.case or nextRepairWO.parentCase == currentWO.case:
                            list_MTBR_sameCase.append([currentWO.serialNumber, currentWO.case, currentWO.number, currentWO.engineer, currentWO.type, str(currentCompletionDate), nextRepairWO.case, nextRepairWO.parentCase, nextRepairWO.number, nextRepairWO.engineer, nextRepairWO.type, str(nextInitialDate)])

            #add touch to engineer
            if openEnded: #get current datetime
                # currentDateTime = datetime.now()
                currentDateTime = MostRecentTS
                delta = currentDateTime-currentCompletionDate
                hours = delta.days*24 + delta.seconds/(60*60) #datetime subtraction only generates days and seconds variables
                mtbr = str(format(hours/24, '.2f'))
                
            con = "Billable"
            if currentWO.type == "Installation":
                con = "Warranty"
            else:
                con = findContractAtTimeOfWO(currentCompletionDate, dict_assets[currentWO.serialNumber])

            #get total labor days and total duration of onsite/depot labor
            laborDays = findNumberOfOnsiteTimestamps(currentWO)
            oLaborDuration:float = float(0.0)
            for _ts in currentWO.onsiteTimesheets:
                ts:Timesheet = _ts
                oLaborDuration += float(ts.duration)

            tLaborDuration:float = float(0.0)
            for _ts in currentWO.travelTimesheets:
                ts:Timesheet = _ts
                tLaborDuration += float(ts.duration)

            dLaborDuration:float = float(0.0)
            for _ts in currentWO.depotTimesheets:
                ts:Timesheet = _ts
                dLaborDuration += float(ts.duration)

            #generate touch obj
            nextWO = 'NA'
            if not openEnded:
                nextWO = nextRepairWO.number
            touch = Touch(asset.ProductType,currentWO.serialNumber, currentWO.number, currentWO.type, mtbr, nextWO, containsMiddleWO, openEnded, con, ts_final.finalTimestamp, laborDays, currentWO.territory, str(round(oLaborDuration, 2)), asset.status, str(round(tLaborDuration, 2)), str(round(dLaborDuration, 2)))    

            #add to engineer dictionary
            if currentWO.engineer in dict_engineers:
                eng:Engineer = dict_engineers[currentWO.engineer]
                eng.addTouch(touch)
            else:
                eng = Engineer(currentWO.engineer, currentWO.territory)
                eng.addTouch(touch)
                dict_engineers[currentWO.engineer] = eng

            #add to territory dictionary
            if touch.territory in dict_territories:
                territory:Territory = dict_territories[touch.territory]
                territory.addTouch(touch)
            else:
                territory:Territory = Territory(touch.territory)
                territory.addTouch(touch)
                dict_territories[touch.territory] = territory

    writeFiles()
def writeFiles():
    if flag_quickLoad:
        folderName = checkIfFolderExists("OUTPUT")
        os.makedirs(folderName)
    else:
        folderName = filedialog.askdirectory(parent = root, initialdir=flag_quickload_initDir, title="Please Select Folder to Write Files")
    
    #write raw FSE TBR file
    filepath_rawFSEtouches = folderName+"/TBR_raw_FSE.csv"
    writeRawFSETouchFileTo(filepath_rawFSEtouches)

    #write raw Asset MTBF file
    filepath_assetMTBF = folderName + "/MTBF_raw_Asset.csv"
    writeRawAssetMTBFfileTo(filepath_assetMTBF)

    #write same case file (list of WOs with the same case or WOs with same parent case)
    filepath_sameCase = folderName + "/log_same_case.csv"
    writeSameCaseFileTo(filepath_sameCase)

    #write contract MTBF file
    filepath_contractMTBF = folderName + "/MTBF_raw_Contracts.csv"
    writeRawContractMTBFfileTo(filepath_contractMTBF)
    
    #write log file
    filepath_logs = folderName + "/log_main.csv"
    writeLogFileTo(filepath_logs)

    print("earliest timestamp:  "+ str(OldestTS))
    print("most recent timestamp:  "+ str(MostRecentTS))
    print("FINISHED!")  
    exit()  
def writeRawFSETouchFileTo(filepath:str):
    file = open(filepath, 'w+', newline ='', encoding='utf-8')
    with file:
        write = csv.writer(file)
        ftf_str = str(ftf_threshold) + '-day FTF'
        write.writerow(["Quarter", "Engineer", "WO Number", "WO Type", "Serial Number", "Product", "Product SubType", "Status", "TBR", "Next Repair WO Number", ftf_str, "Contains Middle WO", "Open Ended", "Contract Type During Visit", "FinalOnsiteTimestamp", "# Onsite WO days", "Service Territory", "Onsite Labor Hours", "Travel Labor Hours", "Depot Labor Hours"])
        for key in dict_engineers:
            e:Engineer = dict_engineers[key]
            for _t in e.touches:
                # if e.name != '__GLOBAL FSEs':
                t:Touch = _t

                #set product sub-type
                subType = "NA"
                productType = t.assetType
                if "Aurora" in t.assetType and "CS" not in t.assetType:
                    subType = t.assetType
                    productType = "Aurora"

                #find if FTF rate (if not adequate time to reach ftf threshold, mark 'NA')
                valid_ftf = True
                diff = MostRecentTS-getDateAndTime(t.finalTimestamp)
                if int(diff.days) < ftf_threshold or t.TBR == '' or float(t.TBR) < 0:
                    valid_ftf = 'NA'
                elif float(t.TBR) < ftf_threshold:
                    valid_ftf = False
                
                #write to file
                write.writerow([t.quarter, e.name, t.WO_number, t.WO_type, t.serialNumber, productType, subType, t.assetCurrentStatus, t.TBR, t.next_WO_number, str(valid_ftf), t.containsMiddleWO, t.openEnded, t.contractTypeAtTime, t.finalTimestamp, t.onsiteWOdays, t.territory, t.onsiteLaborDuration, t.travelLaborDuration, t.depotLaborDuration])

    print('Engineers with MTBF:  ' + str(len(dict_engineers)))
def writeRawAssetMTBFfileTo(filepath:str):
    assetMTBFFile = open(filepath, 'w+', newline='', encoding='utf-8')

    list_assetMTBF_byInstallDate = list() #create list of asset MTBRs based on install date and contract date
    list_assetMTBF_byInstallDate.append(["Serial #", "Product", "Product SubType", "Status", "Account", "Estimated Territory", "Install Date", "Primary FSE", "# of Repairs", "MTBF"])
    

    for sn in dict_assets:
        asset:Asset = dict_assets[sn]
        repairCount = 0
        for wo in asset.workOrders: #find number of WOs in dictionary
            wo:WorkOrder

            #filter out erroneous WOs
            if wo.type == "Repair" or wo.type == "Service":
                if wo.engineer != "Lakshmi Poluru":
                    repairCount+=1
        
        if asset.installationDate == "":
            log.append(["writeRawAssetMTBFfileTo", "asset " + asset.serialNumber + " doesn't have an installation date.  skipping MTBF for this asset."])
            continue            

        # numRepairs = str(count) #find lifespan between most recent timestamp and either installation date or oldest TS, whichever is most recent
        dt_installation = getDateAndTime(asset.installationDate)
        lifeSpan:timedelta
        if dt_installation > OldestTS:
            lifeSpan = MostRecentTS-dt_installation
        else:
            lifeSpan = MostRecentTS-OldestTS
        
        #find territory based on primary FSE
        territory:Territory = findAssetTerritory(asset, "writeRawAssetMTBFfileTo")
        
        subType = "NA"
        productType = asset.ProductType
        if "Aurora" in asset.ProductType and "CS" not in asset.ProductType:
            subType = asset.ProductType
            productType = "Aurora"

        if repairCount>0:
            mtbf:float = float(lifeSpan.days)/float(repairCount)
            mtbf_str = str(round(mtbf, 1))
            list_assetMTBF_byInstallDate.append([asset.serialNumber, productType, subType, asset.status, asset.account, territory.name, asset.installationDate, asset.primaryFSE, str(repairCount), mtbf_str])
        else: #don't divide by zero
            span = str(lifeSpan.days)
            list_assetMTBF_byInstallDate.append([asset.serialNumber, productType, subType, asset.status, asset.account, territory.name, asset.installationDate, asset.primaryFSE, "0", span])

    with assetMTBFFile:
        write = csv.writer(assetMTBFFile)
        print("Assets with MTBF:  "+str(len(list_assetMTBF_byInstallDate)))
        write.writerows(list_assetMTBF_byInstallDate)
    
    # return dict_productTypeMTBR
def writeRawContractMTBFfileTo(filepath:str):
    contractMTBFFile = open(filepath, 'w+', newline='', encoding='utf-8')

    list_contractMTBF = list() #create list of contract MTBFs 
    list_contractMTBF.append(["Serial #", "Installation Date", "Current Status", "Product", "Product Subtype", "Account", "Primary FSE", "Estimated Territory", "Contract Number", "Contract Type", "Contract Start Date", "Contract End Date", "Contract Time Span (Days)", "# of Repairs During Contract Period", "MTBF", "Annual Repair Extrapolation"])
    
    for _sn in dict_assets:
        asset:Asset = dict_assets[_sn]
        for row in asset.contracts:
            contract:Contract = row
            dt_startDate = getDateAndTime(contract.startDate)
            dt_endDate = getDateAndTime(contract.endDate)
            td_span = dt_endDate-dt_startDate
            
            repairCount = 0
            #scroll through WOs to count # of repairs during contract time span
            for row in asset.workOrders:
                wo:WorkOrder = row
                if wo.type == "Repair" or wo.type == "Service":
                    if wo.number in dict_finalTimeSheets:
                        finalTimesheet:Timesheet = dict_finalTimeSheets[wo.number]
                        dt_finalTimesheet = getDateAndTime(finalTimesheet.finalTimestamp)
                        if dt_finalTimesheet >= dt_startDate and dt_finalTimesheet <= dt_endDate:
                            repairCount+=1
                    else: #go by wo close date
                        if wo.closeDate != "":
                            dt_closeDate = getDateAndTime(wo.closeDate)
                            if dt_closeDate >= dt_startDate and dt_closeDate <= dt_endDate:
                                repairCount+=1
                            else:
                                if asset.serialNumber != '':
                                    log.append(["writeRawContractMTBFfileTo", "wo " + wo.number + " is not in final timesheets dictionary.  Therefore using close date to determine if Repair WO is within contract time span.  However, WO close date is nil.  Repair will not be logged for conract " + contract.number + ", asset "+ asset.serialNumber + ", account "+asset.account])

            territory:Territory = findAssetTerritory(asset, "writeRawContractMTBFfileTo")

            try:
                mtbf = round(float(td_span.days)/float(repairCount), 1)
            except ZeroDivisionError:
                mtbf = td_span.days

            try:
                annual_repairs = 365/td_span.days*repairCount
            except ZeroDivisionError:
                annual_repairs = 0

            #set product syb types
            subType = "NA"
            productType = asset.ProductType
            if "Aurora" in asset.ProductType and "CS" not in asset.ProductType:
                subType = asset.ProductType
                productType = "Aurora"
            list_contractMTBF.append([asset.serialNumber, asset.installationDate, asset.status, productType, subType, asset.account, asset.primaryFSE, territory.name, contract.number, contract.type, contract.startDate, contract.endDate, str(td_span.days), str(repairCount), str(mtbf), str(round(annual_repairs, 1))])

    with contractMTBFFile:
        write = csv.writer(contractMTBFFile)
        write.writerows(list_contractMTBF)
def writeLogFileTo(filepath:str):
    logFile = open(filepath, 'w+', newline='', encoding='utf-8')
    with logFile:
        write = csv.writer(logFile)
        print("log counts:  "+str(len(log)))
        write.writerows(log)
def writeSameCaseFileTo(filepath:str):
    sameCaseFile = open(filepath, 'w+', newline='', encoding='utf-8')
    with sameCaseFile:
        write = csv.writer(sameCaseFile)
        write.writerows(list_MTBR_sameCase)
def checkIfFileExists(fileName)->str:
    counter = 0
    newFileName = fileName
    while exists(newFileName):
        counter+=1
        spl = fileName.split(".csv")
        newFileName = spl[0]+"("+str(counter)+").csv"
    return newFileName
def checkIfFolderExists(folderName)->str:
    counter = 0
    newFolderName = folderName
    while exists(newFolderName):
        counter+=1
        # spl = folderName.split(".csv")
        newFolderName = folderName + "("+str(counter)+")"
    return newFolderName

#remaining layout
directionsLabel = tk.Label(root, text="Directions:")
directionsLabel.grid(row=0, column=0, sticky='w')
directionsLabel_csv = tk.Label(root, text="(export each file as a .csv)")
directionsLabel_csv.grid(row=1, column=0, sticky='w')
directionsLabel_AT = tk.Label(root, text="Pull Assets report from here (all-time range report):  ")
directionsLabel_AT.grid(row=2, column=0, sticky='e')
directionsLabel_CT = tk.Label(root, text="Pull Contracts report from here (all-time range report):  ")
directionsLabel_CT.grid(row=3, column=0, sticky='e')
directionsLabel_CS = tk.Label(root, text="Pull Cases report from here (set desired range):  ")
directionsLabel_CS.grid(row=4, column=0, sticky='e')
directionsLabel_WO = tk.Label(root, text="Pull Work Order report from here (match range of cases):  ")
directionsLabel_WO.grid(row=5, column=0, sticky='e')
directionsLabel_TS = tk.Label(root, text="Pull Timesheets report from here (match WO range, validate any >15 hrs, ensure no ts after current day):  ")
directionsLabel_TS.grid(row=6, column=0, sticky='e')

label_AT = tk.Label(root, text="LINK", fg="blue", cursor="hand2")
label_AT.grid(row=2, column=1, sticky='w')
label_AT.bind("<Button-1>", lambda e: wb.open_new(link_AT))

label_CT = tk.Label(root, text="LINK", fg="blue", cursor="hand2")
label_CT.grid(row=3, column=1, sticky='w')
label_CT.bind("<Button-1>", lambda e: wb.open_new(link_CT))

label_CS = tk.Label(root, text="LINK", fg="blue", cursor="hand2")
label_CS.grid(row=4, column=1, sticky='w')
label_CS.bind("<Button-1>", lambda e: wb.open_new(link_CS))

label_WO = tk.Label(root, text="LINK", fg="blue", cursor="hand2")
label_WO.grid(row=5, column=1, sticky='w')
label_WO.bind("<Button-1>", lambda e: wb.open_new(link_WO))

label_TS = tk.Label(root, text="LINK", fg="blue", cursor="hand2")
label_TS.grid(row=6, column=1, sticky='w')
label_TS.bind("<Button-1>", lambda e: wb.open_new(link_TS))

root.grid_rowconfigure(7, minsize=50) #set blank size of empty grid row


if flag_quickLoad:
    directionsLabel_load = tk.Label(root, text="Ensure .csv files are within the same folder as the .py script with the following names:")
    directionsLabel_load.grid(row=7, column=0, sticky='w')

    root.grid_rowconfigure(8, minsize=10) #set blank size of empty grid row
    directionsLabel_load = tk.Label(root, text="assets.csv")
    directionsLabel_load.grid(row=9, column=0, sticky='e')
    directionsLabel_load = tk.Label(root, text="contracts.csv")
    directionsLabel_load.grid(row=10, column=0, sticky='e')
    directionsLabel_load = tk.Label(root, text="cases.csv")
    directionsLabel_load.grid(row=11, column=0, sticky='e')
    directionsLabel_load = tk.Label(root, text="WOs.csv")
    directionsLabel_load.grid(row=12, column=0, sticky='e')
    directionsLabel_load = tk.Label(root, text="timesheets.csv")
    directionsLabel_load.grid(row=13, column=0, sticky='e')

    root.grid_rowconfigure(14, minsize=10) #set blank size of empty grid row
    directionsLabel_save = tk.Label(root, text="The script will output files to a folder labeled 'OUTPUT'")
    directionsLabel_save.grid(row=15, column=0, sticky='w')

    root.grid_rowconfigure(16, minsize=10) #set blank size of empty grid row

    openCTFile = tk.Button(root, text="RUN", padx=10, pady=5, foreground="white", background="gray", command=parseAssets)
    openCTFile.grid(row=17, column=0, sticky='e')
else:
    directionsLabel_load = tk.Label(root, text="Load each report using the buttons below.")
    directionsLabel_load.grid(row=7, column=0, sticky='w')
    directionsLabel_save = tk.Label(root, text="Once each report is loaded, it will ask you to save the output file somewhere.")
    directionsLabel_save.grid(row=8, column=0, sticky='w')

    openCTFile = tk.Button(root, text="1. Load Assets Report", padx=10, pady=5, foreground="white", background="gray", command=parseAssets)
    openCTFile.grid(row=9, column=0, sticky='e')
    openCTFile = tk.Button(root, text="2. Load Contracts Report", padx=10, pady=5, foreground="white", background="gray", command=parseContracts)
    openCTFile.grid(row=10, column=0, sticky='e')
    openCSFile = tk.Button(root, text="3. Load Cases Report", padx=10, pady=5, foreground="white", background="gray", command=parseParentCases)
    openCSFile.grid(row=11, column=0, sticky='e')
    openWOFile = tk.Button(root, text="4. Load WO Report", padx=10, pady=5, foreground="white", background="gray", command=parseWorkOrders)
    openWOFile.grid(row=12, column=0, sticky='e')
    openTimesheetFile = tk.Button(root, text="5. Load Timesheets Report", padx=10, pady=5, foreground="white", background="gray", command=parseTimeSheets)
    openTimesheetFile.grid(row=13, column=0, sticky='e')

    label_loaded_AT = tk.Label(root, textvariable=text_loaded_AT, fg="green")
    label_loaded_AT.grid(row=9, column=1, sticky='w')
    label_loaded_CT = tk.Label(root, textvariable=text_loaded_CT, fg="green")
    label_loaded_CT.grid(row=10, column=1, sticky='w')
    label_loaded_CS = tk.Label(root, textvariable=text_loaded_CS, fg="green")
    label_loaded_CS.grid(row=11, column=1, sticky='w')
    label_loaded_WO = tk.Label(root, textvariable=text_loaded_WO, fg="green")
    label_loaded_WO.grid(row=12, column=1, sticky='w')
    label_loaded_TS = tk.Label(root, textvariable=text_loaded_TS, fg="green")
    label_loaded_TS.grid(row=13, column=1, sticky='w')

    root.grid_rowconfigure(14, minsize=50) #set blank size of empty grid row

    # resetButton = tk.Button(root, text="Reset", padx=10, pady=5, foreground="white", background="gray", command=resetApp)
    # resetButton.grid(row=14, column=0, sticky='e')

root.mainloop()