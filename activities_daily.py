'''
File Name      : activities_table_data.py
Author Name    : Lakshmi Damodara
Date           : 01/24/2018
Version        : 1.0
Description    :
This program reads the mechanical tracker file and gets the values for the activities table
The hard-coded value is the directory and the file of the excel_act_tble_config.ini file.
The cell positions are given in the config file.

The program reads the various attributes of the cell for activity columns and writes to output file
using the program excel_writing.py

Functions:
converDate, convertDate1, getActivityNameCellPosition, getUnitNameCellPosition, getContractorNameCellPosition
getPlannedStartCellPosition, getPlannedEndCellPosition

Files need to run this program:
1. excel_act_tble_config.ini
2. excel_file_config.ini

Program dependencies:
1. excel_file_config_reader.py
2. excel_writing.py
3. excel_utilities.py

Log File
1. log_file.txt : has been set at the DEBUG level to log all activities for the run time.

'''

import openpyxl
import configparser
import datetime
import logging

import excel_writing as gwWriter
import excel_utilities as eu
import Database_Insert as dti

### -- Start of Functions --------
# function to convert dates to string in mmddyyyy format
def convertDate(dtt):
    return datetime.datetime.date(dtt) # returns just the date in mm-dd-yyyy format

def convertDate1(dtt):
    return dtt.strftime('%m%d%Y') # returns date in string in mmddyyyy format

def getActivityNameCellPosition(pos): #returns the value of activities_name from excel_act_tble_config.ini
    keyVal = 'activities' + str(pos)
    cell_postion = config1[keyVal]['activities_name']
    return cell_postion

#returns the value of unit name from excel_act_tble_config.ini
def getUnitNameCellPosition(pos):
    keyVal = 'activities' + str(pos)
    cell_postion = config1[keyVal]['activities_unit_name']
    return cell_postion

#returns the value of contractor name from excel_act_tble_config.ini
def getContractorNameCellPosition(pos):
    keyVal = 'activities' + str(pos)
    cell_postion = config1[keyVal]['activities_contractor_name']
    return cell_postion

#returns the value of planned start date from excel_act_tble_config.ini
def getPlannedStartCellPosition(pos):
    keyVal = 'activities' + str(pos)
    cell_postion = config1[keyVal]['activities_planned_start']
    return cell_postion

#returns the value of planned end date from excel_act_tble_config.ini
def getPlannedEndCellPosition(pos):
    keyVal = 'activities' + str(pos)
    cell_postion = config1[keyVal]['activities_planned_end']
    return cell_postion

#function to get output file name for writing out the results csv file
def outfile():
    return config1['outputFileName']['fname']

#function to get output directory file name for writing out the results csv file
def outfileDir():
    return config1['outputFileName']['fdirectory']

# returns incremented value of a variable
def incrementfnc(tval):
    tval = tval + 1
    return tval

def prResultData(rList):
    print(rList)

def extractActualStartDate(rList):
     data = rList[0][1]
     L_actual_start_date.append(data)
     return data

def extractActualEndDate(rList):
        leng_list = len(rList)
        data = rList[leng_list -1][1]
        L_actual_end_date.append(data)
        return data

### ---------End of Functions -----

# import excel_file_config_readyer.py to get all its functions
import excel_file_config_reader as efcr_activities

# open the config parser to read the activities config file
config1 = configparser.ConfigParser()
config1.read('..\excel_act_tble_config.ini')
#config1.read('..\excel_activities_config.in')

# get the filename and directory for logfile writing
Log_FileName = efcr_activities.logfileDirectory() + efcr_activities.logfileName() # directory + filename
#print(Log_FileName)
logging.basicConfig(filename=Log_FileName,level=logging.DEBUG,)

logging.debug('Program: activities_table_data.py........')

# get the filename and directory : Excel fileName and directory for reading values
L_FileName = efcr_activities.fileDirectory() + efcr_activities.fileName() # directory + filename
logging.debug('activities_table_data.py : opening excel file name - %s'%L_FileName)
# passing the file name and creating an instance of the workbook
act_wb = openpyxl.load_workbook(L_FileName,data_only='True')

# getting the active worksheet
wrksheet_names = act_wb.sheetnames

##########-----------------------------------------#############
L_result_data_sheet = []
L_actual_start_date = []
L_actual_end_date = []

asheets = efcr_activities.getActivitySheets() # get the list of activity sheets from excel_file_config.ini
len_asheets = len((asheets))
for i in range(0,len_asheets,2):
    sheet_val = asheets[i] # getting the active worksheet number
    L_result_data_sheet = eu.getSheetResult(act_wb,sheet_val) # calling function getSheetResult()
    L_actualStart_date = extractActualStartDate(L_result_data_sheet)
    L_actualEnd_date = extractActualEndDate(L_result_data_sheet)

#########-------------------------------------------#############

#get the total activities in the sheet
tot_activity = config1['TotalActivities']['total_activities']

# initializing a list
Final_List = list()

# This for loop is to go through the excel sheet
# Take key values of excel_act_tble_config.ini as arguments
# search each cell to get the values
logging.debug('Entering into For loop to get values from excel sheet')

# get the active sheet
activityName_active_sheet = wrksheet_names[0]
# pass the active sheet name
sheet = act_wb[activityName_active_sheet]

L1 = []

for i in range(0,int(tot_activity)):
    L_activityName_cell_value = sheet[getActivityNameCellPosition(i)]
    L_activities_unit_name_cell_value = sheet[getUnitNameCellPosition(i)]
    L_activities_contractor_name_cell_value = sheet[getContractorNameCellPosition(i)]
    LC_activities_planned_start_cell_value = sheet[getPlannedStartCellPosition(i)]
    L_activities_planned_start_date = convertDate(LC_activities_planned_start_cell_value.value).strftime('%m%d%Y')
    LC_activities_planned_end_cell_value = sheet[getPlannedEndCellPosition(i)]
    L_activities_planned_end_date = convertDate(LC_activities_planned_end_cell_value.value).strftime('%m%d%Y')
    L_activities_actual_start_date = L_actual_start_date[i]
    L_activities_actual_end_date = L_actual_end_date[i]

    # Depending on the number of activities, the if loop will load the list
    j = i - 1
    L1.insert(j,L_activityName_cell_value.value)
    L1.insert(incrementfnc(j+1),L_activities_unit_name_cell_value.value)
    L1.insert(incrementfnc(j+2),L_activities_contractor_name_cell_value.value)
    L1.insert(incrementfnc(j+3),L_activities_planned_start_date)
    L1.insert(incrementfnc(j+4),L_activities_planned_end_date)
    L1.insert(incrementfnc(j+5),L_activities_actual_start_date)
    L1.insert(incrementfnc(j+6), L_activities_actual_end_date)
    final_list = [L1]

    # output file
    output_FileName1 = outfileDir() + str(outfile())
    output_FileName = output_FileName1.replace("'","")
    logging.debug('activities_table_data.py : sending the list to excel_writing.py file ')
    # Now pass the list along with filename to the writer python file
    gwWriter.writeCSVFile(output_FileName,final_list,Log_FileName)
    dti.executeSQL_Activities(final_list)
    L1 = []

# close or delete all the open instances, Lists, and connections
# clears all the variables from memory

del final_list
del L1
del L_result_data_sheet
del L_actual_start_date
del L_actual_end_date
act_wb.close()
config1.clear()
del efcr_activities
del eu
del dti

# --- End of Program ---