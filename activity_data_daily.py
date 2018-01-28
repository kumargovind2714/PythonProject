'''
File Name      : activity_data_daily.py
Author Name    : Lakshmi Damodara
Date           : 01/26/2018
Version        : 1.0
Description    :
This program reads the individual sheets in mechanical tracker file and gets the values for the daily activity Data table
The hard-coded value is the directory and the file of the excel_act_tble_config.ini file.
The cell positions are given in the config file.

The program reads the various attributes of the cell for activity columns and writes to output file
using the program excel_writing.py

Functions:
converDate, convertDate1

Files need to run this program:
1. excel_act_tble_config.ini
2. excel_file_config.ini

Program dependencies:
1. excel_file_config_reader.py
2. excel_writing.py

Log File
1. log_file.txt : has been set at the DEBUG level to log all activities for the run time.

'''

import openpyxl
import configparser
import datetime
import logging


import excel_writing as ewWriter

### -- Start of Functions --------
# function to convert dates to string in mmddyyyy format
def convertDate(dtt):
    return datetime.datetime.date(dtt) # returns just the date in mm-dd-yyyy format

def convertDate1(dtt):
    return dtt.strftime('%m%d%Y') # returns date in string in mmddyyyy format

# returns incremented value of a variable
def incrementfnc(tval):
    tval = tval + 1
    return tval

def remove_items_list(listVar,removeListVar):
    # remove the unnecessary values from the result data
    for i in sorted(removeListVar, reverse=True):
        del listVar[i]
    return listVar

def getSheetResult(wbe, sheet_name):
    # get the max count of rows and cols
    m_row = sheet_name.max_row
    m_col = sheet_name.max_column
    m_row = m_row + 1
    #print(m_row, m_col)

    # initialize list
    result_data = []
    for curr_row in range(7, m_row, 1):
        if not sheet_name.row_dimensions[curr_row].hidden == True:  # dont read if the row is hidden
            row_data = []
            for curr_col in range(2, 8, 1):
                # print('I am in row :%d' %curr_row)
                data = sheet_name.cell(row=curr_row, column=curr_col)
                # print(len(data.value))
                if isinstance(data.value, datetime.datetime):
                    row_data.append(convertDate(data.value).strftime('%m%d%Y'))
                else:
                    row_data.append(data.value)
        result_data.append(row_data)

    popping_Var = []
    LC_pop_len = len(result_data)

    ## Now accessing the list of list result_data : result_data = [][]
    for i in range(0, LC_pop_len, 1):
        for j in range(0, 6, 1):
            LC_data = result_data[i][j]
            if type(LC_data) == str:
                if len(LC_data) > 10:
                    popping_Var.append(i)
                break

    # Call the function to remove the list containing None or null values
    result_data = remove_items_list(result_data, popping_Var)
    #print(result_data)

### -------------------------------------
### Remove the list of result Data which has None or null values
### --------------------------------------
    popping_Var_None = []
    LC_pop_len_none = len(result_data)
    ## Now accessing the list of list result_data with None Value
    for i in range(0, LC_pop_len_none, 1):
        LC_data = result_data[i][5]
        if LC_data == None:
            popping_Var_None.append(i)

    # Call the function to remove the list containing None or null values
    result_data=remove_items_list(result_data,popping_Var_None)

    print(result_data)
    return result_data
### ---------End of Functions -----

# import excel_file_config_readyer.py to get all its functions
import excel_file_config_reader as efcr

# get the filename and directory for logfile writing
Log_FName = efcr.logfileName()
Log_Dname = efcr.logfileDirectory()
Log_FileName = Log_Dname + Log_FName # directory + filename
print(Log_FileName)
logging.basicConfig(filename=Log_FileName,level=logging.DEBUG,)

logging.info('Program: activities_data_daily.py........')

# get the filename and directory : Excel fileName and directory for reading values
L_FName = efcr.fileName()
L_Dname = efcr.fileDirectory()
L_FileName = L_Dname + L_FName # directory + filename
logging.info('activities_data_daily.py : opening excel file name - %s'%L_FileName)
# passing the file name and creating an instance of the workbook
wb = openpyxl.load_workbook(L_FileName,data_only='True')

#Reading for each sheet
# getting the required sheet
result_data_sheet1 = []
result_data_sheet2 = []
result_data_sheet3 = []

sheet = wb['Pile Installation']
result_data_sheet1 = getSheetResult(wb,sheet)
ewWriter.write_activity_daily_data_CSV('..\output\daily_PileInstallation_activity_data.csv', result_data_sheet1,'..\log_activity_data.txt')
print('Finished Writing Pile Installation :%d' %len(result_data_sheet1))

sheet = wb['Torque Tube Rows Installed']
result_data_sheet2 = getSheetResult(wb,sheet)
ewWriter.write_activity_daily_data_CSV('..\output\daily_TorqueTubeRowInstalled_activity_data.csv', result_data_sheet2,'..\log_activity_data.txt')
print('Finished writing TorqueTubeRowInstalled :%d' %len(result_data_sheet2))

sheet = wb['Tracker Rows Completed']
result_data_sheet3 = getSheetResult(wb,sheet)
ewWriter.write_activity_daily_data_CSV('..\output\daily_TorqueTubeRow_activity_data.csv', result_data_sheet3,'..\log_activity_data.txt')
print('Finished Writing TrackerRowsCompleted :%d' %len(result_data_sheet3))



#---- End of Program ------