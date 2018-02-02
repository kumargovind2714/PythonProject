'''
File Name      : activity_data_daily.py
Author Name    : Lakshmi Damodara
Date           : 01/28/2018
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
1. excel_file_config.ini

Program dependencies:
1. excel_file_config_reader.py
2. excel_writing.py

Log File
1. log_file.txt : has been set at the DEBUG level to log all activities for the run time.

Output File
1. Based on the number of activities, respective files are written in the output directory
2. The output directory can be configured on excel_file_config.ini file

'''

import openpyxl
import configparser
import datetime
import logging


import excel_writing as ewWriter
import Database_Insert as dbi
import excel_utilities as eut

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

### ---------End of Functions -----

# import excel_file_config_readyer.py to get all its functions
import excel_file_config_reader as efcr

# get the filename and directory for logfile writing
Log_FileName = efcr.logfileDirectory() + efcr.logfileName()  # directory + filename
logging.basicConfig(filename=Log_FileName,level=logging.DEBUG,)

logging.info('##---Program: activities_data_daily.py..........................')
logging.info(datetime.datetime.today())
logging.info('##-------------------- ---------------------------------........')

# get the filename and directory : Excel fileName and directory for reading values
L_FileName = efcr.fileDirectory() + efcr.fileName() # directory + filename
logging.info('activities_data_daily.py : opening excel file name - %s'%L_FileName)
print(L_FileName)
# passing the file name and creating an instance of the workbook with actual values and ignoring the formulas
wb = openpyxl.load_workbook(L_FileName,data_only='True')

# Fist get the active range sheets from excel_file_config.ini using excel_file_config_reader.py
asheets = []
result_data_sheet = []

#----------------------------------------------------------------------------------------------
# -- This section is to collect all the data from various activity sheet
# -- Call the excel writing.py program to write the file into a csv format
# -- The information about sheetname, total sheets etc are derived from excel_file_config.ini
# -- The output file is determined by the name of the sheet.csv
#----------------------------------------------------------------------------------------------

asheets = efcr.getActivitySheets() # get the list of activity sheets from excel_file_config.ini
len_asheets = len((asheets))
for i in range(0,len_asheets,2):
    sheet_val = asheets[i] # getting the active worksheet number
    result_data_sheet = eut.getSheetResult(wb,sheet_val) # calling function getSheetResult()
    loc_fname = efcr.outputDirectory() + efcr.outputfileName() # getting the output directory and filename
    # calling the excel_writing.py to write the data to the file
    ewWriter.write_activity_daily_data_CSV(loc_fname, result_data_sheet)
    dbi.executeSQL_Activities_Daily(result_data_sheet)

# close all the connections
wb.close()
# close or delete all the open instances, Lists, and connections
# clears all the variables from memory

del result_data_sheet
del asheets
del efcr
del ewWriter
del dbi
del eut

#---- End of Program ------