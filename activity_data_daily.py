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

#---------------------------------------------------------------------------
# This function contains the business logic to read the daily activity sheet
#----------------------------------------------------------------------------
def getSheetResult(wbe, sheet_name):
    # get the max count of rows and cols
    m_row = sheet_name.max_row
    m_col = sheet_name.max_column
    m_row = m_row + 1
    #print(m_row, m_col)

    # -- This section is to store the data of the active sheet in to List of Lists
    # -- result_data is the list of list containing : rows and columns of the active sheet
    # -- The reading of row starts at Row# 7 and 6 columns are read starting from column # 2
    # -- As it is being read, the dates are converted to MMDDYYYY format and
    # -- the result is stored in result_data as list of lists

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

    # -- This section is to create a list - popping_Var which contains the row index of result_data
    # -- that needed to be removed as it contains non-date values
    # -- We are reading the col.#2 and checking if the string value is greater than 10
    # -- If yes, then the index value is stored in popping_Var

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

    # Call the function to remove the list of values in result_data which are referenced in popping_Var
    result_data = remove_items_list(result_data, popping_Var)

### -------------------------------------
### Remove the list of result Data which has None or null values
### If the planned to date column has null values, those rows in the result_data list are removed.
### --------------------------------------
    popping_Var_None = []
    LC_pop_len_none = len(result_data)
    ## Now accessing the list of list result_data with None Value
    for i in range(0, LC_pop_len_none, 1):
        LC_data = result_data[i][5]
        if LC_data == None: # checking if the list index[5] in result_data is None
            popping_Var_None.append(i) # store the index value in popping_Var_None list

    # Call the function to remove the list containing None or null values as referenced in popping_Var_None
    result_data=remove_items_list(result_data,popping_Var_None)

    # return the final list to the calling function
    print(result_data)
    return result_data

### ---------End of Functions -----

# import excel_file_config_readyer.py to get all its functions
import excel_file_config_reader as efcr

# get the filename and directory for logfile writing
Log_FName = efcr.logfileName()
Log_Dname = efcr.logfileDirectory()
Log_FileName = Log_Dname + Log_FName # directory + filename
#print(Log_FileName)
logging.basicConfig(filename=Log_FileName,level=logging.DEBUG,)

logging.info('Program: activities_data_daily.py........')

# get the filename and directory : Excel fileName and directory for reading values
L_FName = efcr.fileName()
L_Dname = efcr.fileDirectory()
L_FileName = L_Dname + L_FName # directory + filename
logging.info('activities_data_daily.py : opening excel file name - %s'%L_FileName)

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
    sheet_val = asheets[i]
    sheet = efcr.shName(sheet_val) # get the sheet name
    Asheet = wb[sheet] # creating an instance of the sheet
    result_data_sheet = getSheetResult(wb, Asheet) # getting the values of the sheet in
    outDir = efcr.outputDirectory() # getting the directory for writing the output file
    loc_fname = outDir + sheet + '.csv' # creating the outputfile name with directory,sheetname
    print(loc_fname)
    # calling the excel_writing.py to write the data to the file
    ewWriter.write_activity_daily_data_CSV(loc_fname, result_data_sheet,'..\log_activity_data.txt')

#---- End of Program ------