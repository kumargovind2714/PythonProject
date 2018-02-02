
'''
File Name      : excel_writing.py
Author Name    : Lakshmi Damodara
Date           : 01/24/2018
Version        : 1.0
Description    :
This file is primarily to write the data to a csv file with comma delimiters
It accepts the name of the file with directory and list as an argument
Then it opens the file with append mode and writes to the csv file
'''

# import the csv library package
import csv
import logging


# function to write the data into the csv file.
def writeCSVFile(oFile, listOfListVal,logFname):
    logging.info('excel_writing.py : Opening the csv file for writing the output %s' %oFile)
    with open(oFile, 'a') as outcsv:
        #configure writer to write standard csv file
        writer = csv.writer(outcsv, delimiter=",", lineterminator='\n')
        print(listOfListVal)
        Len_listOfListVal = len(listOfListVal)
        for list in range(0,Len_listOfListVal):
            writer.writerow(listOfListVal[list])

    outcsv.close()


# function to write the data into the csv file.
def write_activity_daily_data_CSV(FileN, rslt):
    logging.info('excel_writing.py : Opening the csv file for writing the output %s' %FileN)
    with open(FileN, 'a') as outcsv:
        #configure writer to write standard csv file
        writer = csv.writer(outcsv, delimiter=',', quotechar='|', quoting=csv.QUOTE_MINIMAL, lineterminator='\n')
        writer.writerows(rslt)

    outcsv.close()
#---- End of program