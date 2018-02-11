''''
File Name      : db_utilities.py
Author Name    : Lakshmi Damodara
Creation Date  : 01/30/2018
Updated Date   : 02/02/2018
Version        : 1.0
Description    :
1. This program is used for various database operations
2. Functions:
    executeSQL_Activities()
    executeSQL_Activities_Daily()
    executeSQL_BaseLine_Activities()
    getConn()

Program Dependencies:
1. Config_Reader.py

Configuration File:
2. dbase_Config.ini
'''

# Loading the library for postgres sql
import psycopg2

# Loading the config_Reader program to get the database connection details
import db_config_reader as configFile
import os
# function to insert values into the Activities table
# takes the list as the argument, extracts the values from the list
# inserts the values into the table

def executeQuery(sqlQuery):

    try:
        conn = getConn()
        curs = conn.cursor()
        # first delete all the existing data from activities table
        curs.execute(sqlQuery)
        conn.commit()
    except (Exception, psycopg2.DatabaseError) as error:
        print(error)

    conn.close()

def executeQueryString(sqlQuery,conn):

    try:
        curs = conn.cursor()
        # first delete all the existing data from activities table
        curs.execute(sqlQuery)
        conn.commit()
    except (Exception, psycopg2.DatabaseError) as error:
        print(error)

def executeSQL_BaseLine_Activities(exSQL,exData):
    try:
        conn = getConn()
        cur = conn.cursor()
        print(exSQL)
        print(exData)
        cur.execute(exSQL, exData)
        conn.commit()
    except (Exception, psycopg2.DatabaseError) as error:
        print(error)

    conn.close()

def executeSQLData(exSQL,exData,conn):
    try:
        cur = conn.cursor()
        print(exSQL)
        print(exData)
        cur.execute(exSQL, exData)
        conn.commit()
    except (Exception, psycopg2.DatabaseError) as error:
        print(error)


# function to get the connection values
# imports the config_Reader.py
# uses the dbase_Config.ini

def getConn():
    try :
        ## creating a connection instance
        myConnection = psycopg2.connect(host=configFile.dbHostName(),port=configFile.dbPortNumber(),user=configFile.dbUserName(),password=configFile.dbPwd(),database=configFile.dbDatabaseName())
        # call the query
        return myConnection
    except (Exception, psycopg2.DatabaseError) as error:
        print(error)

#-- End of Program --