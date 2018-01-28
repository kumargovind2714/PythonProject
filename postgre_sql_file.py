''''
File Name      : postgre_sql_file.py
Author Name    : Lakshmi Damodara
Date           : 01/20/2018
Version        : one
Description    :
1. This program gets the necessary connection details about the postgres database from Config_Reader.py
2. Establishes the connection to the postgres database
3. Runs the query and gets all the values and displays on the screen.
4. Catches the exception if any database error

'''

# Loading the library for postgres sql
import psycopg2
# Loading the config_Reader program to get the database connection details
import Config_Reader as configFile

# Start function: doQuery(): to run a simple query
def doQuery( conn ):
    # creating a cursor to read the records
    cur = conn.cursor()
    # Query
    cur.execute("SELECT id, name FROM units")
    # Parsing through the cursor and printing the details on the screen
    for id, name in cur.fetchall():
        print(id, name)
# -- End of function doQuery()

try :
    ## creating a connection instance
    myConnection = psycopg2.connect(host=configFile.returnHostName(),user=configFile.usrName(),password=configFile.PWD(),database=configFile.dbName())
    # call the query
    doQuery(myConnection)
except (Exception, psycopg2.DatabaseError) as error:
        print(error)
finally:
    if myConnection is not None:
            myConnection.close()
#-- End of Program --