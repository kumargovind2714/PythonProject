''''
File Name      : Config_Reader.py
Author Name    : Lakshmi Damodara
Date           : 01/20/2018
Version        : one
Description    :
This program is a basic function file to return the database related information like
Hostname, UserName, Password, Database

The hard-coded value is the directory and the file of the dbase_Config.ini file.
The config file should be in the same directory as this file, if it is not, provide the right directory paths.

'''

# Library configparser is used to parse the config file
import configparser

# creating an instance of the configparser
config = configparser.ConfigParser()

# hard-coded config file to be read by this porgram
config.read('D:\Anaconda3\Kumar\dbBase_Config.ini')

# returns the hostname of the database
def returnHostName():
    return config['Database2']['hostname']

# returns the user name of the database
def usrName():
    return config['Database2']['username']

# returns the password of the database
def PWD():
    return config['Database2']['password']

# returns the database name of the database
def dbName():
    return config['Database2']['database']

# -- End of Program ---