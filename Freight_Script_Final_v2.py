# -*- coding: utf-8 -*-
"""
Created on Thu May 6 10:14:01 2021

@author: Kenneth Liang
Script to consolidate Freight files and upload to Google Sheets

Packages that need to be installed
pip install --upgrade pip
pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
pip install xlrd
pip install gspread
pip install gspread_dataframe
pip install pandas
pip install pywin32
pip install PySimpleGUI
pip install pyinstaller

virtualenv venv
venv\Scripts\activate
pyinstaller --clean --onefile --hidden-import "pywin32" -w 'Freight_Script_Final.py'
deactivate

"""
# For Google Sheet authentication
from __future__ import print_function
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

# upload results to Google Sheets
import gspread
import gspread_dataframe as gd
import gspread.auth as ga

# Library for moving files into proper directories
import os
import shutil
import sys

import pandas as pd

# Library for sending email through outlook, need the pywin32 library
# import win32com.client as win32

# For GUI
import PySimpleGUI as sim_gui

# determine if application is a script file or frozen exe
if getattr(sys, 'frozen', False):
    directory = os.path.dirname(os.path.realpath(sys.executable))
elif __file__:
    directory = os.path.dirname(os.path.realpath(__file__))

try:
    csv_file_list = os.listdir(os.path.join(directory,'Freight Files','Working','CVS',''))
    xls_file_list = os.listdir(os.path.join(directory,'Freight Files','Working','XLS',''))

except:
    sim_gui.popup('Working and Archive folders are in the wrong directory. Please ensure they are with the Freight script.')
    
# initialize pandas dataframes
df_csv = pd.DataFrame()
df_xls = pd.DataFrame()

if len(csv_file_list)==0 or len(xls_file_list)==0:
    
    sim_gui.popup('There are no files in the CSV or XLS working directories. Please check folders.')

else:
    
    # Message to user that the script is running
    layout = [[sim_gui.Text('Script is running please wait')]]

    # create the window
    window = sim_gui.Window('',layout,keep_on_top=True)
    
    event, values = window.read(timeout=1)
    
    for file in csv_file_list:
        try:
            # read csv file, set data type to string
            data = pd.read_csv(os.path.join(directory,'Freight Files','Working','CVS',file), delimiter=';',low_memory=False, dtype=str)

            # append csv file to pandas dataframe
            df_csv = df_csv.append(data)

            # move CVS file which should have been appended to the dataframe already from Working to Archive folder
            shutil.move(os.path.join(directory,'Freight Files','Working','CVS',file), os.path.join(directory,'Freight Files','Archive','CVS',''))

        except:
            sim_gui.popup('Issue appending csv files or moving to archive folder')

    for file in xls_file_list:
        try:
            # read excel file, set data type to string
            data = pd.read_excel(os.path.join(directory,'Freight Files','Working','XLS',file), dtype=str)

            # append excel file to pandas dataframe
            df_xls = df_xls.append(data)

            # move excel file which should have been appended to the dataframe already from Working to Archive folder
            shutil.move(os.path.join(directory,'Freight Files','Working','XLS',file), os.path.join(directory,'Freight Files','Archive','XLS',''))

        except:
            sim_gui.popup('Issue appending excel files or moving to archive folder')

    try:
        # change field name so that we can join the csv and excel dataframes on invoice # and record #
        df_xls.rename(columns={'Invoice #':'INV_NUM','Record Nbr':'REC_NUM'}, inplace=True)
    except:
        sim_gui.popup('Issue renaming invoice # and record # fields in excel dataframe')

    # join the xls and csv dataframes similar to using join in SQL. join on invoice number and record number
    result = pd.merge(df_xls, df_csv, how="inner", on=['INV_NUM', 'REC_NUM'])

    # Filter and select on appropriate fields
    final = result.loc[result['Cost Object Value1']=='ENTER COST OBJECT',['INV_NUM','REC_NUM','PAYER_ACC','SHIP-FROM_COMPANY','SHIP-FROM_ADDR','SHIP-FROM_CITY','SHIP-FROM_PROV','SHIP-TO_COMPANY','SHIP-TO_ADDR','SHIP-TO_CITY','SHIP-TO_PROV','Cost Object Value1','Category','Customer Ref1','Customer Ref2','Customer Ref3','Customer Ref4','SHIPMENT_BASE_AMT','SHIPMENT_GST','SHIPMENT_PST','SHIPMENT_QST','SHIPMENT_HST','FUEL_SURCHARGE','INV_DATE']].where(result.notnull(), 0)

    # add additional table columns for total frieght cost and splitting date details
    final['BASE+FUEL COST']=final['SHIPMENT_BASE_AMT'].astype(float) + final['FUEL_SURCHARGE'].astype(float)
    final['YEAR']=final['INV_DATE'].str[:4]
    final['MONTH']=final['INV_DATE'].str[4:6]
    final['DAY']=final['INV_DATE'].str[-2:]
    
    # If modifying these scopes, delete the file token.json.
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    
    creds = None
    
    # Create Google credentials if they don't already exist
    # The file token.json stores the user's access and refresh tokens, and is created automatically when the authorization flow completes for the first time.
    if os.path.exists('.json'):
        creds = Credentials.from_authorized_user_file('.json', SCOPES)
    
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                '.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('.json', 'w') as token:
            token.write(creds.to_json())
    
    # change default directory for the credential files needed to log into Google sheets
    def gspread_paths(dir):
        ga.DEFAULT_CONFIG_DIR = dir
        ga.DEFAULT_CREDENTIALS_FILENAME = os.path.join(ga.DEFAULT_CONFIG_DIR, '.json')
        ga.DEFAULT_AUTHORIZED_USER_FILENAME = os.path.join(ga.DEFAULT_CONFIG_DIR, '.json')
        ga.load_credentials.__defaults__ = (ga.DEFAULT_AUTHORIZED_USER_FILENAME,)
        ga.store_credentials.__defaults__ = (ga.DEFAULT_AUTHORIZED_USER_FILENAME, 'token')

    gspread_paths(directory)

    # log into Google Sheets
    gs = gspread.oauth()

    # open the NMS Freight Data Google Sheet using it's key ID
    ws = gs.open_by_key('ENTER GOOGLE SHEET ID').sheet1

    # download the data in Google Sheet and append the Final dataframe to it
    existing = gd.get_as_dataframe(ws,header=0)
    existing = existing.dropna()
    updated = existing.append(final)

    # upload appended dataframe to Google Sheet
    gd.set_with_dataframe(worksheet=ws,dataframe=updated)

    window.close()