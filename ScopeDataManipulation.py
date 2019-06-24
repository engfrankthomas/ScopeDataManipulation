"""
""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
Code name:      Scope data manipulation script
Requirement:    Python 3
Description:    This script receive data collected by UNHCR and connert them into a format compatible with Scope.
Author:         WFP Tanzania/FT
""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
"""

#python packages
import subprocess
import xml.etree.cElementTree as et
import pandas as pd
from collections import OrderedDict
from pandas import ExcelWriter
from pandas import ExcelFile
import numpy as np
import datetime as dt
import dateutil
import tkinter as tk
import tkinter.filedialog as tkf
import os
application_window = tk.Tk()
application_window.withdraw()

#acquire file path
data_path = tkf.askopenfilename(parent=application_window, initialdir=os.getcwd(),title="Please select a excel file with data", filetypes=[('xml file', '.xml')])

#import XML file.
import xml.etree.ElementTree as ET
import pandas as pd
xml_final_1 = data_path
tree = ET.parse(xml_final_1)
root = tree.getroot()
dfcols = ['Given_name_of_HR1','Middle_name_of_HR1','Family_name_of_HR1','Household_Number', 'Age_of_HR1', 'Sex_of_HR1','Address_Location_Level_4', 'Photo_of_HR1', 'Feeding_Family_size_6_months_and_above']
data_table = pd.DataFrame(columns=dfcols)
for record in tree.findall(".//record"):
    Given_name_of_HR1 = record.find("./Given_name_of_HR1")
    Middle_name_of_HR1 = record.find("./Middle_name_of_HR1")
    Family_name_of_HR1 = record.find("./Family_name_of_HR1")
    Household_Number = record.find("./Household_Number")
    Household_Size = record.find("./Household_size")
    Individual_Number_of_HR1 = record.find("./Individual_Number_of_HR1")
    Age_of_HR1 = record.find("./Age_of_HR1")
    Sex_of_HR1 = record.find("./Sex_of_HR1")
    Address_Location_Level_4 = record.find("./Address_Location_Level_4")
    Photo_of_HR1 = record.find("./Photo_of_HR1")
    Feeding_Family_size_6_months_and_above = record.find("./Feeding_Family_size_6_months_and_above")
    if Household_Number is not None:
        data_table = data_table.append(pd.Series([Given_name_of_HR1.text, Middle_name_of_HR1.text ,Family_name_of_HR1.text, Household_Number.text, Age_of_HR1.text, Sex_of_HR1.text, Address_Location_Level_4.text, Photo_of_HR1.text, Feeding_Family_size_6_months_and_above.text], index=dfcols), ignore_index=True)
#print(data_table['Age_of_Designate_-if_applicable'])
data_table['Location']=data_table.apply(lambda x: 'NY FDP 1', axis=1)
data_table['Document Type']=data_table.apply(lambda x: 'Refugee Proof of Registration', axis=1)
data_table['Household Role']=data_table.apply(lambda x: 'HD', axis=1)
data_table['Recipient']=data_table.apply(lambda x: 'P', axis=1)
data_table['Sex_of_HR1']=data_table['Sex_of_HR1'].replace({'Male':'M','Female':'F'})
data_table['Marital Status']=data_table.apply(lambda x: ' ', axis=1)
data_table['Date of Birth']=data_table.apply(lambda x: ' ', axis=1)
data_table['Phone Number']=data_table.apply(lambda x: ' ', axis=1)
data_table['Mobile Number']=data_table.apply(lambda x: ' ', axis=1)
data_table['Bank Account Number']=data_table.apply(lambda x: ' ', axis=1)
data_table['Household Vulnerable']=data_table.apply(lambda x: ' ', axis=1)
data_table['Language Code']=data_table.apply(lambda x: ' ', axis=1)
data_table['E Card Status']=data_table.apply(lambda x: ' ', axis=1)
data_table['Physical Disability Status']=data_table.apply(lambda x: ' ', axis=1)
data_table['Mental Disability Status']=data_table.apply(lambda x: ' ', axis=1)
data_table['Orphan']=data_table.apply(lambda x: ' ', axis=1)
data_table['Height']=data_table.apply(lambda x: ' ', axis=1)
data_table['Weight']=data_table.apply(lambda x: ' ', axis=1)
data_table['MUAC']=data_table.apply(lambda x: ' ', axis=1)
data_table['Breastfeeding']=data_table.apply(lambda x: ' ', axis=1)
data_table['Malnourished']=data_table.apply(lambda x: ' ', axis=1)
data_table['Pregnant']=data_table.apply(lambda x: ' ', axis=1)
data_table['Qualified']=data_table.apply(lambda x: ' ', axis=1)
data_table['Pregnancy Due Date']=data_table.apply(lambda x: ' ', axis=1)
data_table['Child Birth Date']=data_table.apply(lambda x: ' ', axis=1)
data_table['Birth Certificate Provided']=data_table.apply(lambda x: ' ', axis=1)
data_table['Household Arrival Date']=data_table.apply(lambda x: ' ', axis=1)

#renaming headers in dataframe
data_table.rename(columns={
    'Sex_of_HR1':'Gender',
    'Address_Location_Level_4':'Address',
    'Age_of_HR1':'Age',
    'Family_name_of_HR1':'Last Name',
    'Feeding_Family_size_6_months_and_above':'Household Size',
    'Given_name_of_HR1':'First Name',
    'Household_Number':'Household Name',
    'Middle_name_of_HR1':'Middle Name',
    'Photo_of_HR1':'photo_data',}, inplace=True)
data_table['Document Num'] = data_table['Household Name']

scope_data = data_table[['Household Name','Location','Address','Household Size','Document Type','Document Num','Last Name', 'First Name', 'Middle Name', 'Household Role', 'Recipient', 'Gender', 'Marital Status', 'Date of Birth','Age','Phone Number','Mobile Number','Bank Account Number','Household Vulnerable', 'Language Code','E Card Status', 'Physical Disability Status', 'Mental Disability Status', 'Orphan', 'Height', 'Weight','MUAC', 'Breastfeeding', 'Malnourished', 'Pregnant', 'Qualified', 'photo_data', 'Pregnancy Due Date', 'Child Birth Date', 'Birth Certificate Provided', 'Household Arrival Date']]

#export file
export_file_name = os.path.splitext(os.path.basename(data_path))[0]
export_file_path = 'export/%s.xlsx' %export_file_name

#write result to a export file, truncate if it exist and has something
scope_data.to_excel(export_file_path, sheet_name='Data', index = False, header= True)

#open destination folder.
subprocess.Popen(r'explorer /select,"export"')
