# -*- coding: utf-8 -*-
"""
Created on Thu Mar  5 15:24:34 2020

@author: EKKORIR
"""

import pandas as pd
import datetime as dt
import os
import glob
import shutil
from openpyxl import load_workbook
import re

#determine the correct way to pick the date
dateTimeObj = dt.datetime.today()
# get the date object from datetime object
dateObj = dateTimeObj.date()
if dateObj.day < 10:
    day_date = str(0)+str(dateObj.day)+str('-')
else:
    day_date = str(dateObj.day)+str('-')
if dateObj.month < 10:
    month_date = str(0)+str(dateObj.month)+str('-')
else:
    month_date = str(dateObj.month)+str('-')
today_date = day_date+month_date+str(dateObj.year)
today_date = dt.datetime.strptime(today_date[0:10], '%d-%m-%Y')

#bandia today date
#today_date = '04-03-2020'
#today_date = dt.datetime.strptime(today_date[0:10], '%d-%m-%Y')
#today_date

#redirect python to operate from the desktop folder
os.chdir(r'C:\Users\EKKORIR\Desktop\Operational Excellence\Reports\Degredation Report')
fileName = glob.glob('Degradation_Report_*.xlsx')
print('Identified files:', fileName)

# algorithm to extract the old file by file name
for f in range(len(fileName)):
    date_val = fileName[f].split("_")
    date_value = dt.datetime.strptime(date_val[2][0:10], '%d-%m-%Y')
    day_difference = abs(today_date - date_value).days
    if (day_difference > 1) and (day_difference<=7):
        correct_file = fileName[f]
        print('Last report file: ',correct_file)
        
past_report = pd.read_excel(correct_file,sheet_name='Degradation',index_col=None,header=3)

#loop through and determine the incident that meets the specified criteria
reported_links = []
for i in range(len(past_report)):
    if (past_report.iloc[i][7] == 'Closed'): #and (dateObj.day - past_report.iloc[i][2].day) <= 4:
        #print (past_report.iloc[i][0],'----',past_report.iloc[i][11],'----',past_report.iloc[i][14],'----',past_report.iloc[i][16])
        reported_links.append(past_report.iloc[i,14])
reported_links = list(set(reported_links))
reported_links

# Move the old file
destination = r'C:\Users\EKKORIR\Desktop\Operational Excellence\Reports\Degredation Report\OldReport'
source = r'C:\Users\EKKORIR\Desktop\Operational Excellence\Reports\Degredation Report'
print(os.listdir(destination))
print(os.listdir(source))
dest = shutil.move(source+'\\'+correct_file,destination)
print(os.listdir(destination))
print(os.listdir(source))

# Get the latest file
# redirect python to operate from the desktop folder
os.chdir(r'C:\Users\EKKORIR\Desktop\Operational Excellence\Reports\Degredation Report')
fileName2 = glob.glob('Degradation_Report_*.xlsx')
fileName2[0]

current_report = pd.read_excel(fileName2[0],sheet_name='Degradation',index_col=None,header=3)
#loop through and determine the incident that meets the specified criteria
newly_reported_links = []
for i in range(len(current_report)):
    if (current_report.iloc[i][7] != 'Closed'): #and (dateObj.day - past_report.iloc[i][2].day) <= 4:
        #print (current_report.iloc[i][0],'----',current_report.iloc[i][1],'----',current_report.iloc[i][2],'----',current_report.iloc[i][3])
        newly_reported_links.append(current_report.iloc[i,14])
newly_reported_links = list(set(newly_reported_links))
newly_reported_links

#Determine recurred links
def intersection(list1, list2):
    return list(set(list1) & set(list2))

recurred_links = intersection(reported_links, newly_reported_links)

# Remove nan from the list
semifinal_recurred_list = [r for r in recurred_links if r == r]

# Match out only links from the list
final_recurred_list = []
for x in range(len(semifinal_recurred_list)):
    if re.search('GigabitEthernet',semifinal_recurred_list[x]):
        print(re.search('GigabitEthernet',semifinal_recurred_list[x]).string)
        final_recurred_list.append(re.search('GigabitEthernet',semifinal_recurred_list[x]).string)

final_recurred_list

# create new dataframe and rename columns
updated_report = current_report
updated_report.columns = ['Incident ID', 'Last Name', 'First Name', 'Department', 'Summary',
       'Service*+', 'Priority', 'Status', 'Assigned Group', 'Assignee',
       'Reported Date+', 'Last Resolved Date', 'SLM Real Time Status',
       'Resolution', 'Degraded Links', 'Incident Type*', 'Responded Date+',
       'Target Date', 'QKE-CauseOfFailureValue']

#Write to excel file
writer = pd.ExcelWriter("Degradation_Publishing_Report.xlsx", engine="openpyxl")
writer.book = load_workbook('Degradation_Publishing_Report.xlsx')
writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
reader = pd.read_excel(r'Degradation_Publishing_Report.xlsx')
for k in final_recurred_list:
    
    df = pd.DataFrame(updated_report[updated_report['Degraded Links'] == k][['Last Resolved Date','Degraded Links','Incident ID']])
    df['Week No.'] = dt.datetime.today().isocalendar()[1]
    df['Region_initials'] = k.split("_")[1]
    df['Region'] = 'Region Unspecified'
    if df['Region_initials'].values[0] == 'NE':
        df['Region'] == 'NARIOBI EAST'
    elif df['Region_initials'].values[0] == 'NW':
        df['Region'] == 'NARIOBI WEST'
    elif df['Region_initials'].values[0] == 'RV':
        df['Region'] == 'RIFT VALLEY'
    elif df['Region_initials'].values[0] == 'WN':
        df['Region'] == 'WESTERN NYANZA'
    elif df['Region_initials'].values[0] == 'CO':
        df['Region'] == 'COAST'
    elif df['Region_initials'].values[0] == 'MK':
        df['Region'] == 'MOUNT KENYA'
    del df['Region_initials']
    df.to_excel(writer,index=False,header=True)   
    writer.close()