# this compiles the data from individual sheets into a master attendance sheet

import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import xlrd
import xlwt

scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('gspreadhw.json', scope)
client = gspread.authorize(creds)

samplefile="General_generalM.xlsx"
sheet=client.open(samplefile).sheet1
list_of_hashes=sheet.get_all_records()
Col_labels=['Dep','Loc','Act','Name','ID','Mon','Days_Attendance','Total_Sessions']
testdf=pd.DataFrame(list_of_hashes, index=None, columns=Col_labels)

sample2="General_VaniM.xlsx"
sheet2=client.open(sample2).sheet1
list_of_hash2=sheet2.get_all_records()
test2=pd.DataFrame(list_of_hash2, columns=Col_labels)

error_log=[]
for i, row in testdf.iterrows():
    if (row['Days_Attendance'] > row['Total_Sessions']):
        print("Too many Seva Days ", i),
        error_log.append(("Too Many",i))
    elif (row['Days_Attendance']==""):
        print("Seva Days Blank ", i),
        error_log.append(("Blank",i))
error_log

attend_master_df=pd.DataFrame(columns=Col_labels)
attend_master_df=attend_master_df.append(testdf)
attend_master_df=attend_master_df.append(test2)
attend_master_df.shape

writer = pd.ExcelWriter("Attendance_Master_Compile.xlsx")
attend_master_df.to_excel(writer)
writer.save()


'''
import Validate
import pandas as pd 

Month_Attendance = pd.DataFrame()

def compile_attendance(entry_form, month):
    if entry_form.Validate:
        entry_form.read(month)
        output_dataframe=entry_form.calculate_total_days_per_ID(month)
        entry_form.write_to_master(output_dataframe)  #upload to google drive

def read(entry_form, month):
    pass     #calling Sheets API and read entry_form

def calculate_total_days_per_ID(entry_form,month):
    temp_data=pd.DataFrame()

    if entry_form.type == "weekly":
        for ID in entry_form.ID:
            for mon in month:
                count += entry_form.month.count("Y") # count the # of Y
            temp_data.append(entry_form.ID, entry_form.month)
            temp_data.Seva_attendance_days_per_month['ID'==ID]=count
            count = 0
        month_sessions = entry_form.month_session.count("Y") # count how many sessions held in month
        temp_data.total_sessions_per_month=month_sessions
        return temp_data
    
    elif entry_form.type == "monthly":
        # direct sum of each field seva_attendance and month_sessions
    
def write_to_master(DataFrame):
    # write the dataframe to google drive - using sheets API

'''


