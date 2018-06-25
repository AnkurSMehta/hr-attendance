# this compiles the data from individual sheets into a master attendance sheet

import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import xlrd
import xlwt

scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('gspreadhw.json', scope)
client = gspread.authorize(creds)

lookupfile = "20180516_sevarthi_master_list.xlsx"
sheet = client.open(lookupfile).sheet1
list_of_hashes=sheet.get_all_records(head=2)
lookup_col_labels=['Sevarthi Name','M/F','sub center','Round Year Dept','Event Dept', 'ID','Phone','Phone 2','E/Y',
                   'HRMUM ID','Current Status Feb 2018','Notes to verify','Additional Notes']
lookupdf=pd.DataFrame(list_of_hashes, index=None, columns=lookup_col_labels)

def monthly_file_compile(filename, error_log):
    sheet=client.open(filename).sheet1
    list_of_hashes=sheet.get_all_records()
    Col_labels=['Dep','Loc','Act','Name','ID','Mon','Days_Attendance','Total_Sessions']
    testdf=pd.DataFrame(list_of_hashes, index=None, columns=Col_labels)
    
    for i, row in testdf.iterrows():
        if (row['Days_Attendance'] > row['Total_Sessions']):
            #err_msg = filename + " Too Many " + str(i)
            error_log.append((filename," Too Many Days", str(i)))
        elif (row['Days_Attendance']==""):
            #err_msg = filename + " Blank Att " + str(i)
            error_log.append((filename," Blank Att ", str(i)))
        elif lookupdf.ID.isin([row['ID']])[0]==False:
            #err_msg = filename + " New ID " + str(i)
            error_log.append((filename," New ID ", str(i)))
    return testdf

Col_labels=['Dep','Loc','Act','Name','ID','Mon','Days_Attendance','Total_Sessions']
attend_master_df=pd.DataFrame(columns=Col_labels)
compiledf=pd.DataFrame(columns=Col_labels)

error_log=[]

for item in client.openall():
    file_to_compile = item.title
    if file_to_compile.endswith("M.xlsx"):
        compiledf = monthly_file_compile(file_to_compile, error_log)
    attend_master_df=attend_master_df.append(compiledf)

writer = pd.ExcelWriter("Attendance_Master_Compile2.xlsx")
attend_master_df.to_excel(writer)
writer.save()

'''
getting error when compiledf = monthly_file  function is called
error_log.append((filename," New ID ", str(i)))
TypeError: 'str' object is not callable
'''
