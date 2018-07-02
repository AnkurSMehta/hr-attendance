# this compiles the data from individual sheets into a master attendance sheet

import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import xlrd
import xlwt
import csv
from datetime import datetime

scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('gspreadhw.json', scope)
client = gspread.authorize(creds)
'''
lookupfile = "20180516_sevarthi_master_list.xlsx"
sheet = client.open(lookupfile).sheet1
list_of_hashes=sheet.get_all_records(head=2)
lookup_col_labels=['Sevarthi Name','M/F','sub center','Round Year Dept','Event Dept', 'ID','Phone','Phone 2','E/Y',
                   'HRMUM ID','Current Status Feb 2018','Notes to verify','Additional Notes']
lookupdf=pd.DataFrame(list_of_hashes, index=None, columns=lookup_col_labels)
'''
lookup_col_labels=['Sevarthi Name','M/F','sub center','Round Year Dept','Event Dept', 'ID','Phone','Phone 2','E/Y',
                   'HRMUM ID','Current Status Feb 2018','Notes to verify','Additional Notes']
lookupdf=pd.read_excel("20180607_sevarthi_master_list.xlsx",header=0, skiprows=1)

def monthly_file_compile(filename, error_log, monthlist):
    sheet=client.open(filename).sheet1
    list_of_hashes=sheet.get_all_records()
    month_subset=[]
    for entry in list_of_hashes:
        for month in monthlist:
            if entry['Mon']==month:
                month_subset.append(entry)
    Col_labels=['Dep','Loc','Act','Name','ID','Mon','Days_Attendance','Total_Sessions']
    testdf=pd.DataFrame(month_subset, index=None, columns=Col_labels)
    
    for i, row in testdf.iterrows():
        if (row['Days_Attendance'] > row['Total_Sessions']):
            #err_msg = filename + " Too Many " + str(i)
            error_log.append((filename," Too Many Days", row['Name'], row['Mon']))
        
        if (row['Days_Attendance']==""):
            #err_msg = filename + " Blank Att " + str(i)
            error_log.append((filename," Blank Att ", row['Name'], row['Mon']))
        
        try:
            if lookupdf.ID[lookupdf.ID==int(row['ID'])].index.tolist()==[]:
                #err_msg = filename + " New ID " + str(i)
                #print row['ID'], type(row['ID'])
                #print lookupdf.ID.isin([row['ID']])[0]
                error_log.append((filename," New ID ", row['Name'], row['Mon']))
        except:
            if lookupdf.ID.isin([row['ID']])[0]==False:
                #err_msg = filename + " New ID " + str(i)
                #print row['ID'], type(row['ID'])
                #print lookupdf.ID.isin([row['ID']])[0]
                error_log.append((filename," New ID ", row['Name'], row['Mon']))

    return testdf

def weekly_file_compile(filename, error_log, monthlist):
    sheet=client.open(filename).sheet1
    list_of_hashes=sheet.get_all_values()
    df=pd.DataFrame(list_of_hashes)

    start_processing = False
    sevarthis = {}
    all_sevarthi_ids = []
    seva_days = {}
    masters = {}
    sevarthi_id=""
    dept=""
    loc=""
    act=""
    name=""

    for i,row in df.iterrows():
        row = row.to_dict()
        dept = row[1]
        loc = row[2]
        act = row[3]
        name = row[4]
        sevarthi_id = row[5]
        if sevarthi_id != "":
            masters[sevarthi_id] = "\t".join([dept, loc, act, name])

    for item, frame in df.iteritems():
        if str(frame[2]).strip() == "ID":
            for sevarthi_id in frame[3:]:
                sevarthis[sevarthi_id] = {}
                all_sevarthi_ids.append(sevarthi_id)
            start_processing = True
            continue
    
        if start_processing:
            #if(item.find("Unnamed") == -1):
            #    item = item.split(".")[0]
            formatted_ts = ""

            try:
                session_date = datetime.strptime(frame[2], "%Y-%m-%d")
                formatted_ts = session_date.strftime("%b")
            except:
                #print("Cannot parse date")
                continue
        
            if formatted_ts not in seva_days:
                seva_days[formatted_ts] = 0

                for sevarthi_id in all_sevarthi_ids:
                    sevarthis[sevarthi_id][formatted_ts] = 0

            if frame[1] == "y":
                seva_days[formatted_ts] += 1

            for i, attendance in enumerate(frame[3:]):
                if attendance.lower() == "y":
                    sevarthis[all_sevarthi_ids[i]][formatted_ts] += 1
    
    all_rows = []
    headers = ['Dep','Loc','Act','Name','ID','Mon','Days_Attendance','Total_Sessions']

    for current_id in sevarthis:
        for time_frame in sevarthis[current_id]:
            for month in monthlist:
                if time_frame==month:
                    try:
                        values = masters[current_id].split("\t") + [current_id, time_frame, sevarthis[current_id][time_frame], seva_days[time_frame]]
                        row = dict(zip(headers, values))
                        #print row
                        all_rows.append(row)
                    except:
                        #print("Some error")
                        #print row
                        continue 

    final_df = pd.DataFrame(all_rows, index=None, columns=headers)
    for i, row in final_df.iterrows():
        if (row['Days_Attendance'] > row['Total_Sessions']):
            #err_msg = filename + " Too Many " + str(i)
            error_log.append((filename," Too Many Days", row['Name'], row['Mon']))
        
        if (row['Days_Attendance']==""):
            #err_msg = filename + " Blank Att " + str(i)
            error_log.append((filename," Blank Att ", row['Name'], row['Mon']))
        
        try:
            if lookupdf.ID[lookupdf.ID==int(row['ID'])].index.tolist()==[]:
                #err_msg = filename + " New ID " + str(i)
                #print row['ID'], type(row['ID'])
                #print lookupdf.ID.isin([row['ID']])[0]
                error_log.append((filename," New ID ", row['Name'], row['Mon']))
        except:
            if lookupdf.ID.isin([row['ID']])[0]==False:
                #err_msg = filename + " New ID " + str(i)
                #print row['ID'], type(row['ID'])
                #print lookupdf.ID.isin([row['ID']])[0]
                error_log.append((filename," New ID ", row['Name'], row['Mon']))      
    
    #final_df.to_csv("Monthly_Conslidate.csv", sep=",", columns=headers)
    return final_df



Col_labels=['Dep','Loc','Act','Name','ID','Mon','Days_Attendance','Total_Sessions']
attend_master_df=pd.DataFrame(columns=Col_labels)
compiledf=pd.DataFrame(columns=Col_labels)
compile_weekly_df= pd.DataFrame(columns=Col_labels)

error_log=[]

for item in client.openall():
    file_to_compile = item.title
    if file_to_compile.endswith("MONTHLY.xlsx"):
        compiledf = monthly_file_compile(file_to_compile, error_log, ['Jul','Aug'])
        attend_master_df=attend_master_df.append(compiledf)
    if file_to_compile.endswith("WEEKLY.xlsx"):
        compile_weekly_df = weekly_file_compile(file_to_compile, error_log, ['Jul','Aug'])
        attend_master_df=attend_master_df.append(compile_weekly_df)    

writer = pd.ExcelWriter("Attendance_Master_Compile3.xlsx")
attend_master_df.to_excel(writer, columns=Col_labels)
writer.save()

with open("error-log.csv", "w") as file_to_write:
    for entry in error_log:
        #print entry
        rec = "%s\n" % ",".join(list(entry))
        file_to_write.write(rec)