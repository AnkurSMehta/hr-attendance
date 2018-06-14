# this compiles the data from individual sheets into a master attendance sheet

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




