# this python file generates an entry form for each center
import json
import pandas as pd

with open("config.json") as conf:
    CONF = json.load(conf)

sev_list_path = CONF("sevarthi_list_path")
center_list = CONF("Center_list")
activity_list=CONF("Activity_List")

sevarthi_dataframe = pd.read_excel(sev_list_path)

def generate(forms=ALL):

    list_of_all_forms = {}
    for indiv_center_form in forms:
        indiv_list = get_sevarthi_list(indiv_center_form)
        indiv_filled_form=populate_form(indiv_list, indiv_center_form)
        upload_google_drive(indiv_filled_form)

def get_sevarthi_list(form):
    return sevarthi_dataframe['Dep'=form.dep & 'Loc'=form.loc]

def populate_form(name_list, center):
    if center.type = "monthly":
        temp_form=pd.DataFrame()
        temp_form.append(name_list)
        temp_form.add_column("Activity")
        temp_form.Activity="execution" # this is hardcoded as a placeholder - to be changed
        temp_form.add_column("month", "Seva_attendance_days_per_month","total_sessions_per_month","remarks")
        temp_form.total_sessions_per_month=4  #this is hardcoded - to be parameter from config.json
    elif center.type = "weekly":
        # similar, but need to add weekly sunday calendar
    elif center.type = "GNC":
        # need to add weekly and additional info about mahatma attendance, maybe all in single excel workbook
    elif center.type = "event":
        # similar as monthly
    
    return temp_form

def upload_google_drive(form):
    # temp_form to be uploaded - using config.json again?
     


    


