# this is the validation code, to check accuracy and completeness of individual forms
import json

status = "unvalidated"

with open("config.json") as conf:
    CONF = json.load(conf)

sevarthi_dataframe = pd.read_excel(sev_list_path)

def Validate(entry_form, month):
    if !entry_form.check_missing_values(entry_form,month): # if there are missing values in ID or days_attendance or month
        status_update(entry_form, month)  # send email to administrator
    elif !entry_form.optional_missing_values():  # if there are missing values in name, sessions, activity
        HR_internal_update(entry_form, month)  # send email to HR administrator to fill missing values manually
    elif entry_form.unmatched_ID():
        status_update(entry_form, month)  # send email to administrator
    else:
        return True

def check_missing_values(entry_form,month):
    if entry_form.ID == "":
        status = "missing_ID"
        return False
        break
    elif entry_form.days_attendance == "":
        status = "missing_days"
        return False
        break
    elif entry_form.month == "":
        status = "missing_month"
        return False
        break
    else:
        return True

def status_update(entry_form,month):
    # email to CONF("HR_Admin_Email") send status, later status can go to dept coordinator

def unmatched_ID(entry_form,month):
    # vlookup ID in sevarthi_dataframe.ID send status to HR




