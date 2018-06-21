# this python file generates an entry form for each center
import json
import pandas as pd
import xlrd
import xlwt

with open("config.json") as conf:
    CONF = json.load(conf)

#round_year_sev_list = CONF("round_year_sevarthi_list")

#round_year_center_list = CONF("Round_Year_Center_list")
#Output_path = CONF("Attendance_Directory")

#center_list = pd.read_excel(round_year_center_list) - getting error 'dict' object not callable
#sevarthi_dataframe = pd.read_excel(round_year_sev_list)

# Direct implementation - read in the 2 excel configuration files
center_list = pd.read_excel("E://Mumbai Center Dada//2014 04 HR//attendance_software//Round_Year_Center_List.xlsx")
sevarthi_dataframe = pd.read_excel("E://Mumbai Center Dada//2014 04 HR//attendance_software//Sevarthi_List_Round_Year.xlsx")

# filter into 2 dataframes - monthly and weekly
center_list_monthly=center_list[center_list.Type=="M"]
center_list_weekly=center_list[center_list.Type=="W"]

# generate a list of dataframes, each element has the sevarthi list
len=center_list_monthly.shape[0]
test=[]
for i in range(len):
    test.append(pd.merge(center_list_monthly[i:i+1], sevarthi_dataframe.iloc[:,0:5], on=['Dep','Loc']))

# generate excel files with multiple blocks - 1 for each month
Month_list=['Jul','Aug','Sep','Oct','Nov','Dec']
for i in range(len):
    temp_df=test[0].head(0)
    #temp_df['Mon']=""
    for month in Month_list:
        test[i]['Mon']=month
        temp_df=temp_df.append(test[i])  
    temp_df['Days_Attendance']=""
    temp_df['Total_Sessions']=""
    del temp_df['Type']
    
    writer=pd.ExcelWriter(test[i].Dep[0]+"_"+test[i].Loc[0]+"M"+".xlsx")
    temp_df.to_excel(writer,"Attendance_Form")
    writer.save()


'''
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
'''
     


    


