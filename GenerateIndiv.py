# this python file generates an entry form for each center
import json
import pandas as pd
import xlrd
import xlwt
from datetime import date, timedelta

with open("config.json") as conf:
    CONF = json.load(conf)

round_year_sev_list = CONF["round_year_sevarthi_list"]

round_year_center_list = CONF["round_year_center_list"]
dest_generate = CONF["output_indiv_path"]

center_list = pd.read_excel(round_year_center_list)
sevarthi_dataframe = pd.read_excel(round_year_sev_list)

# Direct implementation - read in the 2 excel configuration files
#center_list = pd.read_excel("E://Mumbai Center Dada//2014 04 HR//attendance_software//Round_Year_Center_List.xlsx")
#sevarthi_dataframe = pd.read_excel("E://Mumbai Center Dada//2014 04 HR//attendance_software//Sevarthi_List_Round_Year.xlsx")

# filter into 2 dataframes - monthly and weekly
center_list_monthly=center_list[center_list.Type=="M"]
center_list_weekly=center_list[center_list.Type=="W"]
center_list_weekly=center_list_weekly[center_list.Dep.str.contains("LMHT") | center_list.Dep.str.contains("BMHT")]

'''
# generate a list of dataframes, each element has the sevarthi list
len=center_list_monthly.shape[0]
list_of_all_forms=[]
for i in range(len):
    list_of_all_forms.append(pd.merge(center_list_monthly[i:i+1], sevarthi_dataframe.iloc[:,0:5], on=['Dep','Loc']))

# generate excel files with multiple blocks - 1 for each month
Month_list=['Jul','Aug','Sep','Oct','Nov','Dec']
for i in range(len):
    temp_df=list_of_all_forms[0].head(0)
    #temp_df['Mon']=""
    for month in Month_list:
        list_of_all_forms[i]['Mon']=month
        temp_df=temp_df.append(list_of_all_forms[i])  
    temp_df['Days_Attendance']=""
    temp_df['Total_Sessions']=""
    del temp_df['Type']
    
    writer=pd.ExcelWriter(dest_generate+list_of_all_forms[i].Dep[0]+"_"+list_of_all_forms[i].Loc[0]+"MONTHLY"+".xlsx")
    temp_df.to_excel(writer,"Attendance_Form")
    worksheet=writer.sheets['Attendance_Form']
    workbook=writer.book
    grayfont = workbook.add_format({'font_color':'gray'})
    worksheet.set_column("A:A",2, grayfont)
    worksheet.set_column("B:D",18)
    worksheet.set_column("E:E",30)
    worksheet.set_column("F:G",8)
    worksheet.set_column("H:I",18)
    worksheet.conditional_format('B3:D200', {'type':'cell','criteria': '=','value':"B2",'format': grayfont})
    #worksheet.set_column('A:A', None, grayfont)
    writer.save()
'''

# to generate excel files for weekly sevarthi lists
first_date = date(2018,6,1)+timedelta(6-date(2018,6,1).weekday())
date_list=[]
while first_date.year==2018:
    date_list.append(str(first_date))
    first_date += timedelta(days=7)
date_listdf=pd.DataFrame(date_list)
date_listdf.columns=['% Attendance']
date_listdf=date_listdf.T
title="Session_Held?(Y/N)-->>"

list_week_forms=[]
lenw=center_list_weekly.shape[0]
for i in range(lenw):
    list_week_forms.append(pd.merge(center_list_weekly[i:i+1], sevarthi_dataframe.iloc[:,0:5], on=['Dep','Loc']))
for i in range(lenw):
    del list_week_forms[i]['Type']
    writer = pd.ExcelWriter(dest_generate+list_week_forms[i].Dep[0]+"_"+list_week_forms[i].Loc[0]+"WEEKLY"+".xlsx")
    list_week_forms[i].to_excel(writer,"Attendance_Form", startrow=2)
    date_listdf.to_excel(writer, "Attendance_Form", startrow=2, startcol=6, header=False)
    sevarthi_count=list_week_forms[i].shape[0]
    worksheet=writer.sheets['Attendance_Form']
    workbook=writer.book
    grayfont = workbook.add_format({'font_color':'gray'})
    percent_fmt=workbook.add_format({'num_format':'0%'})
    worksheet.write(1,6,title)
    j=6
    worksheet.write(sevarthi_count+j,1,"No of kids present")
    worksheet.write(sevarthi_count+j+1,1,"No of sevarthis present")
    worksheet.write(sevarthi_count+j+2,1,"Prasadi item")
    worksheet.write(sevarthi_count+j+3,1,"Prasadi cost")
    worksheet.write(sevarthi_count+j+4,1,"Other Misc cost")
    worksheet.write(sevarthi_count+j+5,1,"Avg cost per kid")
    worksheet.write(sevarthi_count+j+7,1,"Prasadi Team")
    for n in range(20):
        worksheet.write(sevarthi_count+j+8+n,0,n+1)
    for rownum in range(sevarthi_count):
        worksheet.write_formula(3+rownum,6,'=iferror(countif($H%d:$AL%d,"y")/countif($H$2:$AL$2,"y"), "")' % (4+rownum, 4+rownum))
    
    worksheet.write(1,7,"n")
    worksheet.write(1,8,"y")
    worksheet.write(3,7,"n")
    worksheet.write(3,8,"y")

    worksheet.set_column("A:A",2,grayfont)
    worksheet.set_column("B:B",22)
    worksheet.set_column("C:D",18)
    worksheet.set_column("E:E",30)
    worksheet.set_column("F:F",8)
    worksheet.set_column("G:G",22, percent_fmt)
    worksheet.set_column("H:BZ",10)
    worksheet.conditional_format('B4:D200', {'type':'cell','criteria': '=','value':"B3",'format': grayfont})
    #worksheet.set_column('A:A', None, grayfont)
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
     


    


