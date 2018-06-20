import xlrd
import pandas as pd
import xlwt

path = 'E://Mumbai Center Dada//2014 04 HR//2017 mum att//20180518_sevarthi_master_list.xlsx'
path1 = 'E://Mumbai Center Dada//2014 04 HR//2017 mum att//20180315_attendance_master_with_pivots.xlsx'
book = xlrd.open_workbook(path1)
#print book.nsheets
#print book.sheet_names()
datasheet_index = 25
#datasheet = book.sheet_by_index(datasheet_index)
#testdatasheet=book.sheet_by_name('master data')
#test_val = testdatasheet.row_values(1)
#print datasheet.row_slice(rowx=5, start_colx=2, end_colx=10)
#print datasheet.row_values(1)
'''
book_w = xlwt.Workbook()
sheet1 = book_w.add_sheet("PySheet1")
for index,val in enumerate(test_val):
    sheet1.write(1,index,val)
#sheet1.write(1,0,test_val)
book_w.save("xlwt_test.xls")
'''
df=pd.read_excel(path1, sheetname=datasheet_index, header=0, skiprows=1)
#print df.head()
df_working = df.iloc[:,0:16]
df_working.columns = ['Dep','Loc','Act','Name','ID','Mon','Days','Sess','Rem1','Rem2','Rem3','Hrs','AdjHrs','TotHrs','MF','EY']
#print df_working.head()
Sev_names=df_working['ID'].unique()
#print Sev_names[0:10]
df_working.drop_duplicates(inplace=True, subset=['Dep','Loc','ID'])
df_working=pd.concat([df_working.iloc[:,0:5],df_working.iloc[:,8],df_working.iloc[:,14:16]], axis=1)
writer=pd.ExcelWriter("Sev_Dep_Unique_Combo.xlsx")
df_working.to_excel(writer,'Unique_list')
writer.save()
df_dep_names=df_working.iloc[:,0:2]
df_dep_names.drop_duplicates(inplace=True)
write2=pd.ExcelWriter("Dep_Unique_Combo.xlsx")
df_dep_names.to_excel(write2,'Unique_Dep')
write2.save()
