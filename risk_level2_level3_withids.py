import pandas as pd
import glob,xlrd,xlwt
import datetime
now = datetime.datetime.now()
worksheet_data1_2 = []
worksheet_data2_3 = []
worksheet_data_level1 = []
worksheet_data_level2 = []
level1_id={}
level2_id = {}
risk_excel = 'Risk_Universe_Master1.xlsx'
risk_e1 = pd.read_excel(risk_excel,sheetname = 1,index_col = 0)
risk_e2 = pd.read_excel(risk_excel,sheetname = 2,index_col = 0)
risk_e3 = pd.read_excel(risk_excel,sheetname = 3,index_col = 0)
risk_e4 = pd.read_excel(risk_excel,sheetname = 4,index_col = 0)
risk_e5 = pd.read_excel(risk_excel,sheetname = 5,index_col = 0)
risk_e6 = pd.read_excel(risk_excel,sheetname = 6,index_col = 0)
risk_e7 = pd.read_excel(risk_excel,sheetname = 7,index_col = 0)
risk_e8 = pd.read_excel(risk_excel,sheetname = 8,index_col = 0)
risk_e9 = pd.read_excel(risk_excel,sheetname = 9,index_col = 0)
risk_all = pd.concat([risk_e1,risk_e2,risk_e3,risk_e4,risk_e5,risk_e6,risk_e7,risk_e8,risk_e9])
risk_pivot2_3 = risk_all[['Level 2 Risk','Level 3 Risk']].drop_duplicates(subset=['Level 3 Risk']).dropna()
risk_pivot2_3.to_excel('output3.xlsx',index='true')
workbook2_3 = xlrd.open_workbook('output3.xlsx',on_demand = True)
workbook_level2 = xlrd.open_workbook('Risk_test.xls',on_demand = True)
worksheet2_3 = workbook2_3.sheet_by_index(0)
worksheet_level2 = workbook_level2.sheet_by_index(0)
for x in range(worksheet2_3.nrows):
    worksheet_data2_3.append(worksheet2_3.row_values(x))
for k in range(worksheet_level2.nrows):
    worksheet_data_level2.append(worksheet_level2.row_values(k))
for z in range(0,len(worksheet_data2_3)):
    del worksheet_data2_3[z][0]
#print(worksheet_data2_3)#level 2 - level 3 mapping
level2_id = {worksheet_data_level2[x][1]:worksheet_data_level2[x][3] for x in range(1,len(worksheet_data_level2))}#dictionary with level 2 risks and ids
new_wb = xlwt.Workbook()
new_ws = new_wb.add_sheet('Risk_1')
new_ws.write(0,0,'Level 2 Risk')
new_ws.write(0,1,'Level 3 Risk')
new_ws.write(0,2,'Parent Object_id')
new_ws.write(0,3,'Level3 Object_id')
for y in range(1,len(worksheet_data2_3)):
    new_ws.write(y,0,worksheet_data2_3[y][0])
    new_ws.write(y,1,worksheet_data2_3[y][1])
    new_ws.write(y,2,level2_id[worksheet_data2_3[y][0]])
    new_ws.write(y,3,'RU3-'+str(now.strftime("%d%m%y"))+str(y))
new_wb.save('Risk_test_level3.xls') 

