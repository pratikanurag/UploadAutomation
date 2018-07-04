import pandas as pd
import glob,xlrd,xlwt
import datetime
now = datetime.datetime.now()
worksheet_data = []
worksheet_data_level1 = []
level1_id={}
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
risk_pivot = risk_all[['Level 1 Risk','Level 2 Risk']].drop_duplicates(subset=['Level 2 Risk']).dropna()
#risk_pivot = risk_pivot.drop_duplicates(subset=['Level 1 Risk','Level 2 Risk'])
risk_pivot.to_excel('output2.xlsx',index='true')
workbook = xlrd.open_workbook('output2.xlsx',on_demand = True)
workbook_level1 = xlrd.open_workbook('rsklevel1.xlsx',on_demand = True)
worksheet = workbook.sheet_by_index(0)
worksheet_level1 = workbook_level1.sheet_by_index(0)
for i in range(worksheet.nrows):
    worksheet_data.append(worksheet.row_values(i))
for k in range(worksheet_level1.nrows):
    worksheet_data_level1.append(worksheet_level1.row_values(k))
#print(worksheet_data) #worksheet data with groupby index
for j in range(0,len(worksheet_data)):
    del worksheet_data[j][0]
#print(worksheet_data) #worksheet data without group by index
#print(len(worksheet_data_level1))
level1_id = {worksheet_data_level1[x][0]:worksheet_data_level1[x][2] for x in range(1,len(worksheet_data_level1))}#dictionary with level 1 risks and ids
#print(level1_id)
#print(worksheet_data)
new_wb = xlwt.Workbook()
new_ws = new_wb.add_sheet('Risk')
new_ws.write(0,0,'Level 1 Risk')
new_ws.write(0,1,'Level 2 Risk')
new_ws.write(0,2,'Parent Object_id')
new_ws.write(0,3,'Level2 Object_id')
for y in range(1,len(worksheet_data)):
    new_ws.write(y,0,worksheet_data[y][0])
    new_ws.write(y,1,worksheet_data[y][1])
    new_ws.write(y,2,level1_id[worksheet_data[y][0]])
    new_ws.write(y,3,'RU2-'+str(now.strftime("%d%m%y"))+str(y))
new_wb.save('Risk_test.xls') 
    
    