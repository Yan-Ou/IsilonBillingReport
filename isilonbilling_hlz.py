import paramiko
import pandas as pd
import numpy as np
import xlsxwriter
from datetime import datetime,date
import os.path

workbook = pd.ExcelFile("akl-isilon.xlsx")
worksheet = workbook.parse("HLZ")

paths = worksheet['Folder Location']
names = worksheet['Folder Name']
customers = worksheet['Customer ID Code ']

dt = date.today()

LINhex = []

ssh = paramiko.SSHClient()
ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
ssh.connect('', username='',password='')

for path in paths:
  stdin,stdout,stderr = ssh.exec_command("isi get -Dd "+path+"| grep LIN")
  result = stdout.readlines()
  LINhex.append(str((result[0].split())[2]).replace(':',''))

LINs = [str(int(element,16)) for element in LINhex]
dbpath = '/ifs/.ifsvar/modules/fsa/pub/latest/'
dbquery = 'for i in `ls -lrt '+dbpath+'|grep db|grep -v journal|grep disk_usage| awk -F" " \'{print $9}\'`; do sqlite3 -header -column '+dbpath+ '$i "select * from disk_usage where lin=' 

lograwsize = []
phyrawsize = []

for LIN, name in zip(LINs,names):
  stdin,stdout,stderr = ssh.exec_command(dbquery+LIN+'"; done;')
  print LIN, name
  read_result = (stdout.readlines()[2]).split()
  lograwsize.append(read_result[7])
  phyrawsize.append(read_result[8])

logsize = [float(element)/1024/1024/1024 for element in lograwsize]
physize = [float(element)/1024/1024/1024 for element in phyrawsize]

stdout,stdout,stderr = ssh.exec_command('ls -lrth /ifs/.ifsvar/modules/fsa/pub/latest | awk -F" " \'{print $6,$7,$8}\'')
rawdbcreation = str(stdout.readlines()).replace("[u'","").replace("\\n\']","")

dbcreation = datetime.strptime(rawdbcreation+' '+str(datetime.now().year), '%b %d %H:%M %Y')


worksheet['Logical Size (GB)'] = logsize
worksheet['Physical Size (GB)'] = physize
worksheet['Cost'] = worksheet['Logical Size (GB)']*worksheet['Charging Rate (per GB)']
worksheet['Time of Checking'] = dbcreation

reportName = 'AKL_iSilon_Billing_'+dbcreation.strftime("%B, %Y")+'.xlsx'

writer = pd.ExcelWriter(reportName,engine='xlsxwriter')
wb = writer.book
money_fmt = wb.add_format({'num_format': '$#,##0', 'bold': True})
total_fmt = wb.add_format({'align': 'right', 'num_format': '0.000','bold': True})
date_fmt = wb.add_format({'num_format': 'YYYY MMM DD'})

for name, customer in zip(names,customers):
  df = worksheet[(worksheet['Folder Name']==name)&(worksheet['Customer ID Code ']==customer)]
  if os.path.isfile(reportName):
    xls = pd.ExcelFile(reportName)
    sheetname = str(customer)+'_'+str(name)
    if sheetname in xls.sheet_names:
      df = xls.parse(sheetname)   
      df = df.drop(df[df['Customer ID Code ']== 'Average'].index)
      newrow = worksheet[(worksheet['Folder Name']==name)&(worksheet['Customer ID Code ']==customer)]
      df = df.append(newrow)
      df = df.drop_duplicates(['Time of Checking'], keep='first')
    
  df.loc[-1] = ['Average','','','','','','',df['Logical Size (GB)'].mean(), df['Physical Size (GB)'].mean(), df['Cost'].mean(),'']
  df.to_excel(writer, index=False,sheet_name=customer+'_'+name)
  ws = writer.sheets[customer+'_'+name]
  ws.set_column('H:H',20, total_fmt)
  ws.set_column('I:I',20, total_fmt)
  ws.set_column('J:J', 20, money_fmt)
  ws.set_column('K:K', 30, date_fmt)
  ws.set_column('C:D',30)
  ws.set_column('B:B',20)
  ws.set_column('E:E',10)
  ws.set_column('A:A',20)

writer.save()