#!/usr/bin/env python
import requests
import json
import pandas as pd
import calendar
from datetime import datetime,timedelta
import numpy as np
import tkinter
from tkinter import ttk
from tkinter import filedialog
import win32api
import math
from decimal import Decimal

headers = {"x-api-key": "KEY", "x-api-secret":"SECRET"}

root= tkinter.Tk()
root.title("Certify")
canvas = tkinter.Canvas(root, width = 400, height = 300)
canvas.pack()

#label for input boxes
dateFromLabel = tkinter.Label(root,text = "Date From").place(x = 40,y = 50) 

dateToLabel = tkinter.Label(root,text = "Date To").place(x = 40,y = 90) 

expenseID = tkinter.Label(root,text = "Expense ID").place(x = 40,y = 130) 

#input boxes
dateFrom = tkinter.Entry(root) 
canvas.create_window(200, 60, window=dateFrom)

dateTo = tkinter.Entry(root) 
canvas.create_window(200, 100, window=dateTo)

expenseID = tkinter.Entry(root) 
canvas.create_window(200, 140, window=expenseID)

# progressbar
progress = ttk.Progressbar(
    root,
    orient='horizontal',
    mode = 'determinate',
    length=280
)


def bar():
    import time
    progress['value'] = 20
    root.update_idletasks()
    time.sleep(1)
  
    progress['value'] = 40
    root.update_idletasks()
    time.sleep(1)
  
    progress['value'] = 50
    root.update_idletasks()
    time.sleep(1)
  
    progress['value'] = 60
    root.update_idletasks()
    time.sleep(1)
  
    progress['value'] = 80
    root.update_idletasks()
    time.sleep(1)
    progress['value'] = 100
  


def validate_trial():
	dateFromInput = dateFrom.get()
	dateToInput = dateTo.get()      
	try:
		datetime.strptime(dateFromInput, '%Y-%m-%d') 
		datetime.strptime(dateToInput, '%Y-%m-%d')            
		get_expenses(dateFromInput, dateToInput)

	except ValueError:        
		win32api.MessageBox(0, 'Incorrect data format, should be YYYY-MM-DD', 'Incorrect date')

#submit button
submit_button = tkinter.Button(root,text = "Submit",width=30, command=validate_trial).place(x = 40,y = 180)

def run(startdate,enddate):
    threading.Thread(target=get_expense_data(startdate,enddate)).start()

def get_expense_data(startdate,enddate):
	root.title("Connecting to the API")
	# place the progressbar
	canvas.create_window(200, 255, window=progress)
	bar()
	print(startdate)
	print(enddate)
	df = pd.DataFrame()

	for i in range(1,100):  

		url = "https://api.certify.com/v1/expenses?startdate=" + startdate + "&enddate=" + enddate + "&page=" + str(i) 

		r = requests.get(url, headers=headers)
		print(r.status_code) #if error status code 403
		if(r.status_code == 403):
			print("error 403")
			win32api.MessageBox(0, "API pulled all the pages!",'Reached the last page!')
			break
		else:        
			content = json.loads(r.text)
			expenseReps = content['Expenses']      

			new_df = pd.json_normalize(expenseReps)
			df = df.append(new_df,ignore_index=True)

	return df

def get_expenses(startdate, enddate):
	departments = {
  		"Major Productions": 8,
  		"Clinical": 11,
	  	"Facilities": 103,
	  	"IT": 5,
	  	"Laboratory": 10,
	  	"Specialist Productions" : 9,
	  	"Business Development" : 6,
	  	"Human Resources" : 4,
	  	"Finance" : 3,
	  	"Management" : 2,
	  	"Quality" : 104,
	  	"Sales" : 101,
	  	"Operations": 7
	} 

	df = get_expense_data(startdate,enddate) 		

	currentMonth = datetime.now().month
	currentYear = datetime.now().year
	lastDay = calendar.monthrange(currentYear, currentMonth)[1]

	endDate = datetime.strptime(str(lastDay)+str(currentMonth)+str(currentYear),'%d%m%Y').strftime('%d-%b-%Y')

	DueDate = datetime.strptime(str(lastDay)+str(currentMonth)+str(currentYear),'%d%m%Y')
	DueDatePlus10= DueDate + timedelta(days=10)
	DueDateStr = DueDatePlus10.strftime('%d-%b-%Y')

	df["Date"] = endDate
	df["Date Due"] = DueDateStr
	df["ProductLine"]=""
	df["Product Line"]=df["ProductLine"]
	df['DateLine'] = pd.to_datetime(df['ExpenseDate'], format='%Y-%m-%d').dt.strftime('%d-%b-%Y')
	df['Date Line'] = df['DateLine']
	df['Department'] = ''
	df['Tax Code'] = df['ExpenseReportGLD2Code']
	df['Expense Report Number'] = ''
	df['Expenses Category'] = df['ExpenseCategory']
	df['Memo'] = df['Reason']
	df['Net Amount'] = df['ReimAmount']
	df = df.assign(rounding = lambda x: np.true_divide(np.ceil(x['Net Amount'] * 10**2), 10**2))
	df['Net Amount'] = df['rounding']
	df['Customer'] = df['ExpenseReportGLD1Code']	
	df.loc[df.Customer != '', 'Product Line'] = "2"
	df = df.replace({"DepartmentName": departments})


	df['Employee Name'] = df['FirstName'] + ' ' + df['LastName']
	df = df.sort_values("Employee Name")
	
	expenseIDInput = expenseID.get()
	expense_report_ids = []
	n = 0
	for count, employee in enumerate(df['Employee Name'].unique()):
		n +=1 
		new_report = {"External ID": expenseIDInput +'#'+"%03d" % n,"Employee Name": employee}
		expense_report_ids.append(new_report)
		

	report_id_df = pd.DataFrame.from_dict(expense_report_ids, orient='columns')

	df = df.merge(report_id_df, how="left", on="Employee Name")

	df = df[['External ID','Expense Report Number',"Employee Name","Date","Date Due","DepartmentName","Expenses Category", "Date Line","Memo","Net Amount","Tax Code","Product Line","Customer"]]
	report_id_df = df[['External ID','Expense Report Number',"Employee Name","Date"]]
	try:
		df.to_csv('my_data.csv', index=False)
		report_id_df.to_csv('header.csv', index=False)
		win32api.MessageBox(0, 'Data Has been Downloaded!','Data Downloaded')   
		root.title("Data Downloaded") 
	except PermissionError as e:
		win32api.MessageBox(0, "Make sure your excel file is shut!",'Error')   		
		print(e)
	#except KeyError(key) as k:
		#win32api.MessageBox(0, "No Reports found in that range!",'Reports not found')
		#print(k)


root.mainloop()