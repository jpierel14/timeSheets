from __future__ import print_function
from tkinter import *
from tkinter import messagebox
import pandas as pd
import numpy as np
import time
from collections import OrderedDict as odict 
import httplib2
import os
import json
from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
from oauth2client.service_account import ServiceAccountCredentials
from apiclient import errors
import datetime
import pickle

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

import gspread
from os import path

__users__={
	'cathy':'pieces',
	'clarissa':'assembly1',
	'ellie':'assembly1',
	'hannah':'assembly1',
	'nikki':'assembly1'
}

file=open('currentPayPeriod.txt','r')
dates=file.read()
file.close()
global __filename__
__filename__='PW_TimeSheet_'+dates
print(__filename__)


__hours__=['0','1','2','3','4','5','6','7','8','9','10','11','12']
__minutes__=['0','5','10','15','20','25','30','35','40','45','50','55']
__companies__=['IT','Loo Hoo','Serendipity','Solomons','Myler','RPM']
#__items__=['Cups','Medical Rubber','Lids','Loo hoo']
json_key = json.load(open('gspread-test.json'))['installed']
scope = [
'https://spreadsheets.google.com/feeds',
'https://www.googleapis.com/auth/drive'
]
credentials = ServiceAccountCredentials.from_json_keyfile_name('gspread-test3.json', scope)
gc=gspread.authorize(credentials)

def userCheck():
	global __users__
	if path.isfile('timesheet.pkl'):    
		inFile = open('timesheet.pkl', 'rb')
		__users__=pickle.load(inFile)
		inFile.close()
		
def center(toplevel):
		toplevel.update_idletasks()
		w = toplevel.winfo_screenwidth()
		h = toplevel.winfo_screenheight()
		size = tuple(int(_) for _ in toplevel.geometry().split('+')[0].split('x'))
		x = w/2 - size[0]/2
		y = h/2 - size[1]/2
		toplevel.geometry("%dx%d+%d+%d" % (size + (x, y)))
		
def readItems():
	#gc=gspread.authorize(credentials)
	global __filename__
	user="Clarissa"
	print(__filename__)
	wks=gc.open(__filename__).worksheet(user)
	tempCell=wks.find('Pay period')
	period=str(wks.cell(tempCell.row+2,tempCell.col).value)
	labor_list=dict([])
	values_list = wks.col_values(2)
	index=values_list.index('Product Information')
	values_list=values_list[index+1:]
	temp=wks.col_values(4)
	temp=temp[index+1:]
	for i in range(len(values_list)):
		if values_list[i] not in __companies__ and values_list[i] !='':
			labor_list[values_list[i]]=temp[i]
	values_list=[value.strip() for value in values_list if value.strip() not in __companies__ and  value != '']
	return period,labor_list,sorted(values_list+['Meetings','Transition time','Recording','Misc'],key=lambda s: s.lower())

	
def importDrive(sheet):
	global __filename__
	#user='Clarissa'
	user=app.userVariable.get()
	user=user.title()
	#wks=gc.open("PW_TimeSheet")
	#try:
	wks=gc.open(__filename__).worksheet(user)
	#except:
	#	wks=gc.create("PW_TimeSheet_")
	for item in sheet.itemsSelected.keys():
		if 'min' in sheet.itemsSelected[item].keys() and 'hour' in sheet.itemsSelected[item].keys() and 'num' in sheet.itemsSelected[item].keys(): 
			cell=wks.find(item)
			if item in ['Meetings','Transition time','Recording','Misc']:
				wks.update_cell(cell.row,cell.col+2,str(int(sheet.itemsSelected[item]['hour'])*60)+sheet.itemsSelected[item]['min'])
			else:
				wks.update_cell(cell.row,cell.col+3,sheet.itemsSelected[item]['num'])
				wks.update_cell(cell.row,cell.col+5,str(int(sheet.itemsSelected[item]['hour'])*60+int(sheet.itemsSelected[item]['min'])))
	
	tempCell=wks.find('Gross Pay')
	return str(wks.cell(tempCell.row+1,tempCell.col).value)
	#print(dateRange.input_value[0:3])
	#wks.update_acell("B2","Hi Ellie")
	#sh = gc.create('testSpreadsheet')
	#sh.share('pieceworksinc@gmail.com', perm_type='user', role='owner')
	
def getTotal(user):
	global __filename__
	user=user.title()
	wks=gc.open(__filename__).worksheet(user)
	tempCell=wks.find('Gross Pay')
	return str(wks.cell(tempCell.row+1,tempCell.col).value)
	
class login(Tk):
	global __users__
	def __init__(self,parent):
		Tk.__init__(self,parent)
		self.parent = parent
		self.minsize(width=400,height=100)
		center(self)
		self.initialize()
	
	def initialize(self):
		self.grid()
		self.label1Variable=StringVar()
		self.label1Variable.set("Please enter your Username and Password")
		label1=Label(self,textvariable=self.label1Variable,anchor="w",fg='black')
		label1.grid(column=1,row=0,columnspan=2,sticky='EW')
		
		
		self.userVariable=StringVar()
		self.user=Entry(self,textvariable=self.userVariable)
		self.user.grid(column=1,row=2,columnspan=1,sticky='Ew')
		self.user.bind("<Return>",self.OnPressEnter)
		self.userVariable.set(u"Username")
		self.user.focus_set()
		self.user.selection_range(0,END)
		self.passVariable=StringVar()
		self.password=Entry(self,textvariable=self.passVariable)
		self.password.grid(column=1,row=3,columnspan=1,sticky='Ew')
		self.password.bind("<Return>",self.OnPressEnter)
		self.passVariable.set(u"Password")
		
		button = Button(self,text=u"Continue",command=self.OnButtonClick)
		button.grid(column=3,row=4)
		
		passChanger=Button(self,text=u"Change password",command=self.passChangeFunc)
		passChanger.grid(column=1,row=4)
		
		self.grid_columnconfigure(1,weight=1)
		self.resizable(True,True)
		self.update()

	def passChangeFunc(self):
		self.destroy()
		change=changePassword(None)
		change.title('Change Password')
		change.mainloop()
	
	def OnPressEnter(self,event):
		if (self.userVariable.get()).lower() in __users__.keys():   
			if self.passVariable.get() == __users__[(self.userVariable.get()).lower()]:
				self.destroy()
				if str(app.userVariable.get()).lower()=='cathy' and str(app.passVariable.get())==__users__['cathy']:
					owner=own(None)
					owner.title('Time Sheet Keeper-President & CEO')
					owner.mainloop()
				else:
					global __total__
					__total__=getTotal(str(app.userVariable.get()))
					sheet = timeSheet(None,user=str(app.userVariable.get()),password=str(app.passVariable.get()))
					sheet.title('Time Sheet Keeper:'+__period__+' - '+self.userVariable.get().title())
					sheet.mainloop()
				#export(sheet)
			else:
				self.labelVariable = StringVar()
				label=Label(self,textvariable=self.labelVariable,anchor="w",fg="white",bg="Red")
				label.grid(column=0,row=4,columnspan=2,sticky='EW')
				self.labelVariable.set("Invalid Password.")
				self.password.focus_set()
				self.password.selection_range(0,END)
		else:
			self.labelVariable = StringVar()
			label=Label(self,textvariable=self.labelVariable,anchor="w",fg="white",bg="Red")
			label.grid(column=0,row=4,columnspan=2,sticky='EW')
			self.labelVariable.set("Invalid Username.")
			self.user.focus_set()
			self.user.selection_range(0,END)
			
	def OnButtonClick(self):
		if (self.userVariable.get()).lower() in __users__.keys():   
			if self.passVariable.get() == __users__[(self.userVariable.get()).lower()]:
				self.destroy()
				if str(app.userVariable.get()).lower()=='cathy' and str(app.passVariable.get())==__users__['cathy']:
					owner=own(None)
					owner.title('Time Sheet Keeper-President & CEO')
					owner.mainloop()
				else:
					global __total__
					__total__=getTotal(str(app.userVariable.get()))
					sheet = timeSheet(None,user=str(app.userVariable.get()),password=str(app.passVariable.get()))
					sheet.title('Time Sheet Keeper:'+__period__+' - '+self.userVariable.get().title())
					sheet.mainloop()
				#export(sheet)
			else:
				self.labelVariable = StringVar()
				label=Label(self,textvariable=self.labelVariable,anchor="w",fg="white",bg="Red")
				label.grid(column=0,row=4,columnspan=2,sticky='EW')
				self.labelVariable.set("Invalid Password.")
				self.password.focus_set()
				self.password.selection_range(0,END)
		else:
			self.labelVariable = StringVar()
			label=Label(self,textvariable=self.labelVariable,anchor="w",fg="white",bg="Red")
			label.grid(column=0,row=4,columnspan=2,sticky='EW')
			self.labelVariable.set("Invalid Username.")
			self.user.focus_set()
			self.user.selection_range(0,END)

class changePassword(Tk):
	global __users__
	def __init__(self,parent):
		Tk.__init__(self,parent)
		self.parent=parent
		self.minsize(width=400,height=100)
		center(self)
		self.initialize()

	def initialize(self):
		self.grid()
		self.userVariable=StringVar()
		self.user=Entry(self,textvariable=self.userVariable)
		self.user.grid(column=0,row=0,columnspan=1,sticky='Ew')
		self.user.bind("<Return>",self.OnPressEnter)
		self.userVariable.set(u"Old Username")
		self.passVar=StringVar()
		self.password=Entry(self,textvariable=self.passVar)
		self.password.grid(column=0,row=1,columnspan=1,sticky='Ew')
		self.password.bind("<Return>",self.OnPressEnter)
		self.passVar.set(u"Old Password")
		self.passChangeVar=StringVar()
		self.passChange=Entry(self,textvariable=self.passChangeVar)
		self.passChange.grid(column=0,row=2,columnspan=1,sticky='Ew')
		self.passChange.bind("<Return>",self.OnPressEnter)
		self.passChangeVar.set(u"New Password")
		self.passConfirmVar=StringVar()
		self.passConfirm=Entry(self,textvariable=self.passConfirmVar)
		self.passConfirm.grid(column=0,row=3,columnspan=1,sticky='Ew')
		self.passConfirm.bind("<Return>",self.OnPressEnter)
		self.passConfirmVar.set(u"Confirm Password")
		button = Button(self,text=u"Continue",command=self.OnButtonClick)
		button.grid(column=3,row=4)
		self.grid_columnconfigure(1,weight=1)
		self.resizable(True,True)
		self.user.focus_set()
		self.update()
		
	def OnPressEnter(self,event):
		if (self.userVariable.get()).lower() in __users__.keys():   
			if self.passVar.get() == __users__[(self.userVariable.get()).lower()]:
				if self.passChangeVar.get()=="New Password" or self.passChangeVar.get()=='':
					messagebox.showwarning("Error","Please enter a new password")
				elif self.passConfirmVar.get() != self.passChangeVar.get():
					messagebox.showwarning("Error","Incorrect password confirmation")
				else:
					__users__[self.userVariable.get().lower()]=self.passChangeVar.get()
					messagebox.showwarning("","Success!")
					self.destroy()
					app = login(None)
					app.title('Time Sheet Keeper:'+__period__)
					app.mainloop()
			else:
				messagebox.showwarning("Error","Invalid Password")
		else:
			messagebox.showwarning("Error","Invalid Username")
			
	def OnButtonClick(self):
		if (self.userVariable.get()).lower() in __users__.keys():   
			if self.passVar.get() == __users__[(self.userVariable.get()).lower()]:
				if self.passChangeVar.get()=="New Password" or self.passChangeVar.get()=='':
					messagebox.showwarning("Error","Please enter a new password")
				elif self.passConfirmVar.get() != self.passChangeVar.get():
					messagebox.showwarning("Error","Incorrect password confirmation")
				else:
					__users__[self.userVariable.get().lower()]=self.passChangeVar.get()
					messagebox.showwarning("","Success!")
					self.destroy()
					app = login(None)
					app.title('Time Sheet Keeper:'+__period__)
					app.mainloop()
			else:
				messagebox.showwarning("Error","Invalid Password")
		else:
			messagebox.showwarning("Error","Invalid Username")
			
class own(Tk):
	global __users__
	global __filename__
	def __init__(self,parent):
		Tk.__init__(self,parent)
		self.parent = parent
		self.minsize(width=500,height=500)
		center(self)
		self.initialize()

	def initialize(self):
		self.finish=Button(self,text=u"Finish",command=self.OnButtonClick)
		self.finish.place(relx=.95,rely=.95,anchor='center')
		cancel=Button(self,text=u"Cancel",command=self.cancel)
		cancel.place(relx=.85,rely=.95,anchor='center')
		myImage=PhotoImage(file='pieceworks.gif')
		myImage=myImage.subsample(3,3)
		self.im=Label(self,image=myImage)
		self.im.image=myImage
		#self.im.place(relx=.5,rely=.4,anchor=CENTER)
		self.im.place(relx=.01,rely=.92)
		self.grid()
		new = Button(self,text=u"New Pay Period",command=self.reset)
		new.grid(column=0,row=1)
		addUser=Button(self,text=u"Add worker",command=self.newUser)
		addUser.grid(column=0,row=2)
		
	def reset(self):
		self.startLabel=StringVar()
		start=Label(self,textvariable=self.startLabel,anchor='center')
		start.grid(column=1,row=0,columnspan=1,sticky='EW')
		self.endLabel=StringVar()
		end=Label(self,textvariable=self.endLabel,anchor='center')
		end.grid(column=2,row=0,columnspan=1,sticky='EW')
		self.startVariable = StringVar()
		self.startVariable.set('')
		self.entry=Entry(self,textvariable=self.startVariable)
		self.entry.grid(column=1,row=1,sticky='Ew')
		self.entry.bind("<Return>",self.OnPressEnter)
		self.endVariable=StringVar()
		self.endVariable.set('')
		self.entry2=Entry(self,textvariable=self.endVariable)
		self.entry2.grid(column=2,row=1,sticky='EW')
		self.entry2.bind("<Return>",self.OnPressEnter)
	
	def OnButtonClick(self):
		if self.endVariable != '' and self.startVariable!='':
			try:
				start=datetime.datetime.strptime(self.startVariable.get(),'%m/%d')
				end=datetime.datetime.strptime(self.endVariable.get(),'%m/%d')
			except RuntimeError:
				messagebox.showwarning("Error","Dates need to be in MM/DD Format")
			f=open('currentPayPeriod.txt','w')
			f.write(self.startVariable.get()+'-'+self.endVariable.get())
			self.resetSheets()
					
			self.destroy()


	def OnPressEnter(self,event):
		if self.endVariable != '' and self.startVariable!='':
			try:
				start=datetime.datetime.strptime(self.startVariable.get(),'%m/%d')
				end=datetime.datetime.strptime(self.endVariable.get(),'%m/%d')
			except RuntimeError:
				messagebox.showwarning("Error","Dates need to be in MM/DD Format")
			open('currentPayPeriod.txt','w').close()
			f=open('currentPayPeriod.txt','w')
			__filename__=self.startVariable.get()+'-'+self.endVariable.get()
			f.write(__filename__)
			f.close()
			self.resetSheets()
					
	def newUser(self):
		pass
	'''
	def resetSheets(self):
		self.progressBar = StringVar()
		progress=Label(self,textvariable=self.progressBar,anchor="center",fg="black")
		progress.grid(column=0,row=100,columnspan=1,sticky='EW')
		self.progressBar.set(u"Updating Information...0%")
		self.update()
		userNum=1
		for user in __users__:
			if user != 'cathy':
				wks=gc.open("PW_TimeSheet").worksheet(user.title())
				for item in __items__:
					try:
						tempcell=wks.find(item)
						if item in ['Meetings','Transition time','Recording','Misc']:
							wks.update_cell(tempcell.row,tempcell.col+2,'0')
						else:
							wks.update_cell(tempcell.row,tempcell.col+3,'0')
							wks.update_cell(tempcell.row,tempcell.col+5,'0')
					except:
						continue
				self.progressBar.set(u"Updating Information..."+str(userNum/(len(__users__)-1)*100)+"%")
				self.update()
	'''
	def resetSheets(self):
		sh=gc.create(__filename__)
		sh.share('pieceworksinc@gmail.com', perm_type='user', role='owner')
		for user in [x for x in __users__.keys()+'blank' if x !='cathy']:
			wks=sh.add_worksheet(title=user)
	def cancel(self):
		self.destroy()
	
class timeSheet(Tk):
	global __users__
	def __init__(self,parent,user,password):
		Tk.__init__(self,parent,user,password)
		self.parent = parent
		self.minsize(width=500,height=500)
		center(self)
		self.initialize()

	def initialize(self):
		self.finish=Button(self,text=u"Finish",command=self.OnButtonClick,state='disabled')
		self.finish.place(relx=.95,rely=.95,anchor='center')
		self.cancel=Button(self,text=u"Cancel",command=self.cancel)
		self.cancel.place(relx=.85,rely=.95,anchor='center')
		self.totalVariable=StringVar()
		totalLabel=Label(self,textvariable=self.totalVariable,anchor='center')
		totalLabel.place(relx=.1,rely=.95)
		self.totalVariable.set("Your current pay period total is: "+__total__)
		totalLabel.config(font=("Courier",10))
		myImage=PhotoImage(file='pieceworks.gif')
		myImage=myImage.subsample(3,3)
		self.im=Label(self,image=myImage)
		self.im.image=myImage
		#self.im.place(relx=.5,rely=.4,anchor=CENTER)
		self.im.place(relx=.01,rely=.92)
		self.grid()
		self.row=2
		self.variable=False
		self.itemsSelected=odict([])

		self.labelVariable=StringVar()
		label=Label(self,textvariable=self.labelVariable,anchor='w')
		label.grid(column=0,row=0,columnspan=1,sticky='EW')
		self.labelVariable.set(" ")
		
		self.labelVariable=StringVar()
		label=Label(self,textvariable=self.labelVariable,anchor='w')
		label.grid(column=0,row=1,columnspan=1,sticky='EW')
		self.labelVariable.set("Welcome!")
		
		
		
		button = Button(self,text=u"Add Item",command=self.addItem)
		button.grid(column=0,row=2)
		
	def addHours(self):
		self.labelVariable = StringVar()
		label=Label(self,textvariable=self.labelVariable,anchor="center",fg="black")
		label.grid(column=3,row=1,columnspan=1,sticky='EW')
		self.labelVariable.set(u"Hours")
		self.hours=StringVar()
		self.hours.set("Hours")
		self.hours.trace('w',self.hourChange)
		hour=OptionMenu(self,self.hours,*__hours__)
		hour['width']=5
		hour.grid(column=3,row=self.row)
		
	def addMinutes(self):
		self.labelVariable = StringVar()
		label=Label(self,textvariable=self.labelVariable,anchor="center",fg="black")
		label.grid(column=4,row=1,columnspan=1,sticky='EW')
		self.labelVariable.set(u"Minutes")
		self.mins=StringVar()
		self.mins.set("Minutes")
		self.mins.trace('w',self.minChange)
		min=OptionMenu(self,self.mins,*__minutes__)
		min['width']=5
		min.grid(column=4,row=self.row)
		
	def addItemNum(self):
		self.labelVariable = StringVar()
		self.label=Label(self,textvariable=self.labelVariable,anchor="center",fg="black")
		self.label.grid(column=2,row=1,columnspan=1,sticky='EW')
		self.labelVariable.set(u"Number of Pieces")
		self.entryVariable = StringVar()
		self.entryVariable.set('')
		self.entryVariable.trace('w',self.itemNumChange)
		self.entry=Entry(self,textvariable=self.entryVariable)
		self.entry.grid(column=2,row=self.row,sticky='Ew')
		self.entry.bind("<Return>",self.OnPressEnter)
	
	def itemNumChange(self,*args):
		if self.entryVariable.get() !='' and self.entryVariable.get() not in self.itemsSelected[self.variable.get()].keys():
			self.itemsSelected[self.variable.get()]['num']=self.entryVariable.get()
		
	def itemChange(self,*args):
		if self.variable.get() !="Choose Item" and self.variable.get() not in self.itemsSelected:
			self.itemsSelected[self.variable.get()]=odict([])	
			if self.variable.get() in ['Meetings','Transition time','Recording','Misc'] or __labor__[self. variable.get()]=='hr':
				self.label.destroy()
				self.entry.destroy()
				self.itemsSelected[self.variable.get()]['num']=''
			
	
	def hourChange(self,*args):
		if self.hours.get() !="Minutes" and self.hours.get() not in self.itemsSelected[self.variable.get()]:
			self.itemsSelected[self.variable.get()]['hour']=self.hours.get()
	
	def minChange(self,*args):
		if self.mins.get() !="Minutes" and self.mins.get() not in self.itemsSelected[self.variable.get()]:
			self.itemsSelected[self.variable.get()]['min']=self.mins.get()
		if self.finish['state']=='disabled' and self.itemsSelected and 'min' in self.itemsSelected[self.variable.get()].keys() and 'hour' in self.itemsSelected[self.variable.get()].keys() and 'num' in self.itemsSelected[self.variable.get()].keys():
			self.finish['state']='normal'
		
	def addItem(self):
		if not [item for item in __items__ if item not in self.itemsSelected] or self.row-1>len(__items__):
			messagebox.showwarning("Error","No more item options.")
		elif self.variable and (self.variable.get() == 'Choose Item' or self.hours.get()=="Hours" or self.mins.get()=="Minutes" or self.entryVariable==''):
			messagebox.showwarning("Error","Please fill all fields")
		else:
			self.addItemNum()
			self.addHours()
			self.addMinutes()
			self.labelVariable = StringVar()
			label=Label(self,textvariable=self.labelVariable,anchor="center",fg="black")
			label.grid(column=3,row=0,columnspan=2,sticky='EW')
			self.labelVariable.set(u"Total Time")
			self.entry.focus_set()
			self.entry.selection_range(0, END)
			self.variable=StringVar(self)
			self.variable.set('Choose Item')
			self.variable.trace('w',self.itemChange)
			option=OptionMenu(self,self.variable,*[item for item in __items__ if item not in self.itemsSelected])
			#option['width']=15
			option.grid(column=1,row=self.row,columnspan=1,sticky='ew')
			self.row+=1
		
		
	def OnButtonClick(self):
		if self.finish['state']=='normal':
			self.labelVariable = StringVar()
			self.labelVariable.set(u"Updating information...")
			self.label=Label(self,textvariable=self.labelVariable,anchor="center",fg="black")
			self.label.grid(column=0,row=100,columnspan=1,sticky='EW')
			self.update()
			cell=importDrive(self)
			messagebox.showwarning("Totals","Your updated total for this pay period is: "+cell)
			self.destroy()


	def OnPressEnter(self,event):
		if self.finish['state']=='normal':
			self.labelVariable = StringVar()
			self.label=Label(self,textvariable=self.labelVariable,anchor="center",fg="black")
			self.label.grid(column=0,row=100,columnspan=1,sticky='EW')
			self.labelVariable.set(u"Updating information...")
			cell=importDrive(self)
			messagebox.showwarning("Totals","Your updated total for this pay period is: "+cell)
			self.destroy()

			
	def cancel(self):
		self.destroy()

class opening(Tk):
	def __init__(self,parent):
		Tk.__init__(self,parent)
		self.parent = parent
		self.minsize(width=100,height=50)
		center(self)
		self.initialize()

	def initialize(self):
		self.openVariable=StringVar()
		self.openVariable.set(u"Initializing...")
		self.openVar=Label(self,textvariable=self.openVariable,anchor='center',fg='black')
		self.openVar.place(relx=.5,rely=.6,anchor=CENTER)
		self.openVar.config(font=("Courier",12))
		myImage=PhotoImage(file='pieceworks.gif')
		self.im=Label(self,image=myImage)
		self.im.image=myImage
		#self.im.place(relx=.5,rely=.4,anchor=CENTER)
		self.im.pack()
		self.update()

def saveUsers():
	global __users__
	output = open('timesheet.pkl', 'wb')
	pickle.dump(__users__, output)
	output.close()
		
if __name__ == "__main__":
	userCheck()
	openApp=opening(None)
	openApp.title('Time Sheet Keeper')
	period,labor,items=readItems()
	global __items__
	global __labor__
	global __period__
	__period__=period
	__items__=items
	__labor__=labor
	openApp.destroy()
	app = login(None)
	app.title('Time Sheet Keeper:'+__period__)
	app.mainloop()
	saveUsers()
	
	