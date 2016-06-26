'''
Created on 30/1/2016

@author: Sean D. O'Connor

A self-contained "launch" module for MS Excel.
This module is spesfic to Excel because Excel behaves 
differently to other MS Office applications in the sense 
that new workbooks (documents) are opend in an existing
instance by default. This is not always desireable (particualrly
when threading or multiprocessing), so this module
allows manual control over Excel's starting behaviour.

There are also methods to manage workbook objects, such
as saving, deletion, and closing the running instance.
'''

import sys, time, pythoncom, os, datetime
from collections import namedtuple, deque
from pywintypes import com_error
import math
import numpy as np
import pywintypes

from win32com.client.gencache import EnsureDispatch
from win32com.client import constants as c, Dispatch, DispatchEx, GetObject

#################################################
####### Make sure the PYTHONPATH is setup #######
#################################################
sys.path.append(r'C:\Python27\lib\site-packages\win32com')
# PS H:\> [System.Environment]::SetEnvironmentVariable("PYTHONPATH", "C:\Python27;C:\Python27\DLLs;C:\Python27\Lib\lib-tk;
# C:\Users\sdo\Documents\eclipse workspace\Zone Substation Trending;","User")

#################################################
####### Excel spesfic application launcher ######
#################################################
class Excel(object):
	"""A utility class to make it easier to get at Excel and manage 
	a single (1) running instance from one class."""

	def __init__(self, **kwargs):
		"""
		xlApp: instacne of the excel application
		xlBook: the active excel workbook
		"""
		visible = kwargs.get("visible", True) # The state of the current xl window
		self.number_sheets = 1 # By default excel will open with one sheet
		AppName = 'Excel.Application'
		self.xlApp = None # The window frame
		self.xlBook = None # The name of te file open inside the window (there can be many)
		self.filename = None
		
		if kwargs.get('runninginstance'):
			# We have to attach to an existing application instance
			try:
				filename = kwargs.get('filename')
				# GetObject(Class='Excel.Application') will get the first Excel instance over and over again
				self.xlBook = GetObject(filename) #will return a running instance of the file/program (if exists), else starts a new instance
				#self.xlApp = GetObject(Class=AppName)
				self.xlApp = kwargs.get('instance', self.xlBook.Application) # The only way to access the parent instnace is to get it from the proccess that started it
			except TypeError:
				print "Could not attach to any %s" % AppName
				return
			except:
				# Produces a com_error if the file is not avliable
				self.xlApp = GetObject(Class=AppName)
				self.xlBook = self.xlApp.ActiveWorkbook
		else:
			try:
				if kwargs.get('newinstance'):
					# Create a new instance of the application
					if kwargs.get('earlybinding') or kwargs.get('staticproxy'):
						self.xlApp = EnsureDispatch(AppName)
					else:
						self.xlApp = DispatchEx(AppName) # Will open a read only copy, if the file is already open
						
					if kwargs.get('filename'):
						try:
							self.filename = kwargs.get('filename')
							self.xlBook = self.xlApp.Workbooks.Open(self.filename)
						except:
							# The file does not exist
							pass
							#self.addWorkbook() # Binds a new xl workbook
							#self.xlBook = self.xlApp.ActiveWorkbook
				else:
					pass
					#Handled by line: "if kwargs.get('filename'):" for more details
					#self.xlApp = Dispatch(AppName)
			except TypeError:
				print "Could not dispatch %s" % AppName
				return
			
			if kwargs.get('filename') and self.xlBook is None:
				self.filename = kwargs.get('filename')
				try:
					#self.xlBook = self.xlApp.Workbooks.Open(self.filename)
					self.xlBook = GetObject(self.filename) # This looks like it is opening new files in an existing instance regardless of the newinstance variable being set
					self.xlApp = self.xlBook.Application
					#self.xlApp.Windows(self.xlBook.Name).Visible = kwargs.get("BookVisible", True) # The visibility of the workbook in the instace of excel
				except com_error:
					# The file does not exist...
					if self.xlApp:
						# A new instance of excel was launched, but it has not loaded any workbooks, so ActiveWorkbook is None
						self.xlApp.Workbooks.Add()
						self.xlBook = self.xlApp.ActiveWorkbook
					else:
						# A new instance was not started, and the file doesn't exist, so start a new workbook in the running instace
						self.xlApp = Dispatch(AppName)
						self.xlApp.Workbooks.Add() # Binds a new xl workbook
						self.xlBook = self.xlApp.ActiveWorkbook
			elif self.xlApp is None:
				self.xlApp = Dispatch(AppName)
				self.xlApp.Workbooks.Add() # Binds a new xl workbook
				self.xlBook = self.xlApp.ActiveWorkbook
			
			# Launch the iHistorian Excel plug-in separately (no plug-ins are launched with the Excel COM object)
			try:
				self.launchiHistorian()
			except:
				pass
		
		self.xlApp.Windows(self.xlBook.Name).Visible = kwargs.get("BookVisible", True) # The visibility of the workbook in the instace of excel
		self.xlApp.Visible = visible
		self.FinalTests()
		
		# (win32com.client)
		# Dispatch
		# Starts a new program process, but doesn't neccessarialy create a instance of the application
		
		# DispatchEx
		# Starts a new instance of the program (allows you to specify if you want the process on this machine or another one, in/out of proccess etc.)
		# Historian add-in not available
		
		# gencache.EnsureDispatch
		# Ensures a static proxy exists by either creating it or returning an existing one
		# Use in place of Dispatch() if you always want early binding
		
		# GetObject(filename, Class="")
		# Attaching to an existing application
		# GetObject attaches to the last opened instance (using GetActiveObject), else starts a new instance (if no other instances of the program e.g. excel are open, else it binds to an existing excel instance)
		
	def __repr__(self):
		filename = self.filename
		if not self.filename:
			filename = "Excel: " + str(self.xlApp)
		return filename
	
	def launchiHistorian(self):
		"""This is for loading the GE Proficy Historian Excel Add-in"""
		iHistorianPath = r"C:\Program Files (x86)\Microsoft Office\Office14\Library\iHistorian.xla"
		
		self.xlApp.DisplayAlerts = False
		assert os.path.exists(iHistorianPath) == True, "Error: There was a problem locating the iHistorian.xla Excel plug-in."
		self.xlApp.Workbooks.Open(iHistorianPath)
		self.xlApp.RegisterXLL(iHistorianPath)
		self.xlApp.Workbooks(iHistorianPath.split("\\")[-1]).RunAutoMacros = True
		self.xlApp.DisplayAlerts = True

	def FinalTests(self):
		"""Ensure a static proxy exists"""
		try:
			c.xlDown
		except:
			from win32com.client import makepy
			sys.argv = ["makepy", r"C:\Program Files (x86)\Microsoft Office\Office14\Excel.exe"]
			makepy.main ()
		
		
	def save(self, newfilepath=None):
		IntialAlertState = self.xlApp.DisplayAlerts
		self.xlApp.DisplayAlerts = False
		if newfilepath:
			self.filename = newfilepath
			directory = os.path.dirname(newfilepath)
			if not os.path.exists(directory):
				os.makedirs(directory)
			self.xlBook.SaveAs(newfilepath)
		else:
			if os.path.exists(self.filename):
				#os.remove(self.filename)
				self.xlBook.Save()
			else:
				directory = os.path.dirname(self.filename)
				if not os.path.exists(directory):
					os.makedirs(directory)
				self.xlBook.SaveAs(self.filename)
		self.xlApp.DisplayAlerts = IntialAlertState # Probably be set back to true, but maybe not

	def closeWorkbook(self, SaveChanges=0):
		self.xlApp.Windows(self.xlBook.Name).Visible = True
		self.xlBook.Close(SaveChanges) # Avoids a prompt when closing out
		#self.xlApp.Quit()
		#self.xlApp.Windows(self.xlBook.Name).Visible = 0
		self.xlBook = None
		del self.xlBook # Free-up memory
		
	def closeApplication(self):
		"""Closes the entire window/running instance, 
		including any open workbooks."""
		numberOfBooks = 0
		for book in self.xlApp.Windows:
			numberOfBooks += 1
			self.xlApp.Windows(self.xlBook.Name).Visible = True
			book.Close(SaveChanges=0)
		#self.xlBook.Close(SaveChanges=0) # Avoids a prompt when closing out

		self.xlApp.Visible = 0 # Must do this, else the excel.exe process does not quit
		self.xlApp.Quit()
		self.xlApp = None
		del self.xlApp # Free-up memory
		
	def closeInstance(self):
		"""Closes the currently attached instanace, but leaves the window running"""
		self.xlApp = None
		del self.xlApp


class Instances(object):
	def __init__(self, name):
		self.appname = name
		self.filepath = ""


class Excel_Handler(object):
	Instances = []

# OOP: https://jeffknupp.com/blog/2014/06/18/improve-your-python-python-classes-and-object-oriented-programming/
class Excel_Instance(Excel_Handler):
	def __init__(self, **kwargs):
		self.xlAPP = None

		app_id = kwargs.get("index", 0)
		app_book = kwargs.get("book", None)

		if app_book:
			xlBook = GetObject(os.path.join(os.path.abspath(__file__), app_book))
			self.xlAPP = xlBook.Application
		elif app_id > 0:
			pass
		else:
			self.xlAPP = DispatchEx("Excel.Application")

		# Set the visibility of the new instance, defaults to visible
		self.set_visibility(kwargs.get("visible", True))

		# See if we find the current book in any of the excel bokks that are open
		for xlapp in xlAPPs:
			for xlbook in xlapp.Workbooks:
				if xlbook.Name == app_book:
					pass


	def __repr__(self):
		return "Excel: %s" % self.name

	#def number_instances(self):
	#	return 0

	def set_visibility(self, status):
		self.xlAPP.Visible = status

	def close(self):
		pass

class Excel_Book(): # Excel_Instance
	def __init__(self):
		self.xlBOOK = None

	def new_book(self, filename):
		pass

	def get_book(self, filename):
		pass

	def set_visibility(self, status):
		self.xlBOOK.Visible = status

	def close(self):
		pass