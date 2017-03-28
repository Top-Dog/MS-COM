'''
Created on 30/3/2016

@author: Sean D. O'Connor
'''

import sys, time, pythoncom, os, datetime
from collections import namedtuple, deque
from pywintypes import com_error
import math
import numpy as np
import pywintypes
from ..Launch import c

#################################################
####### User defined class objects ##############
#################################################
shtRange = namedtuple('xlrange', 'sheet xlrange row1 col1 row2 col2')
Cell = namedtuple('Cell', 'sheet row col')

class Sheet(object):
	"""A class for working on excel sheet objects"""
	def __init__(self, xlInstance):
		self.xlInst = xlInstance
		
	def __repr__(self):
		return "Excel Book: " + self.xlInst.xlBook.Name

	def activateSheet(self, sheetname):
		'''Set the named sheet as the currently active sheet'''
		self.xlInst.xlBook.Sheets(sheetname).Activate()
		return self.xlBook.ActiveSheet

	def renameSheet(self, newname, sheetname=''):
		'''Renames the currently active excel sheet. Excel 2010
		opens only 1 sheet at startup, so this will rename the
		default sheet.Takes a string as an argument.'''
		if sheetname:
			self.xlInst.xlBook.Sheets(sheetname).Name = newname
		else:
			self.xlInst.xlBook.ActiveSheet.Name = newname

	def getActiveSheetName(self):
		"""Returns the name of the currently active sheet"""
		return self.xlInst.xlBook.ActiveSheet.Name
		
	def getSheet(self, sheetName):
		return self.xlInst.xlBook.Worksheets(sheetName)

	def getMaxRow(self, sheetName, col, row=1):
		sht = self.xlInst.xlBook.Worksheets(sheetName)
		return sht.Cells(row, col).End(c.xlDown).Row

	def getMaxCol(self, sheetName, col, row):
		sht = self.xlInst.xlBook.Worksheets(sheetName)
		return sht.Cells(row, col).End(c.xlToRight).Column

	def getNextCol(self, sheetname, col, row):
		"""Looks for the next column with something in it. Useful
		if looking for next column over an unknown number of blank cells. """
		sht = self.xlInst.xlBook.Worksheets(sheetname)
		#self.xlInst.xlApp.SendKeys("^{DOWN}", True) # Simulates: CTRL + DOWN ARROW, block until operation completes
		nextcol = sht.Cells(row, col).End(c.xlToRight).Column
		if sht.Cells(row, col+1).Value and not sht.Cells(row, col+1).MergeCells:
			# There is an entry 1 column to the right
			nextcol = col + 1
		elif nextcol < 16384:
			# The last entry in the column
			mergedcols = 1
			while sht.Cells(row, col + mergedcols).MergeCells:
				mergedcols += 1
			nextcol = col + mergedcols

		return nextcol

	def getNextFreeCol(self, sheetname, col, row):
		"""Looks for the next column with nothing in it."""
		sht = self.xlInst.xlBook.Worksheets(sheetname)
		nextcol = sht.Cells(row, col).End(c.xlToRight).Column
		#while nextcol < 16384:
		#    if sht.Cells(row, col+1).Value or sht.Cells(row, col+1).MergeCells:
		#        # There is an entry 1 column to the right
		#        nextcol = col + 1
		#    else:
		#        break
		if nextcol >= 16384:
			nextcol = col + 1
		else:
			nextcol += 1
		if nextcol >= 16384:
			nextcol = col

		return nextcol

	def insertCol(self, sheetName, column):
		"""Insert a column to the left of the spesfied column.
		@param column: a letter code describing a column in Excel or number in the range 1 to n"""
		if type(column) == int:
			colLetter = self.num_to_let(column)
		else:
			colLetter = column
		self.xlInst.xlBook.Worksheets(sheetName).Columns("%s:%s" % (colLetter, colLetter)).Insert(Shift=c.xlToRight, CopyOrigin=c.xlFormatFromLeftOrAbove)

	def getNextRow(self, sheetName, col, row=1):
		"""Looks for the next row with something in it. Useful
		if looking for next record over an unknown number of blank cells. """
		sht = self.xlInst.xlBook.Worksheets(sheetName)
		if sht.Cells(row + 1, col).Value: #and not sht.Cells(row + 1, col).MergeCells
			return row + 1
		elif sht.Cells(row + 1, col).MergeCells:
			mergedrows = 0
			while sht.Cells(row + mergedrows, col).MergeCells:
				mergedrows += 1
			return row + mergedrows
		else:
			return sht.Cells(row, col).End(c.xlDown).Row
		
	def getNextRow2(self, SheetName, col, row=1):
		sht = self.xlInst.xlBook.Worksheets(SheetName)
		#self.xlInst.xlApp.SendKeys("^{DOWN}", True) # Simulates: CTRL + DOWN ARROW, block until operation completes
		nextrow = sht.Cells(row, col).End(c.xlDown).Row
		if sht.Cells(row + 1, col).Value and not sht.Cells(row + 1, col).MergeCells:
			# There is an entry 1 row down
			nextrow = row + 1
		elif nextrow >= 1048576:
			# The last entry in the column
			mergedrows = 1
			while sht.Cells(row + mergedrows, col).MergeCells:
				mergedrows += 1
			nextrow = row + mergedrows

		return nextrow

	def getCol(self, sheet, col, row1, row2):
		"""Return a list of values corrosponding to a column in excel"""
		newlist = self.getRange(sheet, row1, col, row2, col)
		return [item[0] for item in newlist]

	def getRow(self, sheet, row, col1, col2):
		"""Return a list of values corrosponding to row in excel"""
		newlist = self.getRange(sheet, row, col1, row, col2)
		return [item for item in newlist]

	def alpha2number(columnletters):
		"""Given a string represeenting the column in excel,
		return a number representing the column indice."""
		offset = 0
		for letter in columnletters:
			offset *= 26 # Base counting system
			offset += ord(letter.lower()) - 96 # ASCII offset of 'a'
		return offset

	def duplicate_WBO(self, newobjectname):
		"""Checks if a chart or sheet with the same
		name already exists"""
		dupFound = False
		for sheet in self.xlInst.xlBook.Sheets:
			if sheet.Name == newobjectname:
				dupFound = True 
		
		for chart in self.xlInst.xlBook.Charts:
			if chart.Name == newobjectname:
				dupFound = True
		return dupFound
		
	def remove_dup_WBO(self, wboname):
		initalstate = self.xlInst.xlApp.DisplayAlerts
		self.xlInst.xlApp.DisplayAlerts = False
		if self.duplicate_WBO(wboname):
			for sheet in self.xlInst.xlBook.Sheets:
				if sheet.Name == wboname:
					sheet.Delete() 
		
			for chart in self.xlInst.xlBook.Charts:
				if chart.Name == wboname:
					chart.Delete()
		self.xlInst.xlApp.DisplayAlerts = initalstate

	def addSheet(self, sheetname):
		#self.xlBook.Sheets.Add(None, After=self.xlInst.xlBook.Sheets([sht for sht in self.xlInst.xlBook.Sheets][-1].Name)) # Adds after/to the right of the last workbook sheet
		self.xlInst.xlBook.Sheets.Add(None, After=self.getLastWBO("Sheets"))
		self.xlInst.number_sheets += 1
		self.getLastWBO("Sheets", sheetname)
		#[sht for sht in self.xlInst.xlBook.Sheets][-1].Name = sheetname
		
	def addChart(self, chartname, chartType):
		"""Adds a new chart object (as a sheet), to the right of all existing sheets"""
		self.remove_dup_WBO(chartname)
		#lastChart = self.xlInst.getLastWBO("Charts")
		lastSheet = self.getLastWBO("Sheets") # gets both worksheet and chart objects aka "sheets"
		
		# Need a two step approach, else adding a chart to the end of a workbook doesn't work
		chart = self.xlInst.xlBook.Charts.Add()
		self.xlInst.xlBook.Sheets(chart.Name).Move(After=lastSheet)
		
		if type(chartType) is str:
			if os.path.isfile(chartType) and chartType.endswith('.crtx'):
				chart.ApplyChartTemplate(chartType)
			else:
				raise TypeError, "chartType is not recognised as a standard excel chart, or template does not exist" 
		else:
			chart.ChartType = chartType # This needs to be seperate, else it doesn't work (also two variations of the property: .ChartType and .Type)
		#self.xlInst.number_charts += 1
		# Rename the new chart/sheet
		lastChart = self.getLastWBO("Charts", chartname) 

		return lastChart

	def getLastWBO(self, WBO, objectName=""):
		"""Get Last Work Book Object. 
		Objects include: "Sheets", "Charts"
		@return: a name of the last WBO, or None if nothing found"""
		lastWBO = None
		if WBO == "Sheets":
			lastWBO = self.xlInst.xlBook.Sheets([sht for sht in self.xlInst.xlBook.Sheets][-1].Name)
		elif WBO == "Charts":
			ChartList = [chrt for chrt in self.xlInst.xlBook.Charts]
			if ChartList:
				lastChart = ChartList[-1]
				lastWBO = self.xlInst.xlBook.Charts(lastChart.Name)
			
		# Optionally, set the Name of the last object
		if objectName and lastWBO:
			lastWBO.Name = objectName
		return lastWBO
		
	def addWorkbook(self):
		self.xlInst.xlBook = self.xlInst.xlApp.Workbooks.Add()

	def num_to_let(self, num):
		'''Converts a number index to a letter index in excel.
		Uses reccurssion. 1 origin (not 0 origin)'''
		num -= 1 # comment this line to use 0 origin
		baseA = 65 # where 'A' is in the ascii table
		if num > 25:
			return self.num_to_let(num/26) + chr(baseA + (num % 26))
		return chr(baseA+num)
		
	def _rmvSheet(self, sheetname=None):
		if sheetname:
			self.xlInst.xlApp.DisplayAlerts = False
			try:
				self.xlInst.xlBook.Sheets(sheetname).Delete()
			except com_error:
				print "Sheet does not exist"
			self.xlInst.xlApp.DisplayAlerts = True
			self.xlInst.number_sheets -= 1
		else:
			for sht in self.xlInst.xlBook.Sheets:
				#if sht.Name not in self.xlInst.baseSheets:
				self.xlInst.xlApp.DisplayAlerts = False
				sht.Delete()
				self.xlInst.xlApp.DisplayAlerts = True
				self.xlInst.number_sheets -= 1
					
	def rmvSheet(self, removeList=[], keepList=[]):
		"""Removes all sheets in the removeList, if they exist,
		does not remove a sheet if it appears in the keepList
		(even if it also appears in the removeList"""
		if len(removeList) == 0 and len(keepList) == 0:
			for sht in self.xlInst.xlBook.Sheets:
				#if sht.Name not in self.xlInst.baseSheets:
				self.xlInst.xlApp.DisplayAlerts = False
				sht.Delete()
				self.xlInst.xlApp.DisplayAlerts = True
				self.xlInst.number_sheets -= 1
		elif len(keepList) == 0: # Only remove elements in remove list
			for sht in self.xlInst.xlBook.Sheets:
				if sht.Name in removeList:
					self.xlInst.xlApp.DisplayAlerts = False
					sht.Delete()
					self.xlInst.xlApp.DisplayAlerts = True
					self.xlInst.number_sheets -= 1
					#removeList.remove(sht.Name)
		elif len(removeList) == 0: # Remove all sheets except those note in keepList and baseSheets
			localKeepList = keepList #+ self.xlInst.baseSheets
			for sht in self.xlInst.xlBook.Sheets:
				if sht.Name not in localKeepList:
					self.xlInst.xlApp.DisplayAlerts = False
					sht.Delete()
					self.xlInst.xlApp.DisplayAlerts = True
					self.xlInst.number_sheets -= 1
		else:
			# There are items in both the keep and remove list
			for sht in self.xlInst.xlBook.Sheets:
				if sht.Name in removeList and sht.Name not in keepList:
					self.xlInst.xlApp.DisplayAlerts = False
					sht.Delete()
					self.xlInst.xlApp.DisplayAlerts = True
					self.xlInst.number_sheets -= 1
		
	def printSheetNames(self):
		for i in range(self.number_sheets):
			print self.xlInst.xlBook.Sheets(i).Name
			
	def getCell(self, sheetName, row, col):
		"Get VALUE of one cell"
		sht = self.xlInst.xlBook.Worksheets(sheetName)
		return sht.Cells(row, col).Value

	def _getCell(self, sheetName, row, col):
		"Get cell object"
		sht = self.xlInst.xlBook.Worksheets(sheetName)
		return sht.Cells(row, col)
		#return sht.Range(self.xlInst.let_to_num(row) + str(col))
		
	def mv_cell_ref(self, activecell, offsetx, offsety):
		"""
		Moves an excel cell object around
		"""
		return activecell.Offset(offsety + 1, offsetx + 1)

	def setCell(self, sheetName, row, col, value=None):
		"Set value of one cell. If no value is supplied the cell is cleared"
		sht = self.xlInst.xlBook.Worksheets(sheetName)
		#sht = self.xlInst.xlApp.Sheets(sheetname)
		sht.Cells(row, col).Value = value

	def get_calculation_mode(self):
		"""
		@return: the present value of the excel instance's calculation mode 
		"""
		return  self.xlInst.xlApp.Calculation

	def set_calculation_mode(self, state):
		if state is "manual":
			self.xlInst.xlApp.Calculation = c.xlCalculationManual
		elif state is "automatic":
			self.xlInst.xlApp.Calculation = c.xlCalculationAutomatic
		elif state is "semi automatic":
			self.xlInst.xlApp.Calculation = c.xlCalculationSemiautomatic 
		
	def setRow(self, sheet, row, col, valueArray, rowrange=''):
		"""
		Sets a row of values a range
		@param valueArray: A list of comma separated cell values
		"""
		sht = self.xlInst.xlBook.Worksheets(sheet)
		# rowrange is excel range object
		if rowrange:
			sht.Range(rowrange).Value = valueArray
		else:
			sht.Range(sht.Cells(row, col), sht.Cells(row, col+len(valueArray)-1)).Value = valueArray

	def FillDown(self, sheetName, setrow, n, row, col):
		'''Auto Fills a row down n columns'''
		sht = self.xlInst.xlBook.Worksheets(sheetName)
		self.xlInst.setRow(sheetName, row, col, setrow)
		rangestr = "%s%d:%s%d" % (self.num_to_let(col), row, self.num_to_let(col + len(setrow) - 1), row + n)
		sht.Range(rangestr).FillDown()
		
	def mergeCells(self, sheetName, row1, col1, row2, col2):
		sht = self.xlInst.xlBook.Worksheets(sheetName)
		sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Merge()

	def mergeRange(self, sheetName, mergerange):
		# Range is like "A1:B2"
		sht = self.xlInst.xlBook.Worksheets(sheetName)
		sht.Range(mergerange).Merge()
		
	def autofit(self, sheetName, column):
		'''Auto adjusts the column widths to achieve the best fit'''
		sht = self.xlInst.xlBook.Worksheets(sheetName)
		sht.Columns(self.num_to_let(column)).AutoFit()

	# Excel and COM define a date as the number of days since 1/1/1900
	# Python and Unix define a date as the number of seconds since 1/1/1900
	# Excel automatically formats date objects correctly, so they don't appear as floats
	# Excel uses datetime objects, so we need to convert pytime to datetime
	def getPyDateTime(self, dTDate):
		"""
		@param dTDate: A python datetime object
		@return: a pytime datetime object suitable for use with COM or Excel
		"""
		# Ignore the error in eclipse about not being able to locate the function
		return pythoncom.MakeTime(dTDate)
		
	def getDateTime(self, pyDate):
		"""
		@param pyDate: A date time object from COM or Excel
		@return: a conventional python datetime object
		"""
		return datetime.datetime(pyDate.year, pyDate.month, pyDate.day,
						 pyDate.hour, pyDate.minute, pyDate.second,
						 pyDate.msec / 1000)
		
	def str2DateTime(self, dTStr):
		"""Convert a datetime string to a datetime object"""
		t = time.strptime(dTStr, '%d/%m/%Y %H:%M:%S') # Value error is often associated with month and day being swapped
		return datetime.datetime(t.tm_year, t.tm_mon, t.tm_mday, t.tm_hour, t.tm_min, t.tm_sec)

	def getRangeFormula(self, sheet, row1, col1, row2, col2):
		"return a 2d array of values (i.e. tuple of tuples)"
		sht = self.xlInst.xlBook.Worksheets(sheet)
		return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Formula

	def getRange(self, sheet, row1, col1, row2, col2):
		"return a 2d array of values (i.e. tuple of tuples)"
		sht = self.xlInst.xlBook.Worksheets(sheet)
		return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value

	def _getRange(self, sheet, row1, col1, row2, col2):
		"return an excel range object"
		sht = self.xlInst.xlBook.Worksheets(sheet)
		return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2))

	def transposeRange(self, data):
		"""Transposes python lists to/from excel"""
		NoRows = len(data)
		try:
			NoCols = len(data[0])
		except:
			NoCols = 0
			
		TransposedRange = []
		for col in range(NoCols):
			DataRow = []
			for row in range(NoRows):
				DataRow.append(data[row][col])
			TransposedRange.append(DataRow)
		
		return TransposedRange

	def setRange(self, sheet, topRow, leftCol, data):
		"""Insert a 2d array starting at given location.
		Works out the size needed for itself. Needs to be
		a list of lists to insert columns."""
		bottomRow = topRow + len(data) - 1
		try:
			rightCol = leftCol + len(data[0]) - 1
		except TypeError:
			rightCol = leftCol
		sht = self.xlInst.xlBook.Worksheets(sheet)
		sht.Range(sht.Cells(topRow, leftCol),
				   sht.Cells(bottomRow, rightCol)).Value = data
		#return bottomRow, rightCol
		return bottomRow+1, rightCol+1
		
	def getContiguousRange(self, sheet, row, col, direction="both"):
		"""Tracks down and across from top left cell until it
		encounters blank cells; returns the non-blank range.
		Looks at first row and column; blanks at bottom or right
		are OK and return None within the array"""

		sht = self.xlInst.xlBook.Worksheets(sheet)

		# find the bottom row - time consuming (very slow)
		#bottom = row
		#while sht.Cells(bottom + 1, col).Value not in [None, ""]:
		#    bottom = bottom + 1

		# right column - time consuming (very slow)
		#right = col
		#while sht.Cells(row, right + 1).Value not in [None, ""]:
		#    right = right + 1
		
		if direction == "row":
			#xlrange = sht.Range(sht.Cells(row, col), sht.Cells(row, right)).Value
			xlrange = sht.Range(sht.Cells(row, col), sht.Cells(row, col).End(c.xlToRight)).Value
		elif direction == "col":
			#xlrange = sht.Range(sht.Cells(row, col), sht.Cells(bottom, col)).Value
			xlrange = sht.Range(sht.Cells(row, col), sht.Cells(row, col).End(c.xlDown)).Value
		else:
			#xlrange = sht.Range(sht.Cells(row, col), sht.Cells(bottom, right)).Value
			xlrange = sht.Range(sht.Cells(row, col), sht.Cells(row, col).End(c.xlToRight).End(c.xlDown)).Value
		return xlrange

	def fixStringsAndDates(self, aMatrix):
		# converts all unicode strings and times
		newmatrix = []
		for row in aMatrix:
			newrow = []
			for cell in row:
				if type(cell) is UnicodeType: # types.UnicodeType
					newrow.append(str(cell))
				elif type(cell) is TimeType:
					newrow.append(int(cell))
				else:
					newrow.append(cell)
			newmatrix.append(tuple(newrow))
		return newmatrix

	def search(self, shtRange, searchTerm, **kwargs):
		# At some point this function was modified to no longer except excel ranges as valid ranges, instead supply row and col details
		'''Search through the spreadsheet for searchTerm [exact match]. Returns
		a list of cell objects for all the matches found.'''
		sht = self.getSheet(shtRange.sheet)
		if shtRange.xlrange is not None:
			r1 = shtRange.xlrange
		else:
			r1 = sht.Range(sht.Cells(shtRange.row1, shtRange.col1), sht.Cells(shtRange.row2, shtRange.col2))
		cell = r1.Find(What=searchTerm, LookAt=c.xlWhole, MatchCase=kwargs.get("MatchCase", False))
		searchResults = [] # store a list of cells that match the search criteria
		
		if cell:
			try:
				cell.Address
			except:
				print "In search: ", searchTerm, cell, type(cell)
				exit()
			firstAddr = cell.Address
			searchResults.append(cell)
			while 1:
				cell = r1.FindNext(cell)
				if cell.Address == firstAddr:
					break
				searchResults.append(cell)
		return searchResults

	def brief_search(self, SheetName, SearchTerm):
		"""
		@return: An excel cell object (attributes of Row, Column, Value...) or None if no matches are found
		Search for the FIRST instance of a searchterm
		"""
		SearchTermOffset = 0 # returns the 1st search result
		SearchResults = self.search(shtRange(SheetName, None, 1, 1, 1000, 1000), 
								SearchTerm)
		if len(SearchResults):
			return SearchResults[SearchTermOffset]
		else:
			return None
		
	def column_search(self, SheetName, SearchTerms):
		"""
		@param SearchTerms: A list of the column names/text to find (in the same column)
		@return: A tuple of the (x,y) coordinates of the cells NB: this is (col,row)
		"""
		ColumnList = []
		LastRow = self.xlInst.brief_search(SheetName, SearchTerms[0]).Row
		for term in SearchTerms:
			result = self.brief_search(SheetName, term)
			assert result is not None, "Error: can not find the column (%s)" % term
			assert result.Row == LastRow, "Error: search tags span multiple rows"
			LastRow = result.Row
			ColumnList.append((result.Column, result.Row))
			#ColumnList.append(result.Column)
		return ColumnList

	def find_all(self, SheetName, SearchTerm):
		'''Finds all the cells that match  search string (wild card is asterix)
		and return them as list of cell references.'''
		results = self.xlInst.search(shtRange(SheetName, None, 1, 1, 10000, 10000), 
							  SearchTerm)

	def checkEqual(self, iterator):
		'''Returns True if all items in the iterable/list are identical when hashed,
		False otherwise. The problem is that itterable objects that include 2 or more 
		identical elements followed by uninitialised elements will will return True.'''
		return len(set(iterator)) <= 1 and len(iterator) >= 2
			
	def sort(self, sheet, row=None, col=None, rule=None):
		'''Returns an array of the supplied range, sorted according to the rule''' 
		pass

	def roundUp(self, x):
		'''Rounds x up to the nearest 10'''
		return int(math.ceil(x / 10.0)) * 10

	def localMaxima(self, yval):
		'''Finds all local maxima of a dataset. The output is an array of indexes where the local maxima occur
		http://stackoverflow.com/questions/17907614/finding-local-maxima-of-xy-data-point-graph-with-numpy''' 

		yval = np.asarray(yval)
		gradient = np.diff(yval)
		maxima = np.diff((gradient > 0).view(np.int8))
		return np.concatenate((([0],) if gradient[0] < 0 else ()) + 
							  (np.where(maxima == -1)[0] + 1,) + 
							  (([len(yval)-1],) if gradient[-1] > 0 else ()))

	def newGraph(self, feederID, xValues, titles):
		zoneSubID = feederID[:3]
		sht = self.getSheet(feederID=feederID)
		
		shape = sht.Shapes.AddChart()
		chart = shape.Chart
		for seriesIndex in range(1, chart.SeriesCollection().Count + 1):
			chart.SeriesCollection(1).Delete()

		if xValues == "indexes":
			col = 1
			row = 1
			while sht.Cells(row, col).Value not in [feederID]:
				col += 1
			
			r1_name = self.xlInst.getCell(zoneSubID, row, col)
			r1 = self.xlInst.getContiguousRange(zoneSubID, row+1, col, direction="col")
			r1 = sorted([x[0] for x in r1])
			
			graphSheet = self.xlInst.activateSheet("Graphs")
			if graphSheet.Cells(1, 2).Value not in (None, ""):
				col = graphSheet.Cells(1, 1).End(c.xlToRight).Column + 1
			elif graphSheet.Cells(1, 1).Value not in (None, ""):
				col = 2
			else:
				col = 1
			# Create a column of ordered currents
			self.xlInst.setCell("Graphs", row, col, r1_name)
			row1, col1 = self.xlInst.setRange("Graphs", row+1, col, [[I] for I in r1])
			
			sht = self.xlInst.activateSheet(zoneSubID) # reactivate the old sheet that represent the zone sub
			
			end = graphSheet.Cells(2, col).End(c.xlDown).Row
			src = graphSheet.Range(graphSheet.Cells(2, col), graphSheet.Cells(end, col))
			r2 = sht.Range(sht.Cells(2, col), sht.Cells(2, col).End(c.xlDown)) # CT current readings

		elif xValues == "dates":
			col = sht.Cells(1, 1).End(c.xlToRight).Column
			r1 = sht.Range(sht.Cells(2, 1), sht.Cells(2, 1).End(c.xlDown)) # dates
			r2 = sht.Range(sht.Cells(2, col), sht.Cells(2, col).End(c.xlDown)) # CT current readings
			src = self.xlInst.xlApp.Union(r1, r2)

		chart.ChartType = c.xlLine #c.xlXYScatter
		chart.SeriesCollection().Add(Source=src)
		
		series = chart.SeriesCollection(1)
		chart.HasLegend = False
		series.Name = titles.title
		
		# Y axis
		try:
			chart.Axes(c.xlValue).MaximumScale = self.xlInst.roundUp(max(series.Values)) # y axis configuration
			chart.Axes(c.xlValue).MinimumScale = 0
		except:
			print "Error: can't set axis scale for %s" % feederID # weird bug here when running till UND4 - offset by 5 feeders

		try:
			chart.Axes(c.xlValue).MajorUnit = chart.Axes(c.xlValue).MaximumScale / 10 # Resolution
		except com_error:
			print "Got a COM error: ", com_error 
		
		chart.Axes(c.xlCategory, c.xlPrimary).HasTitle = True # xlCategory , xlPrimary (1, 1)
		chart.Axes(c.xlCategory, c.xlPrimary).AxisTitle.Text = titles.xAxis
		chart.Axes(c.xlValue, c.xlPrimary).HasTitle = True # xlValue, xlPrimary (2, 1)
		chart.Axes(c.xlValue, c.xlPrimary).AxisTitle.Text = titles.yAxis
		return [x[0] for x in r2.Value] # return 
		
	def addSeries(self, shtName, chartNumber, xval, yval):
		'''Adds a new series to an exsisting chart'''
		sht = self.getSheet(feederID=shtName)
		shape = sht.Shapes(chartNumber)
		chart = shape.Chart
		numSeries = chart.SeriesCollection().Count
		
		chart.SeriesCollection().Add(Source=sht.Range(sht.Cells(2,1), sht.Cells(3,1))) # Load a dummy range
		series = chart.SeriesCollection(numSeries+1)
		#assert len(xval) > 16384 or len(yval) > 16384, "The max. number of elements assigned to a series has been exceeded"
		series.XValues = xval # limited to arrays of a max len 16384 (2^14)
		series.Values = yval # limited to arrays of a max len 16384 (2^14)

	def arrangeCharts(self, feederID, widths, heights, offset, chartOffset=0):
		'''Arranges the charts in block grid fashion'''
		sht = self.getSheet(feederID=feederID)
		numCharts = sht.ChartObjects().Count
		
		#width = 360 # Width to set each chart
		#height = 220 # Height to set each chart
		#numWide = 3 # The number of charts wide before starting a new line
		if not chartOffset:
			start = 1
			numWide = len(widths)
		else:
			start = chartOffset
			numWide = 1
		for chartIndex in range(start, numCharts + 1):
			if chartIndex % 2 or start > 1: 
				# odd, so it's an "date" chart
				width = widths[0]
				height = heights[0]
			else:
				# even, so it's an "index" chart
				width = widths[1]
				height = heights[1]
				
			widthPrevious = width
			heightPrevious = height
			try:
				chartPrevious = sht.ChartObjects(chartIndex-1)
				widthPrevious = chartPrevious.Width
				heightPrevious = chartPrevious.height
			except:
				pass

			chart = sht.ChartObjects(chartIndex)
			chart.Width = width
			chart.Height = height
			chart.Left = ((chartIndex - start) % numWide) * widthPrevious + offset.x # normally: (chartIndex - 1)....
			chart.Top = int((chartIndex - start) / numWide) * heightPrevious + offset.y
		

class Bunch:
	"""
	A utility class. Usage:
	>>> point = Bunch("class", datum=y, squared=y*y, coord=x)
	>>> point.isok = 1
	"""
	def __init__(self, refname, **kwds):
		self.refname = refname
		self.__dict__.update(kwds)
		
	def __repr__(self):
		return self.refname
