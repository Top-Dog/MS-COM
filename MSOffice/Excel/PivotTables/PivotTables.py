'''
Created on 08/3/2016

@author: Sean D. O'Connor
'''


from ..Launch import c
from ..Worksheets.Worksheet import Sheet,shtRange
from collections import namedtuple

#################################################
####### User defined class objects ##############
#################################################


class PivotTables(object):
	"""A class for working on excel pivot table objects"""
	def __init__(self, xlInstance,xlSheet,TableName,SheetName):
		self.xlInst = xlInstance
		self.xlSheet = xlSheet
		self.TableName = TableName
		self.SheetName = SheetName
		
	def __repr__(self):
		return "Excel Book: " + self.xlInst.xlBook.Name
		
	def createpivot(self,sourcerange,destrange):
		#sheet = self.xlInst.xlBook.Sheets(SheetName)
		PivotSource = self.xlSheet._getRange(sourcerange.sheet, sourcerange.row1, sourcerange.col1, sourcerange.row2, sourcerange.col2)
		PivotDest = self.xlSheet._getRange(destrange.sheet, destrange.row1, destrange.col1, destrange.row2, destrange.col2)
		

		pcache = self.xlInst.xlBook.PivotCaches().Create(SourceType=c.xlDatabase, SourceData = PivotSource, Version=6)
		
		pcache.CreatePivotTable(TableDestination=PivotDest, TableName=self.TableName, DefaultVersion=6)
	
	
	def addpivotfields(self,filtertype,filtername,filterlist):
		
		filtertype = filtertype.capitalize()
		
		
		oridic = {'Filter': c.xlPageField,'Column': c.xlColumnField,'Row': c.xlRowField,'Values': c.xlDataField}
		
		
		if filtertype not in oridic:
			print "%s is not a valid fieldtype. Please select from:" % filtertype
			print oridic.keys()
			return False
		else:
			orientation = oridic[filtertype]
		
		
		sheet = self.xlInst.xlBook.Sheets(self.SheetName)
		
		if filterlist != "All":
			self.togglepivotitems(filtername,filterlist)
		
		sheet.PivotTables(self.TableName).PivotFields(filtername).Orientation = orientation
		
		return True
			
		
		
	def addpivotvalues(self,filtername,filterlist,summary = "Count"):
		
		sheet = self.xlInst.xlBook.Sheets(self.SheetName)
		
		
		
		sheet.PivotTables(self.TableName).PivotFields(filtername).Orientation = c.xlDataField
		
		listoffields = self.listpivotfields("Values")
		newfiltername =  filter(lambda x: filtername in x, listoffields)[0]
		
		
		sheet.PivotTables(self.TableName).PivotFields(newfiltername).Caption = filtername + " "	# Requires a trailing space becuase Excel is dumb
		
		if summary == "Sum":
			sheet.PivotTables(self.TableName).PivotFields(filtername + " ").Function = c.xlSum
			filtername = "Sum of " + filtername
		else:
			sheet.PivotTables(self.TableName).PivotFields(filtername + " ").Function = c.xlCount
			filtername = "Count of " + filtername
		
		self.togglepivotitems(filtername,filterlist)
			
	def listpivotfields(self,FieldType):
		table = self.xlInst.xlBook.Sheets(self.SheetName).PivotTables(self.TableName)
		
		fields = {'Filter': table.PageFields,'Column': table.ColumnFields,'Row': table.RowFields,'Values': table.DataFields,'Visible': table.VisibleFields,'Invisible': table.HiddenFields}
		
		if FieldType not in fields:
			print "%s is not a valid fieldtype. Please select from:" % FieldType
			print fields.keys()
			return False
		else:
		
		
			fieldarea = fields[FieldType]
		
		
		fieldlist = []
		for pvtField in fieldarea:
			fieldlist.append(pvtField.Name)
		
		return fieldlist
	
	def listpivotitems(self,FieldName):
		
		pivotfield = self.xlInst.xlBook.Sheets(self.SheetName).PivotTables(self.TableName).PivotFields(FieldName)
		
		pvtitemlist = []
		for pvtitem in pivotfield.PivotItems():
			pvtitemlist.append(pvtitem.Name)
		return pvtitemlist
		
	def listvisiblepivotitems(self,FieldName):
		
		pivotfield = self.xlInst.xlBook.Sheets(self.SheetName).PivotTables(self.TableName).PivotFields(FieldName)
		
		pvtitemlist = []
		for pvtitem in pivotfield.PivotItems():
			if pvtitem.Visible:
				pvtitemlist.append(pvtitem.Name)
		return pvtitemlist
			
	def listinvisiblepivotitems(self,FieldName):
		
		pivotfield = self.xlInst.xlBook.Sheets(self.SheetName).PivotTables(self.TableName).PivotFields(FieldName)
		
		pvtitemlist = []
		for pvtitem in pivotfield.PivotItems():
			if not pvtitem.Visible:
				pvtitemlist.append(pvtitem.Name)
		return pvtitemlist
	
	def togglepivotitems(self,FieldName,pvtlist):
			
		pivotfield = self.xlInst.xlBook.Sheets(self.SheetName).PivotTables(self.TableName).PivotFields(FieldName)
		
		novisibleitems = len(self.listvisiblepivotitems(FieldName))	# Will be used to ensure not all pvt items made invisible
		
		for pvtitem in pivotfield.PivotItems():
			
			if pvtitem.Name not in pvtlist and novisibleitems != 1:
				pvtitem.Visible = False
				novisibleitems -= 1
				
		
		
		
		

	
		
