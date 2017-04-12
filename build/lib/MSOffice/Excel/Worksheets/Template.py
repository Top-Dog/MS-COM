import MSOffice, time
from MSOffice.Excel.Launch import c
from Worksheet import Sheet, shtRange

class Excel(object):
	def __init__(self, cell_ref):
		self.Instance = cell_ref.Parent.Parent.Parent
		self.Workbook = cell_ref.Parent.Parent
		self.Worksheet = cell_ref.Parent # cell_ref.Worksheet

		# Compatability (coupled to Excel Luanch.py)
		self.xlApp = self.Instance
		self.xlBook = self.Workbook
		self.filename = None

class Template(object):
	"""Copy tables from a existing source document an place them
	in a working document."""
	def __init__(self, TemplateFileName):
		self.xlSheetDest = None
		self.TemplateFileName = TemplateFileName
		self.Dest_Row = 0
		self.Dest_Col = 0
		self.Size_Rows = 0
		self.Size_Cols = 0

	def Place_Template(self, template_name, DestPos):
		# DestPos is the destination postion as an excel range object
		self.CurrentTemplateName = template_name
		self.Dest_Row = DestPos.Row
		self.Dest_Col = DestPos.Column

		self.xlInst = Excel(DestPos)
		self.xlSheetDest = Sheet(self.xlInst)

		# Open the template, and test that the named template exists
		try:
			TemplateSheet = self._Open_Template(template_name)
			self._Copy_Paste_Template(template_name, DestPos)
		except Exception as ex:
			print ex
			#traceback.print_exc()
			print "There was a problem loading the template."

	def Set_Values(self, keymap):
		"""Set the values of the predfiend template keywords.
		Templates are only one-time writable (as the keyword is written over)"""
		SearchRange = shtRange(self.xlInst.Worksheet.Name, None, self.Dest_Row, self.Dest_Col, self.Dest_Row+self.Size_Rows, self.Dest_Col+self.Size_Cols)
		for keyword, value in keymap.iteritems():
			# Find all the cells with this placeholder keyword
			try:
				CellList = self.xlSheetDest.search(SearchRange, "key_"+keyword)
			except Exception as ex:
				print ex
				#traceback.print_exc()
			# Update all the cells with this placeholder keyword
			for cell in CellList:
				cell.Value = value

	def Auto_Fit(self):
		"""Autofits the template you just placed and handles the closing of the template file"""
		for column in range(int(self.Dest_Col), int(self.Dest_Col+self.Size_Cols+1)):
			self.xlSheetDest.autofit(self.xlInst.Worksheet.Name, column)
		self._Close_Template()

	def _Open_Template(self, template_name):
		self.xlTemplate = MSOffice.Excel.Launch.Excel(BookVisible=0, runninginstance=1, filename=self.TemplateFileName) # newinstance
		self.xlSheetSrc = Sheet(self.xlTemplate)
		return self.xlSheetSrc.getSheet(template_name)

	def _Close_Template(self):
		self.xlTemplate.closeWorkbook()

	def _Read_Dimentions(self, template_name):
		"""Every template has the the cell reference A1 = No. Rows, B1 = No. Cols"""
		num_rows = self.xlSheetSrc.getCell(template_name, 1, 1)
		num_cols = self.xlSheetSrc.getCell(template_name, 1, 2)
		return num_rows, num_cols

	def _Copy_Paste_Template(self, template_name, xldestination):
		"""Copy the template"""
		num_rows, num_cols = self._Read_Dimentions(template_name)
		self.Size_Rows = num_rows
		self.Size_Cols = num_cols
		TemplateRange = self.xlSheetSrc._getRange(template_name, 2, 1, 1+num_rows, num_cols)
		try:
			intialval = self.xlTemplate.xlApp.DisplayAlerts
			self.xlTemplate.xlApp.DisplayAlerts = False

			# Option 1
			#TemplateRange.Copy(Destination=xldestination) # Works only if both books are open in the same instance

			# Option 2
			TemplateRange.Copy()
			try:
				# Copy to a new instance and preserve formulae+formatting
				xldestination.Parent.PasteSpecial(Format="XML Spreadsheet", Link=False, DisplayAsIcon=False)
			except:
				# Copy within a instnace/workbbok and preserve formulae+formatting
				xldestination.Parent.Paste(Destination=xldestination)

			self.xlTemplate.xlApp.CutCopyMode = False # Dump clipbaord contents
			self.xlTemplate.xlApp.DisplayAlerts = intialval
		except Exception as ex:
			print ex
			print "Error copying from src"