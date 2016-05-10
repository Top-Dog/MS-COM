import MSOffice
from Worksheet import shtRange

class Template(object):
	"""Copy tables from a existing source document an place them
	in a working document."""
	def __init__(self, xlSheetDest, TemplateFileName):
		#self.template_dir = template_dir
		self.xlSheetDest = xlSheetDest
		self.TemplateFileName = TemplateFileName
		self.Dest_Row = 0
		self.Dest_Col = 0
		self.Size_Rows = 0
		self.Size_Cols = 0

	def Place_Template(self, template_name, DestPos):
		self.CurrentTemplateName = template_name
		# DestPos is the destination postion as an excel range object
		self.Dest_Row = DestPos.Row
		self.Dest_Col = DestPos.Column
		self.Sheet_Obj = DestPos.Worksheet

		# Open the template, and test that the named template exists
		try:
			TemplateSheet = self._Open_Template(template_name)
			self._Copy_Paste_Template(template_name, DestPos)
		except:
			print "There was a problem loading the template."

	def Set_Values(self, keymap):
		"""Set the values of the predfiend template keywords.
		Templates are only one-time writable (as the keyword is written over)"""
		SearchRange = shtRange(self.Sheet_Obj.Name, None, self.Dest_Row, self.Dest_Col, self.Dest_Row+self.Size_Rows, self.Dest_Col+self.Size_Cols)
		for keyword, value in keymap.iteritems():
			# Find all the cells with this placeholder keyword
			CellList = self.xlSheetDest.search(SearchRange, "key_"+keyword)
			# Update all the cells with this placeholder keyword
			for cell in CellList:
				cell.Value = value

	def Auto_Fit(self):
		"""Autofits the template you just placed and handles the closing of the template file"""
		for column in range(int(self.Dest_Col), int(self.Dest_Col+self.Size_Cols+1)):
			self.xlSheetDest.autofit(self.Sheet_Obj.Name, column)
		self._Close_Template()

	def _Open_Template(self, template_name):
		self.xlTemplate = MSOffice.Excel.Launch.Excel(BookVisible=0, runninginstance=1, filename=self.TemplateFileName)
		self.xlSheetSrc = MSOffice.Excel.Worksheets.Worksheet.Sheet(self.xlTemplate)
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
		TemplateRange.Copy(Destination=xldestination)