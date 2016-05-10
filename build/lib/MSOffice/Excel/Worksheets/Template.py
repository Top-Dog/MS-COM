class Template(object):
	"""Copy tables from a existing source document an place them
	in a working document."""
	def __init__(self, TemplateFileName):
		#self.template_dir = template_dir
		self.TemplateFileName = TemplateFileName
		self.Dest_Row = 0
		self.Dest_Col = 0

	def Place_Template(self, template_name, DestPos):
		# DestPos is the destination postion as an excel range object
		self.Dest_Row = DestPos.Row
		self.Dest_Col = DestPos.Column

		# Open the template, and test that the named template exists
		try:
			TemplateSheet = self._Open_Template(template_name)
			self._Copy_Paste_Template(template_name, DestPos)
			self._Close_Template()
		except:
			print "Template %s does not exist" % template_name

	def Set_Values(self, template_name, **kwargs):
		num_rows, num_cols = self._Read_Dimentions(template_name)
		SearchRange = shtRange(template_name, None, self.Dest_Row, self.Dest_Col, self.Dest_Row+num_rows, self.Dest_Col+num_cols)
		for keyword, value in kwargs.iteritems():
			# Find all the cells with this placeholder keyword
			CellList = self.xlSheet.search(SearchRange, "key_"+keyword)
			# Update all the cells with this placeholder keyword
			for cell in CellList:
				cell.Value = value

	def _Open_Template(self, template_name):
		self.xlTemplate = MSOffice.Excel.Launch.Excel(visible=0, runninginstance=1, filename=self.TemplateFileName)
		self.xlSheet = MSOffice.Excel.Worksheets.Worksheet(xlTemplate)
		return self.xlSheet.getSheet(template_name)

	def _Close_Template(self):
		self.xlTemplate.closeWorkbook()

	def _Read_Dimentions(self, template_name):
		"""Every template has the the cell reference A1 = No. Rows, B1 = No. Cols"""
		num_rows = self.xlSheet.getCell(template_name, 1, 1)
		num_cols = self.xlSheet.getCell(template_name, 1, 2)
		return num_rows, num_cols

	def _Copy_Paste_Template(self, template_name, xldestination):
		"""Copy the template"""
		num_rows, num_cols = self._Read_Dimentions(template_name)
		TemplateRange = self.xlTemplate._getRange(template_name, 2, 1, 1+num_rows, num_cols)
		TemplateRange.Copy(Destination=xldestination)