'''
Created on 30/3/2016

@author: Sean D. O'Connor
'''

from ..Launch import c
import os

class XlGraphs(object):
	"""A class to handle the production of the required 
	graphs for feeder/substation analysis. Uses range
	objects as opposed to value arrays for data storage
	in memmory.
	
	Features:
	- Scatter Plot: Feeder Current
	- Line Graph: Feeder kVA(s)
	"""
	def __init__(self, xlInstance, Sheet):
		#super(XlGraphs, self).__init__(xlInstance)
		#super(XlGraphs, self).__init__(xlInstance, FeederIDs)
		self.ChartObjects = [] # A list of lists: [[charObject, xValues], ..., {args to setup chart layout}]
		self.xlInst = xlInstance
		self.Sheet = Sheet
	
	def Create_Chart(self, chartName, xRange, **kwargs):
		"""Create a new (blank) chart object.
		The chartType can be an existing excel chart type (enum), or
		a filepath to an existing template (.crtx)."""
		sheetname = kwargs.get('sheetname', '')
		if sheetname:
			# Add a new shape object 
			shape = self.Sheet.getSheet(sheetname).Shapes.AddChart()
			chart = shape.Chart
			chart.SetSourceData(Source=self.xlInst.xlBook.Sheets(sheetname).Range(xRange)) # This will limit the number of pre-included series, which makes things MUCH faster
			for series in chart.SeriesCollection():
				series.Delete()
			shape.Name = chartName
			#chart.Name = chartName -- can't do this
			chart.ChartType = kwargs.get('chartType', c.xlLine)
		else:
			# Add a new worksheet object
			chart = self.Sheet.addChart(chartName, kwargs.get('chartType', c.xlLine)) # kwargs.get('xlLine', c.xlLine) - sets up a chart like a new sheet
			chart.ChartArea.Clear() # Clears any existing series from the chart
			#chart.SeriesCollection().NewSeries()
			shape = None
			
		# Set the default chart configuration
		chart.SeriesCollection().NewSeries() # need this here, or SetElement won't work
		chart.HasLegend = False
		# Set the title
		chart.HasTitle = True
		chart.SetElement(1) # can't access the 'mso' consts msoElementChartTitleCenteredOverlay
		chart.ChartTitle.Text = chartName
		
		self.Chart_Layout(chart, kwargs)
		
#        chart.HasTitle = True
#        chart.SetElement(1) # can't access the 'mso' consts msoElementChartTitleCenteredOverlay
#        chart.ChartTitle.Text = chartName
#        
#        # Optional: Add axes labels
#        if kwargs.get('xlabel'):
#            chart.Axes(c.xlCategory, c.xlPrimary).HasTitle = True
#            chart.Axes(c.xlCategory, c.xlPrimary).AxisTitle.Text = kwargs.get('xlabel')
#        if kwargs.get('ylabel'):
#            chart.Axes(c.xlValue, c.xlPrimary).HasTitle = True
#            chart.Axes(c.xlValue, c.xlPrimary).AxisTitle.Text = kwargs.get('ylabel')
#        
#        # Optional: Add scaling limits, steps to y (category) axes
#        if kwargs.get('ymin'):
#            chart.Axes(c.xlValue).MinimumScale = kwargs.get('ymin')
#        if kwargs.get('ymax'):
#            chart.Axes(c.xlValue).MaximumScale = kwargs.get('ymax')
#        if kwargs.get('majorunit'):
#            chart.Axes(c.xlValue).MajorUnit = kwargs.get('majorunit')

		self.ChartObjects.append([(shape, chart), xRange, [], kwargs])
	
	def Add_Series(self, chartName, yRange, **kwargs):
		"""Adds a series to an existing chart object"""
		# Find the index of the chart we are working on
		# Only works on charts that were created using this module in the same instance
		chartIndex = self._Get_Chart_Index(chartName)
		
		chart = self.ChartObjects[chartIndex][0][1] # Get the subsclass of the shape.. the chart
		xRange = self.ChartObjects[chartIndex][1]
		self.ChartObjects[chartIndex][2].append(yRange)
		seriesIndex = len(self.ChartObjects[chartIndex][2])
		
		chart.SeriesCollection().Add(Source=yRange, Rowcol=c.xlColumns, 
			SeriesLabels=kwargs.get("serieslabels", False), CategoryLabels=kwargs.get("categorylabels", False)) # Assume the data is in columns and that the selction is only data (no column/row headings)
		chart.SeriesCollection(seriesIndex).XValues = xRange
		chart.DisplayBlanksAs = c.xlNotPlotted
		
		chart.SeriesCollection(seriesIndex).ChartType = kwargs.get("seriestype", chart.ChartType) # You can style individual series' on the chart (might break if a template is used to create the chart)
		#chart.Type = c.xlLine or c.xlAreaStacked
		
		if kwargs.get('seriesname', ''):
			chart.HasLegend = True
			chart.SeriesCollection(seriesIndex).Name = kwargs.get('seriesname', '<no name supplied>')
			
	def Apply_Template(self, chartName, filename, **kwargs):
		"""Applies an exsting template, stored as a file, to an existing chart object.
		This is useful if you want to specify the colour, line styles, legend etc. all in one hit"""
		chartIndex = self._Get_Chart_Index(chartName)
		chart = self.ChartObjects[chartIndex][0][1] # Get the subsclass of the shape.. the chart
		ChartName = chart.ChartTitle.Text
		if os.path.isfile(filename) and filename.endswith('.crtx'):
			chart.ApplyChartTemplate(filename)
			
		# Set the title
		chart.HasTitle = True
		chart.SetElement(1) # can't access the 'mso' consts msoElementChartTitleCenteredOverlay
		chart.ChartTitle.Text = chartName
		
		self.Chart_Layout(chart, kwargs)
		
	def _Get_Chart_Index(self, chartName):
		chartIndex = 0
		for chartStruct in self.ChartObjects: # [(shape, chart), xRange, []]
			shape = chartStruct[0][0]
			if shape is not None:
				shapeName = shape.Name
			else:
				shapeName = ""
			chart = chartStruct[0][1]
			if chart.Name == chartName or shapeName == chartName:
				break
			chartIndex += 1
		assert chartIndex < len(self.ChartObjects), "Could not find a chart matching the supplied chart name"
		return chartIndex

	def Chart_Exists(self, chartName):
		chartIndex = 0
		for chartStruct in self.ChartObjects: # [(shape, chart), xRange, []]
			shape = chartStruct[0][0]
			if shape is not None:
				shapeName = shape.Name
			else:
				shapeName = ""
			chart = chartStruct[0][1]
			if chart.Name == chartName or shapeName == chartName:
				return True
		return False

	def _Unpack_Chart(self, **kwargs):
		"""Not fully implemented, but will return only the requested parameters."""
		if kwargs.get("chartName"):
			chartIndex = self._Get_Chart_Index(kwargs.get("chartName"))
		elif kwargs.get("chartIndex"):
			chartIndex = kwargs.get("chartIndex")
		
		chart = self.ChartObjects[chartIndex][0][1] # Get the subsclass of the shape.. the chart
		xRange = self.ChartObjects[chartIndex][1]
		self.ChartObjects[chartIndex][2].append(yRange)
		seriesIndex = len(self.ChartObjects[chartIndex][2])
		
		return "something"
	
	def Chart_Layout(self, chart, kwargs):        
		# Optional: Add axes labels
		if kwargs.get('xlabel'):
			chart.Axes(c.xlCategory, c.xlPrimary).HasTitle = True
			chart.Axes(c.xlCategory, c.xlPrimary).AxisTitle.Text = kwargs.get('xlabel')
		if kwargs.get('ylabel'):
			chart.Axes(c.xlValue, c.xlPrimary).HasTitle = True
			chart.Axes(c.xlValue, c.xlPrimary).AxisTitle.Text = kwargs.get('ylabel')
		
		# Optional: Add scaling limits, steps to y (category) axes
		if kwargs.get('ymin'):
			chart.Axes(c.xlValue).MinimumScale = kwargs.get('ymin')
		if kwargs.get('ymax'):
			chart.Axes(c.xlValue).MaximumScale = kwargs.get('ymax')
		if kwargs.get('majorunit'):
			chart.Axes(c.xlValue).MajorUnit = kwargs.get('majorunit')

	def Set_Max_X_Value(self, chartname, max_x):
		"""Allows the user to set the maxium x value on a graph"""
		chartIndex = self._Get_Chart_Index(chartname)
		chart = self.ChartObjects[chartIndex][0][1] # Get the subsclass of the shape.. the chart

		#if type(max_x) is datetime.datetime:
		#	# Convert python datetime to something Excel understands
		#	max_x = self.xlInst.xlApp.WorksheetFunction.DATEVALUE(max_x.date.__str__())
		chart.Axes(c.xlCategory).MaximumScale = max_x # will work even if asigning a python datetime object
	
	# self, chartName, xRange, chartType, **kwargs
	def Create_Chart_Shape(self, sheetname, chartName, xRange, chartType, **kwargs):
		"""Create a shape (chart) embeded in an Excel sheet"""
		sht = self.Sheet.getSheet(sheetname)
		shape = sht.Shapes.AddChart()
		chart = shape.Chart
		
		#chart.ChartArea.Clear() # Clears any existing series - also removes the chart
		for series in chart.SeriesCollection():
			series.Delete()
		
		chart.SeriesCollection().NewSeries()
		chart.HasLegend = False
		chart.HasTitle = True
		chart.SetElement(1) # can't access the 'mso' consts msoElementChartTitleCenteredOverlay
		chart.ChartTitle.Text = chartName
		
		self.ChartObjects.append([shape, xRange, []]) # Note that we use the shape object here, not the chart

		chart.ChartType = chartType #c.xlLine #c.xlXYScatter
		chart.SeriesCollection().Add(Source=src)
		
		# The whole series
		series = chart.SeriesCollection(1)
		chart.HasLegend = False
		series.Name = titles.title
		
		
	def Arrange_Shapes(self, sheetname, widths, heights, offset, chartOffset=0):
		'''Arranges the charts in block grid fashion'''
		sht = self.Sheet.getSheet(sheetname)
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
			
			
			# So the <chart object>.Name might return something like "Calculation Chart 3", where Calculation is the sheetname
			# ChartObjects expects a index number of chart name, so only "Chart 3" will work as a reference, not "Calculation Chart 3"
			# The name to use is actually the 'Shape' name instead of the shape.Chart.Name
			chart = sht.ChartObjects(chartIndex) # Must gain acess to the chart object like this
			chart.Width = width
			chart.Height = height
			chart.Left = ((chartIndex - start) % numWide) * widthPrevious + offset.x # normally: (chartIndex - 1)....
			chart.Top = int((chartIndex - start) / numWide) * heightPrevious + offset.y
			
	def Set_Dimentions(self, chartName, width, height):
		"""Set the width and height of a shape object"""
		chartIndex = self._Get_Chart_Index(chartName)
		
		shape = self.ChartObjects[chartIndex][0][0] # Get the subsclass of the shape.. the chart
		#chart = chart = self.ChartObjects[chartIndex][0][1] # Get the subsclass of the shape.. the chart
		#xRange = self.ChartObjects[chartIndex][1]
		#self.ChartObjects[chartIndex][2].append(yRange)
		#seriesIndex = len(self.ChartObjects[chartIndex][2])
		
		if shape is not None:
			shape.Width = width
			shape.Height = height
		
		
	def Set_Position(self, chartName, x, y):
		"""Set the x and y postion of the top left corner of a shape object in the current sheet.
		Origin is the top left corner of the sheet."""
		chartIndex = self._Get_Chart_Index(chartName)
		shape = self.ChartObjects[chartIndex][0][0] # Get the subsclass of the shape.. the chart
		chart = self.ChartObjects[chartIndex][0][1] # Get the chart object itself
		#if chart is not None:
		#	chart.Location(c.xlLocationAsNewSheet, "sheet name")
		if shape is not None:
			shape.Left = x
			shape.Top = y

	def Set_2nd_Position(self, sheetname, chartName, x, y):
		"""Set the x and y postion of the top left corner of a shape object in any sheet.
		Origin is the top left corner of the sheet."""
		chartIndex = self._Get_Chart_Index(chartName)
		shape = self.ChartObjects[chartIndex][0][0] # Get the subsclass of the shape.. the chart
		if shape is not None:
			shape.Left = x
			shape.Top = y
		
	def Arrange_Shapes(self, shapes, XDim, YDim):
		"""@param shapes: A list of shapes to potion in a 2D grid format.
		@param XDim: The number shapes to place horizontally.
		@param YDim: The number of shapes to place vertically.
		Shapes are arranged left to right, top to bottom."""        
		#for i in range(1, len(shapes) + 1):
		TotalWidth = 0
		TotalHeight = 0
		for shape in shapes:
			#shape = sht.ChartObjects(chartIndex)
			shape.Width = width
			shape.Height = height
			shape.Left = ((chartIndex - start) % numWide) * widthPrevious + offset.x # normally: (chartIndex - 1)....
			shape.Top = int((chartIndex - start) / numWide) * heightPrevious + offset.y
		
			if i % 2 or start > 1: 
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
				chartPrevious = sht.ChartObjects(i - 1)
				widthPrevious = chartPrevious.Width
				heightPrevious = chartPrevious.height
			except:
				pass
			
	# Deprocated function
	def Add_Templated_Series(self, chartName, *args, **kwargs):
		"""args is a list of y values (in order) to be applied to existing series'"""
		chartIndex = 0
		for chartStruct in self.ChartObjects: # [chart, xRange, []]
			if chartStruct[0].Name == chartName:
				break
			chartIndex += 1
		
		# Read in chart paramters
		chart = self.ChartObjects[chartIndex][0]
		xRange = self.ChartObjects[chartIndex][1]
		self.ChartObjects[chartIndex][2].append(yRange)
		seriesIndex = len(self.ChartObjects[chartIndex][2])
