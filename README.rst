This is a python wrapper for the Component Object Model (COM) interface in Microsoft Office applications.

Currently the following applications have modules:
> Excel
> Outlook (partialy complete)

This is designed to be an automation package capable of automating those tedious jobs nobody wants to do. It is great for repeat tasks or quick hacks to filter and present data.

Usage Examples:
----------------

For the Excel Module (Sheets)
--------------------------------
>>> import MSOffice
>>> xl = MSOffice.Launch.Excel(visible=True, newinstance=True)
>>> sht = MSOffice.Worksheets.Sheet(xl)
>>> sht.addWorkbook()
>>> sht.renameSheet("my new sheet name") # Rename the active sheet i.e. "Sheet1"
>>> xl.save(r"C:\Temp\my Test File.xlsx")

For the Excel Module (Charts)
--------------------------------
>>> import MSOffice
>>> xl = MSOffice.Launch.Excel(visible=True, newinstance=True)
>>> sht = MSOffice.Worksheets.Sheet(xl)
>>> sht.addWorkbook() # Adds the default sheet called "Sheet1"
>>> from MSOffice.Excel.Charts import XlGraphs
>>> Graphs = XlGraphs(xl, sht)
# There are some x values in the Excel range "A2:A10"
>>> x_range = "='%s'!%s%d:%s%d" % ("Sheet1", "A", 2, "A", 10)
>>> Graphs.Create_Chart("My new chart", x_range)
# To add a chart as shape inside an existing sheet, add the paramtersheetname
# to Create_Chart, for example, sheetname="Name of existing sheet"
>>> Graphs.Add_Series("My new chart", "B2:B10", serieslabels=True) # Series (y) values in column B
