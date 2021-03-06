This is a python wrapper for the Component Object Model (COM) interface in Microsoft Office applications.

Currently the following applications have modules:
> Excel
> Outlook (partialy complete)

This is designed to be an automation package capable of automating those tedious jobs nobody wants to do. It is great for repeat tasks or quick hacks to filter and present data.

**Usage Examples:**
----------------
*For the Excel Module (Sheets)*
--------------------------------
>>> import MSOffice
>>> xl = MSOffice.Launch.Excel(visible=True, newinstance=True)
>>> sht = MSOffice.Worksheets.Sheet(xl)
>>> sht.addWorkbook()
>>> sht.renameSheet("my new sheet name") # Rename the active sheet i.e. "Sheet1"
>>> xl.save(r"C:\Temp\my Test File.xlsx")
>>>
>>> # TearDown
>>> xl.closeWorkbook()
>>> xl.closeApplication()
>>> del xl

*For the Excel Module (Charts)*
--------------------------------
>>> # StartUp
>>> import MSOffice
>>> xl = MSOffice.Launch.Excel(visible=True, newinstance=True) # existinginstance=True
>>> sht = MSOffice.Worksheets.Sheet(xl)
>>> sht.addWorkbook() # Adds the default sheet called "Sheet1"
>>> 
>>> # Add data, so we can see what we've just plotted
>>> # If data is added after the chart is created you 
>>> # will find that Excel defaults to expecting the first
>>> # row to be a heading rather than data
>>> len_data = 10
>>> start_row = 2
>>> start_col = 1
>>> Table = [[x, x**2] for x in range(len_data+1)]
>>> sht.setRange("Sheet1", start_row, start_col, Table)
>>> 
>>> # Add a graph as a new tab (sheet) in the workbook
>>> chartname = "My new chart <name>"
>>> x_range = "='%s'!%s%d:%s%d" % ("Sheet1", "A", start_row, "A", start_row+len_data)
>>> y_range = "='%s'!%s%d:%s%d" % ("Sheet1", "B", start_row, "B", start_row+len_data)
>>> Graphs = XlGraphs(xl, sht)
>>> Graphs.Create_Chart(chartname, x_range)
>>> Graphs.Add_Series(chartname, y_range, serieslabels=False)
>>> # To add a chart as shape inside an existing sheet, add the paramter 'sheetname'
>>> # to Create_Chart, for example, sheetname="Sheet1"
>>>    
>>> # Wait, and have a look at what you've done
>>> raw_input("Done!")
>>> 
>>> # TearDown
>>> xl.closeWorkbook()
>>> xl.closeApplication()
>>> del xl

*For the Excel Module (PivotTables)*
--------------------------------
>>> # StartUp
>>> import MSOffice
>>> from MSOffice.Excel.PivotTables import PivotTables
>>> from MSOffice.Excel.Worksheets.Worksheet import shtRange 
>>> xl = MSOffice.Launch.Excel() 
>>> sht = MSOffice.Worksheets.Sheet(xl)
>>> sht.addWorkbook() # Adds the default sheet called "Sheet1"
>>> sht.addSheet("PIVOT") # Adds additional sheet called 'PIVOT' to store pivot table on
>>> len_data = 10
>>> start_row = 2
>>> start_col = 1
>>>
>>> Table = [[x, x**2] for x in range(len_data+1)]
>>> sht.setRange("Sheet1", start_row, start_col, Table)
>>> sht.setCell("Sheet1", start_row-1, start_col, "X")  # Pivot Tables require data headings
>>> sht.setCell("Sheet1", start_row-1, start_col+1, "Y")
>>>
>>> pt = PivotTables(xl, sht,"PIVOTTABLENAME","PIVOT") # Creates a pivot table called 'PIVOTTABLENAME' on sheet 'PIVOT'
>>> maxrow = sht.getMaxRow("Sheet1", 1, 1)
>>> maxcol = sht.getMaxCol("Sheet1",1,1)
>>> srcrange = shtRange("Sheet1", None, 1, 1, maxrow, maxcol) # Data range to make pivot table'
>>> destrange = shtRange("PIVOT", None, 1, 1, 1, 1) # Data range to put pivot table
>>> pt.createpivot(srcrange, destrange)
>>> pt.addpivotfields("Row","X","All")
>>> pt.addpivotfields("Column","Y","All")
>>> pt.addpivotfields("Values","Y","All")
>>> # Wait, and have a look at what you've done
>>> raw_input("Done!")
>>> 
>>> # TearDown
>>> xl.closeWorkbook()
>>> xl.closeApplication()
>>> del xl
