from setuptools import setup, find_packages
import sys
from win32com.client import makepy

def readme():
    with open('README.rst') as f:
        return f.read()

if __name__ == "__main__":
	setup(name='MS Office pyCOM',
		  version='0.1',
		  description='A python wrapper for MS Excel VBA WinCOM32 API',
		  long_description=readme(),
		  url='some github url (update this!)',
		  author="Sean D. O'Connor",
		  author_email='sdo51@uclive.ac.nz',
		  license='MIT',
		  packages=['MSOffice', 'MSOffice.Excel.Charts', 'MSOffice.Excel.Worksheets'],
		  install_requires=['numpy', 'pypiwin32', 'pyodbc', 'six', 'wheel', 'virtualenv'], #pywin32
		  dependency_links=[], 
		  zip_safe=False,
		  #packages=find_packages()
		  )
	
	# For Excel (fill the cache so we can use constatns with late binding)
	sys.argv = ["makepy", r"C:\Program Files (x86)\Microsoft Office\Office14\Excel.exe"]
	makepy.main()
	userin = raw_input("Did the defintions successfuly build? You should see 'Importing module' as the last line in the console if it did. Type y/n and press enter. ")
	if userin.lower() in ("no", "n"):
		sys.argv= [""]
		print "Choose: Microsoft Excel 14.0 Object Library (1.7)"
		makepy.main()
