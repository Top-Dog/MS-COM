from setuptools import setup, find_packages
from cStringIO import StringIO
import sys

class Capturing(list):
	"""Capture print output to std out"""
	def __enter__(self):
		self._stdout = sys.stdout
		sys.stdout = self._stringio = StringIO()
		return self
	def __exit__(self, *args):
		self.extend(self._stringio.getvalue().splitlines())
		del self._stringio    # free up some memory
		sys.stdout = self._stdout

def readme():
	with open('README.rst') as f:
		return f.read()

if __name__ == "__main__":
	setup(name='MS Office pyCOM',
		  version='0.1',
		  description='A python wrapper for MS Excel VBA WinCOM32 API',
		  long_description=readme(),
		  url='https://github.com/Top-Dog/MS-COM',
		  author="Sean D. O'Connor",
		  author_email='sdo51@uclive.ac.nz',
		  license='MIT',
		  packages=['MSOffice', 'MSOffice.Excel', 'MSOffice.Excel.Charts', 'MSOffice.Excel.Worksheets', 'MSOffice.Excel.PivotTables'],
		  install_requires=['numpy', 'pypiwin32', 'pyodbc', 'six', 'wheel', 'virtualenv'], #pywin32
		  dependency_links=[], 
		  zip_safe=False,
		  #packages=find_packages()
		  )
	from win32com.client import makepy
	# For Excel (fill the cache so we can use constatns with late binding)
	sys.argv = ["makepy", r"C:\Program Files (x86)\Microsoft Office\Office14\Excel.exe"]
	with Capturing() as printoutput:
		makepy.main()
	if len(printoutput):
		if printoutput[0].startswith("Could not locate a type library matching"):
			sys.argv = ["makepy", r"C:\Program Files (x86)\Microsoft Office\Office16\Excel.exe"]
			with Capturing() as printoutput:
				makepy.main()
		if len(printoutput):
			if printoutput[0].startswith("Could not locate a type library matching"):
				sys.argv= [""]
				print "Choose: Microsoft Excel 14.0 Object Library (1.7), or Excel 16.0 Object Library (1.9)"
				makepy.main()
