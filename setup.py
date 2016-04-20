from setuptools import setup, find_packages

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
		  install_requires=['numpy', 'pywin32', 'pyodbc', 'six', 'wheel', 'virtualenv'], #pypiwin32
		  dependency_links=[], 
		  zip_safe=False,
		  #packages=find_packages()
		  )