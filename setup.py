from setuptools import setup, find_packages

VERSION = '0.5.0' 
DESCRIPTION = 'Manipulate Excel file more easily'
LONG_DESCRIPTION = 'Manipulate Excel file more easily'

# Setting up
setup(
        name="openpyxlplus", 
        version=VERSION,
        author="Yifu Yan",
        author_email="yifdev1994@hotmail.com",
        description=DESCRIPTION,
        long_description=LONG_DESCRIPTION,
        packages=find_packages(),
        install_requires=[], # add any additional packages that 
        # needs to be installed along with your package. Eg: 'caer'
        
        keywords=['python', 'excel','spreadsheet','openpyxl'],
        classifiers= [
            "Development Status :: 3 - Alpha",
            "Programming Language :: Python :: 3",
            "Operating System :: Microsoft :: Windows",
        ]
)