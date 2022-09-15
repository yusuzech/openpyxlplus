from setuptools import setup, find_packages

VERSION = '0.5.0' 
DESCRIPTION = 'Based on openpyxl package. Write, modify and style excel worksheet with ease.'
LONG_DESCRIPTION = (
    'Based on numpy package, treat cells in workbook as numpy array.'
    'Enables writing to worksheet, getting cell values and attibutes, '
    'modifying cell attributes and styles in a vectorized manner.'
)

# Setting up
setup(
        name="openpyxlplus", 
        version=VERSION,
        author="Yifu Yan",
        author_email="yifdev1994@hotmail.com",
        description=DESCRIPTION,
        long_description=LONG_DESCRIPTION,
        packages=find_packages(),
        keywords=['python', 'excel','spreadsheet','openpyxl'],
        install_requires = [
            "openpyxl",
            "numpy",
            "pandas"
        ],
        classifiers= [
            "Development Status :: 3 - Alpha",
            "Programming Language :: Python :: 3",
            "Operating System :: Microsoft :: Windows",
        ]
)