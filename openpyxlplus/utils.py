def open_workbook(wb,temp_name="temp.xlsx",prompt=True):
    """
    Open the workbook object, only works in windows. It works by saving the 
        work book and open it using win32.com api. Excel file is closed and
        deleted by pressing enter.
    """
    from sys import platform
    if platform != "win32":
        raise ValueError("Operating System must be windows.")
    import win32com.client as win32
    from datetime import datetime
    import os
    wb.save(temp_name)
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True
    # open workbook
    wb=excel.Workbooks.Open(os.path.abspath(temp_name))

    if prompt:
        # Wait before closing it
        _ = input("Press enter to close Excel")
        excel.Application.Quit()
    else:
        # return excel object and workbook for futher operations
        return(excel,wb)

# codes below will be removed upon release. #
import openpyxl.utils.cell as converter
from numpy import array as Array
from pandas import Series

def auto_replicate(value,target):
    """
    Automatically replicate a value or a list to target with same length

    DROP WHEN RELEASING
    """
    length = len(target)
    ArrayType = type(Array([1]))
    # if value is not any of the list-like object
    if (type(value) is not tuple) and (type(value) is not list) and\
        not isinstance(value,ArrayType) and not isinstance(value,Series):
        return([value] * length)
    else:
        remainder = length%len(value)
        multiple = length//len(value)
        if (remainder != 0) or (multiple <= 0):
            raise ValueError((
                f"Length of target({length}) is not multiple of value({len(value)})"
            ))

    if isinstance(value,ArrayType) or isinstance(value,Series):
        l = value.tolist()
    else:
        l = value
    return(l * (multiple))


def coord_to_address(coordinate):
    """
    coordinate should be (row,column)

    DROP WHEN RELEASING
    """
    rownum = coordinate[0]
    column = converter.get_column_letter(coordinate[1])
    return(column + str(rownum))
    
def boundaries_to_range(start_row,start_col,end_row,end_col):
    """
    convert topleft,rightbottom coordinates to address e.g. X1:Y2

    DROP WHEN RELEASING
    """
    top_left = coord_to_address((start_row,start_col))
    bot_right = coord_to_address((end_row,end_col))
    return(top_left + ":" + bot_right)

def clear_range(worksheet,range):
    """
    DROP WHEN RELEASING
    """
    for row in worksheet[range]:
        for cell in row:
            cell.value = None 
    return(True)

def clear_boundaries(worksheet,start_row,start_col,end_row,end_col):
    """
    DROP WHEN RELEASING
    """
    ret = clear_range(
        worksheet,
        boundaries_to_range(start_row,start_col,end_row,end_col)
    )
    return(ret)