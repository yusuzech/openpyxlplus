from numpy import number
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


def calc_value_shape(value,wrap_text=False,ndigits=2):
    """
    Get number of characters in longest line and number of lines

    If type(value) is numeric:
        round the number to given digits and then convert it text

    If wrap_text == True:
        split by "\n" to get number of lines
        n_horizontal: number of characters in longest line
        n_vertical: number of lines

    If wrap_text == False:
        n_horizontal: number of chracters
        n_vertical: 1

    Parameters:
    -----------

    value: any value, converted to text to get text length
    wrap_text: whether to count lines (split by "\\n")
    ndigits: number of digits to round number to.

    Return:
    -------

    (n_vertical,n_horizontal)
    """
    if isinstance(value,(number,int,float)):
        text = str(round(value,ndigits))
    else:
        text = str(value)

    if wrap_text:
        lines = text.split("\n")
        n_horizontal = max([len(x) for x in lines])
        n_vertical = len(lines)
    else:
        n_horizontal = len(text)
        n_vertical = 1
    return((n_vertical,n_horizontal))

def calc_value_size(
        value,
        wrap_text=False,
        min_width=1.43,
        min_height=13.2,
        max_width=255,
        max_height=409,
        width_factor=0.11,
        height_factor=1.35,
        ndigits=2,
        fontsize=11,
    ):
    """
    Based on text and width/height factor, calculate width and height.

    If wrap_text == False:
    width = fontsize * width factor * number of characters
    height = fontsize * height factor * 1

    If wrap_text = True and newline("\\n") present in cell text:
    width = fontsize * width factor * number of characters in longest line
    height = fontsize * height factor * number of lines

    Parameters:
    -----------

    wrap_text: wrap text
    min_width: width won't be below this number
    min_height: height won't be below this number
    max_width: width won't be above this number
    max_height: height won't be above this number
    width_factor: ideally number around default (found empirically)
    height_factor: ideally number around default (found empirically)
    max_ndigits: For numerical data, use rounded number to ndigits
    ndigits: number of digits (after decimal points) used for calculating width
    fontsize: fontsize of characters in cell

    Return:
    -------
    
    (height,width)
    """
    n_verical, n_horizontal = calc_value_shape(value,wrap_text=wrap_text,ndigits=ndigits)
    
    # height
    height = n_verical * height_factor * fontsize
    if height <= min_height:
        height = min_height
    elif height > min_height and height <= max_height:
        pass
    elif height > max_height:
        height = max_height
    else:
        raise Exception(f"wrong value for height: {height}")

    # width
    width = n_horizontal * width_factor * fontsize
    if width <= min_width:
        width = min_width
    elif width > min_width and width <= max_width:
        pass
    elif width > max_width:
        width = max_width
    else:
        raise Exception(f"wrong value for width: {width}")
    
    return((height,width))

    

# codes below will be removed upon release. #
# import openpyxl.utils.cell as converter
# from numpy import array as Array
# from pandas import Series

# def auto_replicate(value,target):
#     """
#     Automatically replicate a value or a list to target with same length

#     DROP WHEN RELEASING
#     """
#     length = len(target)
#     ArrayType = type(Array([1]))
#     # if value is not any of the list-like object
#     if (type(value) is not tuple) and (type(value) is not list) and\
#         not isinstance(value,ArrayType) and not isinstance(value,Series):
#         return([value] * length)
#     else:
#         remainder = length%len(value)
#         multiple = length//len(value)
#         if (remainder != 0) or (multiple <= 0):
#             raise ValueError((
#                 f"Length of target({length}) is not multiple of value({len(value)})"
#             ))

#     if isinstance(value,ArrayType) or isinstance(value,Series):
#         l = value.tolist()
#     else:
#         l = value
#     return(l * (multiple))


# def coord_to_address(coordinate):
#     """
#     coordinate should be (row,column)

#     DROP WHEN RELEASING
#     """
#     rownum = coordinate[0]
#     column = converter.get_column_letter(coordinate[1])
#     return(column + str(rownum))
    
# def boundaries_to_range(start_row,start_col,end_row,end_col):
#     """
#     convert topleft,rightbottom coordinates to address e.g. X1:Y2

#     DROP WHEN RELEASING
#     """
#     top_left = coord_to_address((start_row,start_col))
#     bot_right = coord_to_address((end_row,end_col))
#     return(top_left + ":" + bot_right)

# def clear_range(worksheet,range):
#     """
#     DROP WHEN RELEASING
#     """
#     for row in worksheet[range]:
#         for cell in row:
#             cell.value = None 
#     return(True)

# def clear_boundaries(worksheet,start_row,start_col,end_row,end_col):
#     """
#     DROP WHEN RELEASING
#     """
#     ret = clear_range(
#         worksheet,
#         boundaries_to_range(start_row,start_col,end_row,end_col)
#     )
#     return(ret)