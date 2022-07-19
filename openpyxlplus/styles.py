"""
This module works but is not easy to use. Please use cell_range module which 
supports all functionality in this module. 
"""
from openpyxl.utils.dataframe import dataframe_to_rows
import openpyxl.utils.cell as converter
import openpyxlplus.utils
from numpy import array as Array
from pandas import Series
from pandas import DataFrame
from openpyxl.styles import Border,Side,NamedStyle


def range_boundaries_rc(range_string):
    """returns [min_row,min_col,max_row,max_col]"""
    boundaries = converter.range_boundaries(range_string)
    ret = (boundaries[1],boundaries[0],boundaries[3],boundaries[2])
    return(ret)

def get_sides(border):
    """
    Saves all sides(left,right,top,bottom) of a Border object to a list of 
    dictionaries. Returns [] if the side style is not specified
    for any sides.
    
    Returns:
    [
        {"direction":"left","side":SideObject},
        {"direction":"right","side":SideObject},
        ...
    ]
    
    """
    sides = []
    for side_string in ["left","right","top","bottom"]:
        side = getattr(border,side_string)
        if side:
            if side.style is not None:
                sides.append({"direction":side_string,"side":side})
    return(sides)

def merge_border(old_border,new_border,new = True):
    """
    Merge two Border objects
    
    Parameters:
    new: if True, overwrite old with new and vice versa
    """
    old_sides = get_sides(old_border)
    new_sides = get_sides(new_border)
    final_sides = {}
    if new:
        sides = old_sides + new_sides
    else:
        sides = new_sides + old_sides
    for item in sides:
        final_sides.update({item["direction"]:item["side"]})
    return(Border(**final_sides))

def cell_append_border(cell,border,new=True):
    """
    Append border to a cell. Only add side specified by border
    
    Parameters:
    new: if True, overwite existing border with new border and vice versa
    """
    new_border = merge_border(cell.border,border,new=new)
    cell.border = new_border
    return(True)
    

def boundaries_append_border(
    worksheet,
    start_row,start_col,end_row,end_col,
    border,new=True
):
    """
    Append same border to all cells in range, can choose between over using new
    borders or keep original borders
    """
    range_address = openpyxlplus.utils.boundaries_to_range(
        start_row,start_col,end_row,end_col
    )
    for cell_row in converter.rows_from_range(range_address):
        for cell_address in cell_row:
            cell = worksheet[cell_address]
            cell_append_border(cell,border,new=new)
    return(True)

def range_append_border(worksheet,range_string,border,new = True):
    """"wraps around boundaries_append_border()"""
    boundaries = range_boundaries_rc(range_string)
    ret = boundaries_append_border(worksheet,*boundaries,border,new)
    return(ret)

def boundaries_apply_style(
    worksheet,
    start_row,start_col,end_row,end_col,
    style
):
    """
    apply named style to all cells in range,new style always overwrite original 
    style.

    style: a named style or a numpy array/list of list/ pandas dataframe with
        each individual element as a named style. The shape must be the same as
        range specified by range_string or boundaries
    """
    range_address = openpyxlplus.utils.boundaries_to_range(
        start_row,start_col,end_row,end_col
    )
    width = end_col - start_col + 1
    height = end_row - start_row + 1
    if isinstance(style,NamedStyle):
        style = Array([style] * (width * height)).reshape((height,width))
    elif type(style) == list:
        style = Array(style)
    elif isinstance(style,DataFrame):
        style = style.values
    elif isinstance(style,type(Array([1]))):
        pass
    else:
        raise ValueError(
            "style must be one of NamedStyle,list,numpy array or pandas DataFrame"
        )

    if style.shape != (height,width):
        raise ValueError((
            f"The shape of style{(height,width)} is must be equal to the size "
            f"of style({style.shape})"
        ))
    for row,cell_row in enumerate(converter.rows_from_range(range_address)):
        for col,cell_address in enumerate(cell_row):
            cell = worksheet[cell_address]
            cell.style = style[row,col]
    return(True)

def range_apply_style(worksheet,range_string,style):
    """"wraps around boundaries_apply_style()"""
    boundaries = range_boundaries_rc(range_string)
    ret = boundaries_apply_style(worksheet,*boundaries,style)
    return(ret) 

def boundaries_append_outline(worksheet,start_row,start_col,end_row,end_col,side,new=True):
    """
    append border around given range using provided side(Side Object).can choose 
        between over using new borders or keep original borders
    new: whether to overwrite original border with new one
    """
    if isinstance(side,Side) is not True:
        raise ValueError(f"side must be object of {Side}")
    
    arr_list = []
    range_address = openpyxlplus.utils.boundaries_to_range(
        start_row,start_col,end_row,end_col
    )
    for cell_row in converter.rows_from_range(range_address):
        arr_list.append(cell_row)

    arr = Array(arr_list)
    four_sides = {
        "left":arr[:,0].tolist(),
        "right":arr[:,-1].tolist(),
        "top":arr[0,:].tolist(),
        "bottom":arr[-1,:].tolist()
    }
    
    for one_side in ["left","right","top","bottom"]:
        cell_addresses = four_sides[one_side]
        for cell_address in cell_addresses:
            cell = worksheet[cell_address]
            cell_append_border(cell,Border(**{one_side:side}),new=new)
    return(True)

def range_append_outline(worksheet,range_string,side,new = True):
    """"wraps around boundaries_append_outline()"""
    boundaries = range_boundaries_rc(range_string)
    ret = boundaries_append_outline(worksheet,*boundaries,side,new)
    return(ret)     


def adjust_column_width(
    ws,
    cell_range= "A1:Y25",
    min_width = 13,
    max_width = 50,
    modifier = 1.2
):
    """
    adjust columns using string values in a range

    cell_range: a list: [start_row,start_col,end_row,end_col] or 
        a string:"A1:Y25"
    """
    if type(cell_range) == str:
        start_row, start_col, end_row, end_col = range_boundaries_rc(cell_range)
    elif (type(cell_range) == list) or (type(cell_range) == tuple):
        start_row, start_col, end_row, end_col = cell_range
    else:
        raise ValueError("cell_range must be string or list/tuple")

    for col in range(start_col,end_col+1):
        width = min_width
        for row in range(start_row,end_row+1):
            cell_value = ws.cell(row,col).value
            if cell_value:
                if type(cell_value) == str and len(cell_value) > min_width:
                    width = round(len(cell_value)*modifier)
        # use cell that contains string with most characters
        column_string = converter.get_column_letter(col)
        ws.column_dimensions[column_string].width = min(width,max_width)
    return(True)