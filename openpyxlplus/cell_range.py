from openpyxl.worksheet.cell_range import CellRange
from openpyxlplus.utils import open_workbook
from openpyxl.styles import Border,Side
import numpy as np
from copy import copy

def getattr_copy(obj,name,*args):
    if len(args) > 0:
        return(copy(getattr(obj,name,args[0])))
    else:
        return(copy(getattr(obj,name)))

v_getattr = np.vectorize(getattr_copy,otypes="O")
v_setattr = np.vectorize(setattr,otypes="O")

class SheetCellRange(CellRange):
    """
    Inheritate from openpyxl.worksheet.cell_range.CellRange. This class allow 
        additional operations on a given cell range such as:
    - Convert SheetCellRange to Cells which is a wrapper on numpy.ndarray.
    - Get cell or cells with coordinate(s).

    """
    def __init__(self,ws, range_string=None, min_col=None, min_row=None,
                 max_col=None, max_row=None, title=None):
        self.ws = ws
        super().__init__(range_string=range_string, min_col=min_col, 
            min_row=min_row, max_col=max_col, max_row=max_row, title=title)
    
    @property
    def cells(self):
        """
        Return cells in range as Cells object. This needs to be a property 
            because operations like shrink,expand,shift could change cell range.
        """
        cells = []
        for row in self.rows:
            l = []
            for cell in row:
                l.append(self.ws.cell(cell[0],cell[1]))
            cells.append(l)
        return(Cells(cells))

    def get_cell(self,coordinate,relative=True):
        """
        Parameters:
        coordinate: coordinate of cell in sheet (row:int,column:int)
        relative: Default to True, uses coordinate relative to top left 
            coordinate of this cell range. Additionally, when relative = True, 
            the coordinate start from 0. When relative = True, uses absolute 
            coordinate in the sheet.

        Return:
        openpyxl.cell.cell.Cell object
        """
        if relative:
            ret = self.ws.cell(
                coordinate[0] + self.min_row,
                coordinate[1] + self.min_col
            )
        else:
            ret = self.ws.cell(coordinate[0],coordinate[1])
        return(ret)
    
    def get_cells(self,coordinates,relative = True):
        """
        subset Cells with coordinates stored in list or numpy.ndarray. Result
        will the same shape as shape of coordinates

        Parameters:
        coordinates: list/tuple or numpy.ndarray which contains coordinate with 
            format of (row:int,column:int)
        relative: whether the coordinates are relative to range. See details in
            .get_cells method

        Return:
        Cells object
        """
        try:
            coordinates[0][0] # test depth of coordinates
            ret = []
            for row in coordinates:
                l = []
                for cell in row:
                    l.append(self.get_cell(cell,relative))
                ret.append(l)
            ret = Cells(ret)
        except:
            ret = Cells([self.get_cell(x,relative) for x in coordinates])
        return(ret)

    def merge_cells(self):
        """
        Merge all cells in range
        """
        self.ws.merge_cells(range_string=self.coord)
        return(self)

    def unmerge_cells(self):
        """
        Unmerge all Cells in range
        """
        self.ws.unmerge_cells(range_string=self.coord)
        return(self)

    def write(self,data,keep_style=True):
        """
        Alias to self.set_value. Write data to this range.

        Parameters:
        data: scaler, list or numpy array. Note that the shape of data should 
            match the shape of range to ensure writing correctly.
        keep_style: whether to keep original style
        """
        self.cells.set_value(data,keep_style=keep_style)
        return(self)

    def add_border(self,side = None,left=True,right=True,top=True,bottom=True):
        """
        Add outside border to cells.

        Parameters:
        side: Side object. If None, use Side(style="thin")
        left: whether to add border to left
        right: whether to add border to right
        head: whether to add border to top
        tail: whether to add border to tail
        """
        self.cells.add_border(
            side=side,left=left,right=right,top=top,bottom=bottom
        )
        return(self)

    @property
    def cell_values(self):
        """
        Display a copy of values in range as numpy array.
        """
        return(self.cells.get_value())

    def autofit_width(self,max_width=None,multiplier=1.2):
        """
        Automatically fit column width in current worksheet by detecting number 
            of characters in each cell. 

        Parameters:
        max_width: maximum column width
        multiplier: increase or reduce column width globally in current range. 
            This value may be adjusted for better view.

        Return:
        self
        """
        max_length = ((max_width if max_width else 999) - 2)/multiplier

        for col_coord in self.cols:
            col_max_length = max_length
            lengths = []
            for cell_coord in col_coord:
                cell = self.ws.cell(*cell_coord)
                if cell.coordinate in self.ws.merged_cells: # not check merged cells
                        continue
                if cell.value:
                    lengths.append(len(str(cell.value)))
                else:
                    lengths.append(0)
            # set column width
            cell_max_length = max(lengths)
            length = col_max_length if cell_max_length > col_max_length \
                else cell_max_length
            self.ws.column_dimensions[cell.column_letter].width = \
                (length + 2) * multiplier
        return(self)

    def show_in_excel(self,temp_name="temp.xlsx"):
        """
        Show this range in excel application.

        Parameters:
        temp_name: temporary file name. Default to "temp.xlsx"
        """
        _ , workbook = open_workbook(self.ws.parent,temp_name=temp_name
            ,prompt=False)

        ws = workbook.Sheets(self.ws.title)
        ws.Select() # make this sheet active
        ws.Range(self.coord).Select()
        print("Remember to close the workbook manually.")
        # Wait before closing it
        # _ = input("Press enter to close Excel")
        # excel_app.Application.Quit()

    # method alias
    show = show_in_excel

class TableRange(SheetCellRange):
    """
    Create a CellRange that represents a table.

    Additional attributes:
    header: a Cellrange object for table header
    index: a Cellrange object for table index
    body: a Cellrange object for table body (not index or header)
    
    """
    def __init__(self,ws,range_string=None, min_col=None, min_row=None,
                max_col=None, max_row=None, title=None,n_index=0,n_header=1):
        """
        Parameters:
        n_index: integer, how many columns from left are index
        n_header: integer, how many rows from top are header 
        """
        super().__init__(ws=ws,range_string=range_string, min_col=min_col, 
            min_row=min_row, max_col=max_col, max_row=max_row, title=title)
        self.n_index = n_index
        self.n_header = n_header

        
    @property
    def header(self):
        if self.n_header == 0:
            return(None)
        else:
            ret = SheetCellRange(
                self.ws,
                min_col = self.min_col + self.n_index,
                min_row = self.min_row,
                max_col = self.max_col,
                max_row = self.min_row + self.n_header-1,
            )
            return(ret)
    @property
    def index(self):
        if self.n_index == 0:
            return(None)
        else:
            ret = SheetCellRange(
                self.ws,
                min_col = self.min_col,
                min_row = self.min_row + self.n_header,
                max_col = self.min_col + self.n_index - 1,
                max_row = self.max_row,
            )
            return(ret)

    @property
    def body(self):
        ret = SheetCellRange(
            self.ws,
            min_col = self.min_col + self.n_index,
            min_row = self.min_row + self.n_header,
            max_col = self.max_col,
            max_row = self.max_row,
        )
        return(ret)

    @property
    def top_left_corner(self):
        """
        When index and header are both enabled, it creates a blank area at top 
        left corner of the table.

        Note: openpyxl.utils.dataframe.dataframe_to_rows adds index name when
        index = True. This may cause confusion sometimes, so it is advised to
        set index = False.
        """
        if self.n_index == 0 or self.n_header == 0:
            return(None)
        else:
            ret = SheetCellRange(
                self.ws,
                min_col = self.min_col,
                min_row = self.min_row,
                max_col = self.min_col + self.n_index - 1,
                max_row = self.min_row + self.n_header - 1,
            )
            return(ret)


class Cells(np.ndarray):
    """
    Wrapper class on numpy.ndarray. Allows following operations:

    - Read values, attributes from the range or subset of Cells
    - Apply style to range or subset of range
    - Write or change values in range
    - Add outside border to the range
    - merge/unmerge cells in range

    returned values preserve the shape of array.
    """
    def __new__(cls,array):
        array = np.array(array).view(cls)
        return(array)


    @staticmethod
    def from_range(ws, range_string=None, min_col=None, min_row=None,
        max_col=None, max_row=None, title=None):
        """
        Construct CellRange from range string or min/max col/row.
        """
        ret = SheetCellRange(ws, range_string=range_string, min_col=min_col, 
            min_row=min_row, max_col=max_col, max_row=max_row, title=title).cells
        return(ret)
    
    def to_range(self):
        """
        Convert current cells to SheetCellRange with top left cell in Cells as
        first cell and bottom right cell as last cell

        Return:
        SheetCellRange object
        """
        top_left = self[0,0]
        bottom_right = self[-1,-1]
        ret = SheetCellRange(
            ws=top_left.parent,
            range_string=f"{top_left.coordinate}:{bottom_right.coordinate}"
        )
        return(ret)

    def get_style(self,style_name):
        """
        Get a copy of cells's style as numpy.ndarray
        """
        ret = v_getattr(self,style_name).view(np.ndarray).copy()
        return(ret)

    def get_style_detail(self,style_name,detail):
        """
        Get detail in style

        Parameters:
        style_name: name of cell style
        details: attribute name of the detail, if detail is nested, provide a 
            list of attribute names. 
            e.g. object.get_style_detail("Border",["left","style"]) 
        """
        ret = self.get_style(style_name)
        if type(detail) == str:
            detail = [detail]
        for x in detail:
            ret = v_getattr(ret,x)
        ret = ret.copy()
        return(ret)

    def set_style(self,style_name,value):
        """
        Set cells' style to value(s), overwrites original style. returns self
        
        Parameters:
        style_name: style name to change
        value: desired style. Note that the shape of data should 
            match the shape of range to ensure writing correctly.
        """
        v_setattr(self,style_name,value)
        return(self)

    def modify_style(self,style_name,style):
        """
        Modify cells' style with given style(s). This method adds style(s) to 
        original one instead of completely overwriting it. If original style 
        already exist, it will ignore that style.
        """
        v_setattr(self,style_name,v_getattr(self,style_name)+style)
        return(self)


    def get_value(self):
        """
        Get a copy of cells' values
        """
        ret = v_getattr(self,"value").view(np.ndarray).copy().astype('O')
        return(ret)

    def set_value(self,value,keep_style = True):
        """
        Set cells values with given value(s)

        Parameters:
        value: value to write to cells.Note that the shape of data should 
            match the shape of range to ensure writing correctly.
        keep_style: whether to preserve original style
        """
        if keep_style:
            original_style = self.get_style("_style")
            v_setattr(self,"value",value)
            self.set_style("_style",original_style)
        else:
            v_setattr(self,"value",value)
        return(self)

    def get_range(self):
        """
        Convert Cells to SheetCellRange using top left and bottom right cell 
            coordinates.
        """
        try:
            top_left = self[0][0]
            bottom_right = self[-1][-1]
        except Exception:
            top_left = self[0]
            bottom_right = self[-1]
        ws = top_left.parent
        ret = SheetCellRange(
            ws,f"{top_left.coordinate}:{bottom_right.coordinate}"
        )
        return(ret)

    def side(self,direction,n):
        """
        Get n columns or rows from given direction to center.
        """
        if direction == "head":
            ret = self[:n,:]
        elif direction == "tail":
            ret = self[-n:,:]
        elif direction == "left":
            ret = self[:,:n]
        elif direction == "right":
            ret = self[:,-n:]
        else:
            raise ValueError(f"direcion:'{direction}'' is not " +\
            "one of 'head', 'tail', 'left' or 'right'")
        return(ret)
    
    def head(self,n=1):
        return(self.side("head",n))

    # method alias
    top = head

    def tail(self,n=1):
        return(self.side("tail",n))

    # method alias
    bottom = tail

    def left(self,n=1):
        return(self.side("left",n))
        
    def right(self,n=1):
        return(self.side("right",n))

    def add_border(self,side = None,left=True,right=True,top=True,bottom=True):
        """
        Add outside border to cells.

        Parameters:
        side: Side object. If None, use Side(style="thin")
        left: whether to add border to left
        right: whether to add border to right
        head: whether to add border to top
        tail: whether to add border to tail
        """
        if side is None:
            side = Side(style="thin")
        if left:
            self.left().modify_style("border",Border(left=side))
        if right:
            self.right().modify_style("border",Border(right=side))
        if top:
            self.head().modify_style("border",Border(top=side))
        if bottom:
            self.tail().modify_style("border",Border(bottom=side))
        return(self)