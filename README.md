# openpyxlplus

Based on openpyxl package. This package stores cells as numpy array which enables easier writing, modifications and styling of worksheets.

1. Use `openpyxlplus.writer` to write scaler value, list, numpy array, pandas dataframe to worksheet. 
2. Use `openpyxlplus.cell_range.SheetCellRange`, `openpyxlplus.cell_range.SheetTableRange` and `openpyxlplus.cell_range.Cells` to write/get cell values and get/write/modify cell attribute easily.

# How to Use

## Writing to Worksheet


```python
# preparation
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxlplus import writer,cell_range
scalar = "s1"
l = ["l1","l2","l3"]
array = np.arange(9).reshape((3,3))
df = pd.DataFrame((np.arange(9)*100).reshape((3,3)),index=["r1","r2","r3"],columns=["c1","c2","c3"])

wb = Workbook()
ws = wb.active

# write scaler
rg1 = writer.write_value(scalar,ws,cell=ws["A1"])

# write list "up" or down (vertically)
rg2 = writer.write_list(l,ws,cell=ws["A3"],direction="down")

# write list "left" or "right" (horizontally)
rg3 = writer.write_list(l,ws,cell=ws.cell(1,3),direction="right")

# write array
rg4 = writer.write_array(array,ws,cell=ws.cell(3,3))

# write dataframe
rg5 = writer.write_dataframe(df,ws,cell=ws.cell(7,1),index=True,header=True)

# write with SheetCellRange.write method. Values are brodcasted
rg6 = cell_range.SheetCellRange(ws,range_string="G1:G5")
rg6.write("a")

# write with cells. Values should have the same shape as numpy array representing the cells
rg7 = cell_range.SheetCellRange(ws,range_string="G6:G8")
rg7.cells.set_value([["b"],["c"],["d"]])

rg_all = cell_range.SheetCellRange(ws,range_string="A1:G10")

display(pd.DataFrame(rg_all.cell_values))
# wb.save("test_workbook.xlsx")
```


<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>0</th>
      <th>1</th>
      <th>2</th>
      <th>3</th>
      <th>4</th>
      <th>5</th>
      <th>6</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>s1</td>
      <td>None</td>
      <td>l1</td>
      <td>l2</td>
      <td>l3</td>
      <td>None</td>
      <td>a</td>
    </tr>
    <tr>
      <th>1</th>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>a</td>
    </tr>
    <tr>
      <th>2</th>
      <td>l1</td>
      <td>None</td>
      <td>0</td>
      <td>1</td>
      <td>2</td>
      <td>None</td>
      <td>a</td>
    </tr>
    <tr>
      <th>3</th>
      <td>l2</td>
      <td>None</td>
      <td>3</td>
      <td>4</td>
      <td>5</td>
      <td>None</td>
      <td>a</td>
    </tr>
    <tr>
      <th>4</th>
      <td>l3</td>
      <td>None</td>
      <td>6</td>
      <td>7</td>
      <td>8</td>
      <td>None</td>
      <td>a</td>
    </tr>
    <tr>
      <th>5</th>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>b</td>
    </tr>
    <tr>
      <th>6</th>
      <td>None</td>
      <td>c1</td>
      <td>c2</td>
      <td>c3</td>
      <td>None</td>
      <td>None</td>
      <td>c</td>
    </tr>
    <tr>
      <th>7</th>
      <td>r1</td>
      <td>0</td>
      <td>100</td>
      <td>200</td>
      <td>None</td>
      <td>None</td>
      <td>d</td>
    </tr>
    <tr>
      <th>8</th>
      <td>r2</td>
      <td>300</td>
      <td>400</td>
      <td>500</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
    </tr>
    <tr>
      <th>9</th>
      <td>r3</td>
      <td>600</td>
      <td>700</td>
      <td>800</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
    </tr>
  </tbody>
</table>
</div>


## Getting Cell Values and Attributes

To get/set cell values, or get/modify cell attributes/styles, you can use `openpyxlplus.cell_range.Cells`.

To convert `SheetCellRange` to `Cells`, use `SheetCellRange.cells` attribute. To convert `Cells` back to `SheetCellRange`, use `Cells.to_range()` method


```python
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxlplus import writer,cell_range
wb = Workbook()
ws = wb.active
df = pd.DataFrame(np.arange(9).reshape((3,3)),index=["r1","r2","r3"],columns=["c1","c2","c3"])
rg = writer.write_dataframe(df,ws,cell=ws.cell(1,1),index=True,header=True)
# rg = cell_range.SheetCellRange(range_string="A1:D4")
display(rg)
display(rg.header) # .header only avilable to SheetTableRange
display(rg.index) # .index only available to SheetTableRange
display(rg.body) # .index only available to SheetTableRange
display(type(rg.cells))
display(pd.DataFrame(rg.cell_values))
```


    <SheetTableRange A1:D4>



    <SheetCellRange B1:D1>



    <SheetCellRange A2:A4>



    <SheetCellRange B2:D4>



    openpyxlplus.cell_range.Cells



<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>0</th>
      <th>1</th>
      <th>2</th>
      <th>3</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>None</td>
      <td>c1</td>
      <td>c2</td>
      <td>c3</td>
    </tr>
    <tr>
      <th>1</th>
      <td>r1</td>
      <td>0</td>
      <td>1</td>
      <td>2</td>
    </tr>
    <tr>
      <th>2</th>
      <td>r2</td>
      <td>3</td>
      <td>4</td>
      <td>5</td>
    </tr>
    <tr>
      <th>3</th>
      <td>r3</td>
      <td>6</td>
      <td>7</td>
      <td>8</td>
    </tr>
  </tbody>
</table>
</div>


`openpyxlplus.cell_range.Cells` is just cells stored in a numpy array, so you can access each individual just like using `openpyxl` package


```python
print(rg.cells)
print(f"Type of cell: {type(rg.cells[1,1])}")
print(f"Value of cell: {rg.cells[1,1].value}")
print(f"Font of cell:\n {rg.cells[1,1].font}")
```

    [[<Cell 'Sheet'.A1> <Cell 'Sheet'.B1> <Cell 'Sheet'.C1> <Cell 'Sheet'.D1>]
     [<Cell 'Sheet'.A2> <Cell 'Sheet'.B2> <Cell 'Sheet'.C2> <Cell 'Sheet'.D2>]
     [<Cell 'Sheet'.A3> <Cell 'Sheet'.B3> <Cell 'Sheet'.C3> <Cell 'Sheet'.D3>]
     [<Cell 'Sheet'.A4> <Cell 'Sheet'.B4> <Cell 'Sheet'.C4> <Cell 'Sheet'.D4>]]
    Type of cell: <class 'openpyxl.cell.cell.Cell'>
    Value of cell: 0
    Font of cell:
     <openpyxl.styles.fonts.Font object>
    Parameters:
    name='Calibri', charset=None, family=2.0, b=False, i=False, strike=None, outline=None, shadow=None, condense=None, color=<openpyxl.styles.colors.Color object>
    Parameters:
    rgb=None, indexed=None, auto=None, theme=1, tint=0.0, type='theme', extend=None, sz=11.0, u=None, vertAlign=None, scheme='minor'
    

### Getting Cell Values


```python
print(f"get all values using .cell_values:\n{rg.cell_values}")
print(f"get all values using .cells.get_value():\n{rg.cells.get_value()}")
print(f"get only header values:\n{rg.header.cell_values}")
print(f"get values from a subset(Slicing syntax is identical with numpy):\n{rg.body.cells[:1,:].get_value()}")
print(f"Cells can be converted to SheetCellRange: {type(rg.body.cells[:1,:].to_range())}")

```

    get all values using .cell_values:
    [[None 'c1' 'c2' 'c3']
     ['r1' 0 1 2]
     ['r2' 3 4 5]
     ['r3' 6 7 8]]
    get all values using .cells.get_value():
    [[None 'c1' 'c2' 'c3']
     ['r1' 0 1 2]
     ['r2' 3 4 5]
     ['r3' 6 7 8]]
    get only header values:
    [['c1' 'c2' 'c3']]
    get values from a subset(Slicing syntax is identical with numpy):
    [[0 1 2]]
    Cells can be converted to SheetCellRange: <class 'openpyxlplus.cell_range.SheetCellRange'>
    

### Getting Cell Attributes/Styles

Please refer to [Openpyxl Documentation](https://openpyxl.readthedocs.io/en/stable/styles.html) on details of styles. With this method, you can use the same keywords to get/modify the style of cells.

The most common styles are: font, fill, border, alignment, number_format

#### Getting Styles


```python
rg.cells[1:2,:].get_style("number_format")
```




    array([['General', 'General', 'General', 'General']], dtype=object)



#### Getting Style detail

For example, the font attribute has many details like font name, size, bold, italic etc.


```python
rg.cells[[1],[1]].get_style("font")
```




    array([<openpyxl.styles.fonts.Font object>
           Parameters:
           name='Calibri', charset=None, family=2.0, b=False, i=False, strike=None, outline=None, shadow=None, condense=None, color=<openpyxl.styles.colors.Color object>
           Parameters:
           rgb=None, indexed=None, auto=None, theme=1, tint=0.0, type='theme', extend=None, sz=11.0, u=None, vertAlign=None, scheme='minor'                              ],
          dtype=object)



If we are only interested in one detailed attribute, then we can use `get_style_detail` method


```python
rg.cells[1:2,:].get_style_detail("font","size")
```




    array([[11.0, 11.0, 11.0, 11.0]], dtype=object)



If the attribute is in another object, we can pass a list to the argument. For example, to theme attribute of font.color.


```python
rg.cells[1:2,:].get_style_detail("font",["color","theme"])
```




    array([[1, 1, 1, 1]], dtype=object)



## Change/Modify Cell Attributes/Styles

There are multiple ways to change styles.
1. `set_style()` method: Set target style with provided style, overwrite all style attributes.
   1. By default, `openpyxl.styles` module uses `None` for any attribute that is not specified. The original attribute will be replaced with `None`.
2. `modify_style()` method: only modify attributes that are provided, if provided attribute is `None`, it won't overwrite original attribute.


```python
# modify style
before = rg.body.cells.get_style("number_format")
rg.body.cells.set_style("number_format","0.00%") # change style (number_format)
after = rg.body.cells.get_style("number_format")
rg.body.clear(value=False,formatting=True) # do not clear values, only clear formatting

print(f"Before:\n{before}")
print(f"Before:\n{after}")
```

    Before:
    [['General' 'General' 'General']
     ['General' 'General' 'General']
     ['General' 'General' 'General']]
    Before:
    [['0.00%' '0.00%' '0.00%']
     ['0.00%' '0.00%' '0.00%']
     ['0.00%' '0.00%' '0.00%']]
    

### Comparision between `set_style()` and `modify_style()` 

Refer to [Openpyxl Documentation](https://openpyxl.readthedocs.io/en/stable/styles.html) on what styles can be used and how to set styles.

Note that by using `set_style` method, if any attribute is not specified, the original attribute will be overwritten with None. This is because `openpyxl.styles` by default set all attributes of a style to `None`.


```python
from openpyxl.styles import Font,Alignment,PatternFill,Border,Side
cells_subset = rg.cells[:1,:]
print(f'Before font name:{cells_subset.get_style_detail("font","name")}')
print(f'Before font size:{cells_subset.get_style_detail("font","size")}')
print(f'Before font bold:{cells_subset.get_style_detail("font","b")}')
print("--------Use .set_style method to set bold to True--------")
cells_subset.set_style("font",Font(b=True)) # set style
print(f'After font name:{cells_subset.get_style_detail("font","name")}')
print(f'After font size:{cells_subset.get_style_detail("font","size")}')
print(f'After font bold:{cells_subset.get_style_detail("font","b")}')
cells_subset.to_range().clear(value=False,formatting=True) # clear format only

```

    Before font name:[['Calibri' 'Calibri' 'Calibri' 'Calibri']]
    Before font size:[[11.0 11.0 11.0 11.0]]
    Before font bold:[[False False False False]]
    --------Use .set_style method to set bold to True--------
    After font name:[[None None None None]]
    After font size:[[None None None None]]
    After font bold:[[True True True True]]
    




    <SheetCellRange A1:D1>



If above is not your desired behavior and you want to only replace original attributes with the ones you provided, use `modify_style` instead


```python
from openpyxl.styles import Font,Alignment,PatternFill,Border,Side
cells_subset = rg.cells[1:2,:]
print(f'Before font name:{cells_subset.get_style_detail("font","name")}')
print(f'Before font size:{cells_subset.get_style_detail("font","size")}')
print(f'Before font bold:{cells_subset.get_style_detail("font","b")}')
print("--------Use .modify_style method to set bold to True--------")
cells_subset.modify_style("font",Font(b=True))
print(f'After font name:{cells_subset.get_style_detail("font","name")}')
print(f'After font size:{cells_subset.get_style_detail("font","size")}')
print(f'After font bold:{cells_subset.get_style_detail("font","b")}')
cells_subset.to_range().clear(value=False,formatting=True) # clear format only

```

    Before font name:[['Calibri' 'Calibri' 'Calibri' 'Calibri']]
    Before font size:[[11.0 11.0 11.0 11.0]]
    Before font bold:[[False False False False]]
    --------Use .modify_style method to set bold to True--------
    After font name:[['Calibri' 'Calibri' 'Calibri' 'Calibri']]
    After font size:[[11.0 11.0 11.0 11.0]]
    After font bold:[[True True True True]]
    




    <SheetCellRange A2:D2>


