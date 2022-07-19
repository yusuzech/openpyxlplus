import unittest
from openpyxl import Workbook,worksheet
import pandas as pd
import numpy as np
from openpyxlplus import writer,cell_range

def is_range_equal_to_array(ws,range_string,array):
    ret = np.array_equal(
        cell_range.SheetCellRange(ws,range_string).cells.get_value(),
        array
    )
    return(ret)

class TestFunctions(unittest.TestCase):
    def test_write_value(self):
        wb = Workbook()
        ws = wb.active
        writer.write_value(1,ws,ws["A1"])
        self.assertEqual(ws["A1"].value,1)

    def test_write_list(self):
        wb = Workbook()
        ws = wb.active
        # down
        writer.write_list([1,2,3],ws,ws["A1"])
        self.assertEqual(
            cell_range.SheetCellRange(ws,"A1:A3").cell_values.flatten().tolist(),
            [1,2,3]
        )
        # up
        # raise Error
        with self.assertRaises(ValueError):
            rg = writer.write_list([1,2,3],ws,ws["A1"],direction="up")
        writer.write_list([1,2,3],ws,ws["A3"],direction="up")
        self.assertEqual(
            cell_range.SheetCellRange(ws,"A1:A3").cell_values.flatten().tolist(),
            [3,2,1]
        )
        # right
        writer.write_list([1,2,3],ws,ws["A1"],direction="right")
        self.assertEqual(
            cell_range.SheetCellRange(ws,"A1:C1").cell_values.flatten().tolist(),
            [1,2,3]
        )

        # left
        with self.assertRaises(ValueError):
            rg = writer.write_list([1,2,3],ws,ws["A1"],direction="left")
        writer.write_list([1,2,3],ws,ws["C1"],direction="left")
        self.assertEqual(
            cell_range.SheetCellRange(ws,"A1:C1").cell_values.flatten().tolist(),
            [3,2,1]
        )

    def test_write_array(self):
        wb = Workbook()
        ws = wb.active
        # write list of list
        writer.write_array([[1,2],[3,4]],ws,ws["A1"])
        self.assertEqual(
            cell_range.SheetCellRange(ws,"A1:B2").cell_values.tolist(),
            [[1,2],[3,4]]
        )

        # fail if dimensions mismatch
        with self.assertRaises(ValueError):
            writer.write_array([[1],[3,4]],ws,ws["A1"])

        # write array
        writer.write_array(np.array([[1,2,3],[2,3,4]]),ws,ws["C3"])
        self.assertEqual(
            cell_range.SheetCellRange(ws,"C3:E4").cell_values.tolist(),
            [[1,2,3],[2,3,4]]
        )

    def test_write_value_merged(self):
        wb = Workbook()
        ws = wb.active
        # form A1 to K1
        writer.write_value_merged(1,ws,right=10)
        self.assertTrue(
            worksheet.cell_range.CellRange("A1:K1") in ws.merged_cells.ranges
        )

        # form A2 to K3
        writer.write_value_merged(1,ws,ws.cell(2,1),down=1,right=10)
        self.assertTrue(
            worksheet.cell_range.CellRange("A2:K3") in ws.merged_cells.ranges
        )

        # negative signs to write up and left
        writer.write_value_merged(1,ws,ws.cell(10,10),up=2,left=2)
        self.assertTrue(
            worksheet.cell_range.CellRange("H8:J10") in ws.merged_cells.ranges
        )
        
        # check if value is correct and centered
        self.assertEqual(ws.cell(8,8).value,1)
        # print(ws.cell(8,8).alignment)
        self.assertEqual(ws.cell(8,8).alignment.horizontal,"center")
        self.assertEqual(ws.cell(8,8).alignment.vertical,"center")

class testWriteDataFrame(unittest.TestCase):
    def setUp(self):
        self.df = pd.DataFrame({
            "percent":[0.01,0.5,2],
            "comma":[1000,500,1000000],
            "mixed":["bold_red","centered","all_border"]
        })
        # initialize class-wise workbook
        self.wb = Workbook()

    def tearDown(self):
        self.wb.close()

    def test_anchor(self):
        array_value = np.array(
            [['percent', 'comma', 'mixed'],
            [0.01, 1000, 'bold_red'],
            [0.5, 500, 'centered'],
            [2.0, 1000000, 'all_border']],dtype='O'
        )
        # at 2,2
        self.wb.create_sheet("sheet1")
        ws = self.wb["sheet1"]
        writer.write_dataframe(
            data = self.df,
            ws = ws,
            cell = ws.cell(2,2),
            index = False,
            header = True
        )
        self.assertTrue(is_range_equal_to_array(ws,"B2:D5",array_value))

        # at 10,3
        self.wb.create_sheet("sheet2")
        ws = self.wb["sheet2"]
        writer.write_dataframe(
            data = self.df,
            ws = ws,
            cell = ws.cell(11,3),
            index=False,
            header=True
        ) 
        self.assertTrue(is_range_equal_to_array(ws,"C11:E14",array_value))

    def test_index_header(self):
        # index and header
        array_value = np.array(
            [[None,'percent', 'comma', 'mixed'],
            [None,None,None,None],
            [0,0.01, 1000, 'bold_red'],
            [1,0.5, 500, 'centered'],
            [2,2.0, 1000000, 'all_border']]
        )
        self.wb.create_sheet("sheet3")
        ws = self.wb["sheet3"]
        writer.write_dataframe(
            data = self.df,
            ws = ws,
            index=True,
            header=True
        ) 
        self.assertTrue(is_range_equal_to_array(ws,"A1:D5",array_value))
        # index
        array_value = np.array(
            [[None,"percent","comma","mixed"],
            [0,0.01, 1000, 'bold_red'],
            [1,0.5, 500, 'centered'],
            [2,2.0, 1000000, 'all_border']],dtype="O"
        )
        self.wb.create_sheet("sheet3")
        ws = self.wb["sheet3"]
        writer.write_dataframe(
            data = self.df,
            ws = ws,
            index=True,
            header=False
        )
        self.assertTrue(is_range_equal_to_array(ws,"A1:D4",array_value))
