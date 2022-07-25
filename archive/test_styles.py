# import unittest
# import openpyxlplus.styles
# from openpyxlplus.styles import get_sides
# from openpyxl import load_workbook
# from openpyxl import Workbook
# from openpyxl.styles import NamedStyle,Font,Alignment,Border,Side
# import openpyxl.utils.cell as converter
# from numpy import array as Array
# from pandas import DataFrame
# import os

# class testAdvanced(unittest.TestCase):
#     def setUp(self):
#         fname = os.path.join(os.path.dirname(__file__), 'test_readonly.xlsx')
#         self.wb = load_workbook(fname)

#     def tearDown(self):
#         self.wb.close()

#     def test_borders(self):
#         # get_sides
#         ws = self.wb["test_borders"]

#         self.assertEqual(
#             openpyxlplus.styles.get_sides(ws["B2"].border)[0]["direction"],
#             "right"
#         )
#         self.assertEqual(
#             openpyxlplus.styles.get_sides(ws["B4"].border)[0]["direction"],
#             "left"
#         )
#         self.assertEqual(
#             openpyxlplus.styles.get_sides(ws["B6"].border)[0]["direction"],
#             "top"
#         )
#         self.assertEqual(
#             openpyxlplus.styles.get_sides(ws["B8"].border)[0]["direction"],
#             "bottom"
#         )

#         self.assertSetEqual(
#             set([x["direction"] for x in\
#                  openpyxlplus.styles.get_sides(ws["B10"].border)]),
#             set(["left","right","top","bottom"])
#         )
#         #merge_border

#         # new = True/False works
#         self.assertEqual(
#             openpyxlplus.styles.merge_border(
#                 Border(left=Side(style="thin")),
#                 Border(left=Side(style="thick")),
#             ).left.style,
#             "thick"
#         )

#         self.assertEqual(
#             openpyxlplus.styles.merge_border(
#                 Border(left=Side(style="thin")),
#                 Border(left=Side(style="thick")),
#                 new = False
#             ).left.style,
#             "thin"
#         )

#         # merging to original borders work
#         new_borders = openpyxlplus.styles.merge_border(
#                 Border(left=Side(style="thin")),
#                 Border(right=Side(style="thick")),
#                 new = False
#             )
#         self.assertEqual(new_borders.left.style,"thin")
#         self.assertEqual(new_borders.right.style,"thick")

#         # writing/appending on excel works
#         wb = Workbook()
#         ws = wb.active
#         ws["A1"].border = Border(left=Side(style="thin"))
#         openpyxlplus.styles.cell_append_border(
#             ws["A1"],Border(right=Side(style="thick"))
#         )
#         cell_border = ws["A1"].border
#         self.assertEqual(cell_border.left.style,"thin")
#         self.assertEqual(cell_border.right.style,"thick")
        
#         # aboundaries_append_border, range_append_border works for single cell
#         three_sides_border = Border(
#             left=Side(style="thin"),
#             right=Side(style="thin"),
#             top=Side(style="thin")
#         )

#         openpyxlplus.styles.boundaries_append_border(ws,1,1,1,1,three_sides_border)
#         cell_border = ws["A1"].border
#         self.assertEqual(cell_border.left.style,"thin")
#         self.assertEqual(cell_border.right.style,"thin")
#         self.assertEqual(cell_border.top.style,"thin")
#         self.assertIsNone(cell_border.bottom)

#         openpyxlplus.styles.range_append_border(ws,"A3",three_sides_border)
#         cell_border = ws["A3"].border
#         self.assertEqual(cell_border.left.style,"thin")
#         self.assertEqual(cell_border.right.style,"thin")
#         self.assertEqual(cell_border.top.style,"thin")
#         self.assertIsNone(cell_border.bottom)
#         # aboundaries_append_border, range_append_border works for more than 
#         # one cell
#         openpyxlplus.styles.boundaries_append_border(ws,1,1,2,2,three_sides_border)
#         openpyxlplus.styles.range_append_border(ws,"C1:D2",three_sides_border)
#         for row in converter.rows_from_range("A1:D2"):
#             for cell_address in row:
#                 cell_border = ws[cell_address].border
#                 self.assertEqual(cell_border.left.style,"thin")
#                 self.assertEqual(cell_border.right.style,"thin")
#                 self.assertEqual(cell_border.top.style,"thin")


#         ws = wb.create_sheet("newsheet1")
#         thin_side = Side(style="thin")
#         # range_append_outline,boundaries_append_outline works for single cell
#         openpyxlplus.styles.boundaries_append_outline(ws,1,1,1,1,thin_side)
#         openpyxlplus.styles.range_append_outline(ws,"C1",thin_side)
#         for cell_address in ["A1","C1"]:
#             cell_border = ws[cell_address].border
#             self.assertEqual(cell_border.left.style,"thin")
#             self.assertEqual(cell_border.right.style,"thin")
#             self.assertEqual(cell_border.top.style,"thin")
#             self.assertEqual(cell_border.bottom.style,"thin")

#         # range_append_outline,boundaries_append_outline works for more than
#         # one cell
#         ws = wb.create_sheet("newsheet2")
#         openpyxlplus.styles.boundaries_append_outline(ws,1,1,3,3,thin_side)
#         openpyxlplus.styles.range_append_outline(ws,"E1:G3",thin_side)
#         for range_string in ["A1:C3","E1:G3"]:
#             arr_list = []
#             for row in converter.rows_from_range(range_string):
#                 arr_list.append(row)
#             arr = Array(arr_list)
#             # top left
#             self.assertSetEqual(
#                 set([x["direction"] for x in get_sides(ws[arr[0,0]].border)]),
#                 set(["left","top"])
#             )
#             # top middle
#             self.assertSetEqual(
#                 set([x["direction"] for x in get_sides(ws[arr[0,1]].border)]),
#                 set(["top"])
#             )
#             # top right
#             self.assertSetEqual(
#                 set([x["direction"] for x in get_sides(ws[arr[0,2]].border)]),
#                 set(["top","right"])
#             )            
#             # middle left
#             self.assertSetEqual(
#                 set([x["direction"] for x in get_sides(ws[arr[1,0]].border)]),
#                 set(["left"])
#             )       
#             # middle middle
#             self.assertEqual(
#                 len([x["direction"] for x in get_sides(ws[arr[1,1]].border)]),
#                 0
#             )       
#             # middle right
#             self.assertSetEqual(
#                 set([x["direction"] for x in get_sides(ws[arr[1,2]].border)]),
#                 set(["right"])
#             )       
#             # bottom left
#             self.assertSetEqual(
#                 set([x["direction"] for x in get_sides(ws[arr[2,0]].border)]),
#                 set(["bottom","left"])
#             )       
#             # bottom middle
#             self.assertSetEqual(
#                 set([x["direction"] for x in get_sides(ws[arr[2,1]].border)]),
#                 set(["bottom"])
#             )       
#             # bottom right
#             self.assertSetEqual(
#                 set([x["direction"] for x in get_sides(ws[arr[2,2]].border)]),
#                 set(["bottom","right"])
#             )            
#         wb.close()

#     def test_style(self):
#         wb = Workbook()
#         ws = wb.active
#         style_bold = NamedStyle("custom_bold",font=Font(bold=True))
#         # range_apply_style, boundaries_apply_style works for single cell
#         openpyxlplus.styles.boundaries_apply_style(ws,1,1,1,1,style_bold)
#         openpyxlplus.styles.range_apply_style(ws,"A3",style_bold)
#         self.assertEqual(ws["A1"].font.b,True)
#         self.assertEqual(ws["A3"].font.b,True)

#         # range_apply_style, boundaries_apply_style works for more than one cell
#         ws = wb.create_sheet("newsheet1")
#         openpyxlplus.styles.boundaries_apply_style(ws,1,1,2,2,style_bold)
#         openpyxlplus.styles.range_apply_style(ws,"C1:D2",style_bold)
#         for row in converter.rows_from_range("A1:D2"):
#             for cell_string in row:
#                 self.assertEqual(ws[cell_string].font.b,True)

#         # dataframe,list,array works
#         # array
#         openpyxlplus.styles.range_apply_style(
#             ws,"A1:B3",Array([style_bold]*6).reshape((3,2))
#         )
#         # dataframe
#         openpyxlplus.styles.range_apply_style(
#             ws,"C1:D3",DataFrame({"x":[style_bold]*3,"y":[style_bold]*3})
#         )
#         # list
#         openpyxlplus.styles.range_apply_style(
#             ws,"E1:F3",[[style_bold]*2,[style_bold]*2,[style_bold]*2]
#         )
#         for row in converter.rows_from_range("A1:F3"):
#             for cell_string in row:
#                 self.assertEqual(ws[cell_string].font.b,True)

#         # raise error if dimension incorrect
#         with self.assertRaises(ValueError):
#             openpyxlplus.styles.range_apply_style(
#             ws,"E1:F3",Array([style_bold]*6).reshape((2,3))
#         )


#     def test_column_width(self):
#         # modify those values according to adjust_column_width() function
#         default_values = {
#             "min_width":13,
#             "max_width":50,
#             "modifier":1.2
#         }
#         ws = self.wb["test_column_width_1"]
#         # adjust column width accordingly
#         openpyxlplus.styles.adjust_column_width(ws,"A1:D2")
#         self.assertEqual(
#             ws.column_dimensions['A'].width,
#             default_values["min_width"]
#         )
#         self.assertEqual(
#             ws.column_dimensions['B'].width,
#             default_values["min_width"]
#         )
#         self.assertEqual(
#             ws.column_dimensions['C'].width,
#             round(len(ws["C1"].value) * default_values["modifier"])
#         )
#         self.assertEqual(
#             ws.column_dimensions['D'].width,
#             default_values["max_width"]
#         )

#         # use specified range correctly, and igonres numbers
#         ws = self.wb["test_column_width_2"]
#         openpyxlplus.styles.adjust_column_width(ws,[1,1,2,10])
#         for col in ["A","B","C","D"]:
#             self.assertEqual(
#                 ws.column_dimensions[col].width,
#                 default_values["min_width"],
#                 f"column {col}"
#             )