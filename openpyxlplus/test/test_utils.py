import unittest
import openpyxlplus.utils as utils

class TestFunctions(unittest.TestCase):
    def test_calc_value_shape(self):

        ## wrap_text = False

        # simple text
        self.assertTupleEqual(
            utils.calc_value_shape("abc"),
            (1,3)
        )

        # simple text with "\n"
        self.assertTupleEqual(
            utils.calc_value_shape("12345\n123"),
            (1,9)
        )

        # number can be converted to text with correct length
        self.assertTupleEqual(
            utils.calc_value_shape(1234.12345678,ndigits=2),
            (1,7)
        )
        self.assertTupleEqual(
            utils.calc_value_shape(12345,ndigits=2),
            (1,5)
        )
        self.assertTupleEqual(
            utils.calc_value_shape(1.2,ndigits=2),
            (1,3)
        )

        ## wrap_text = True
        # simple text with "\n"
        self.assertTupleEqual(
            utils.calc_value_shape("12345\n123\n",wrap_text=True),
            (3,5)
        )
    
    def test_calc_value_size(self):
        fontsize = 1
        min_width = min_height = 3
        max_width = max_height = 5
        width_factor = height_factor = 1
        def calc_value_size(
            value,
        ):
            return(
                utils.calc_value_size(
                    value,
                    wrap_text=True,
                    fontsize=fontsize,
                    min_width=min_width,
                    min_height=min_height,
                    max_width=max_width,
                    max_height=max_height,
                    width_factor=width_factor,
                    height_factor=height_factor,  
                )
            )
        
        # smaller than minimum
        self.assertTupleEqual(
            calc_value_size("11\n11"),
            (3,3)
        )

        # between minimum and maximum
        self.assertTupleEqual(
            calc_value_size("1111\n1111\n1111\n1111"),
            (4,4)
        )

        # greater than maximum
        self.assertTupleEqual(
            calc_value_size("111111\n111111\n111111\n111111\n111111\n111111"),
            (5,5)
        )