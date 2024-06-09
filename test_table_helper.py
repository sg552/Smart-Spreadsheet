import unittest

from table_helper import TableHelper

class TestTableHelper(unittest.TestCase):
    def setUp(self):
        self.table_helper = TableHelper()

    def test_is_table_top_left(self):
        sheet = self.table_helper.open_file('tests/example_0.xlsx')
        self.assertEqual(True, self.table_helper.is_table_top_left(sheet, 1-1, 3-1))
        self.assertEqual(True, self.table_helper.is_table_top_left(sheet, 19-1, 3-1))
        self.assertEqual(True, self.table_helper.is_table_top_left(sheet, 39-1, 3-1))
        self.assertEqual(True, self.table_helper.is_table_top_left(sheet, 51-1, 3-1))
        self.assertEqual(True, self.table_helper.is_table_top_left(sheet, 62-1, 3-1))
        self.assertEqual(True, self.table_helper.is_table_top_left(sheet, 67-1, 3-1))
        self.assertEqual(True, self.table_helper.is_table_top_left(sheet, 73-1, 3-1))
        self.assertEqual(False, self.table_helper.is_table_top_left(sheet, 74-1, 3-1))
        self.assertEqual(True, self.table_helper.is_table_top_left(sheet, 84-1, 3-1))

        sheet = self.table_helper.open_file('tests/example_1.xlsx')
        self.assertEqual(True, self.table_helper.is_table_top_left(sheet, 2, 4))
        self.assertEqual(True, self.table_helper.is_table_top_left(sheet, 21, 2))
        self.assertEqual(True, self.table_helper.is_table_top_left(sheet, 42, 4))
        self.assertEqual(True, self.table_helper.is_table_top_left(sheet, 44, 8))

    def test_is_table_top_right(self):
        sheet = self.table_helper.open_file('tests/example_0.xlsx')
        self.assertEqual(False, self.table_helper.is_table_top_right(sheet, 1-1, 6 ))
        self.assertEqual(True, self.table_helper.is_table_top_right(sheet, 1-1, 7 ))
        self.assertEqual(False, self.table_helper.is_table_top_right(sheet, 1-1, 8 ))

        self.assertEqual(True, self.table_helper.is_table_top_right(sheet, 19-1, 7 ))
        self.assertEqual(True, self.table_helper.is_table_top_right(sheet, 39-1, 4 ))
        self.assertEqual(True, self.table_helper.is_table_top_right(sheet, 51-1, 7 ))
        self.assertEqual(True, self.table_helper.is_table_top_right(sheet, 62-1, 7 ))
        self.assertEqual(True, self.table_helper.is_table_top_right(sheet, 67-1, 7 ))
        self.assertEqual(True, self.table_helper.is_table_top_right(sheet, 73-1, 7 ))
        self.assertEqual(False, self.table_helper.is_table_top_right(sheet, 74-1, 7 ))
        self.assertEqual(True, self.table_helper.is_table_top_right(sheet, 84-1, 5 ))

    def test_is_table_bottom_left(self):
        sheet = self.table_helper.open_file('tests/example_0.xlsx')
        self.assertEqual(False, self.table_helper.is_table_bottom_left(sheet, 15, 2 ))
        self.assertEqual(True, self.table_helper.is_table_bottom_left(sheet, 16, 2 ))
        self.assertEqual(False, self.table_helper.is_table_bottom_left(sheet, 17, 2 ))
        self.assertEqual(False, self.table_helper.is_table_bottom_left(sheet, 21, 2 ))
        self.assertEqual(True, self.table_helper.is_table_bottom_left(sheet, 36, 2 ))
        self.assertEqual(True, self.table_helper.is_table_bottom_left(sheet, 48, 2 ))
        self.assertEqual(True, self.table_helper.is_table_bottom_left(sheet, 59, 2 ))
        self.assertEqual(True, self.table_helper.is_table_bottom_left(sheet, 64, 2 ))
        self.assertEqual(True, self.table_helper.is_table_bottom_left(sheet, 70, 2 ))
        self.assertEqual(True, self.table_helper.is_table_bottom_left(sheet, 93, 2 ))


        sheet = self.table_helper.open_file('tests/example_1.xlsx')
        self.assertEqual(False, self.table_helper.is_table_bottom_left(sheet, 2, 4 ))
        self.assertEqual(False, self.table_helper.is_table_bottom_left(sheet, 17, 4 ))
        self.assertEqual(True, self.table_helper.is_table_bottom_left(sheet, 18, 4 ))
        self.assertEqual(False, self.table_helper.is_table_bottom_left(sheet, 19, 4 ))
        self.assertEqual(True, self.table_helper.is_table_bottom_left(sheet, 39, 2 ))
        self.assertEqual(True, self.table_helper.is_table_bottom_left(sheet, 52, 4 ))
        self.assertEqual(True, self.table_helper.is_table_bottom_left(sheet, 54, 8 ))

    #def test_get_single_table(self):

    #    sheet = self.table_helper.open_file('tests/example_0.xlsx')
    #    table = self.table_helper.get_single_table(sheet)
        #self.assertEqual(matrix1, table.data)

if __name__ == '__main__':
    unittest.main()
