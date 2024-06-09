import unittest

from table_helper import TableHelper

class TestTableHelper(unittest.TestCase):
    def setUp(self):
        self.table_helper = TableHelper('tests/example_0.xlsx')

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

    def test_is_table_top_right(self):

        sheet = self.table_helper.open_file('tests/example_0.xlsx')
        self.assertEqual(False, self.table_helper.is_table_top_right(sheet, 1-1, 6 ))
        self.assertEqual(True, self.table_helper.is_table_top_right(sheet, 1-1, 7 ))
        self.assertEqual(False, self.table_helper.is_table_top_right(sheet, 1-1, 8 ))

    #def test_get_single_table(self):

    #    sheet = self.table_helper.open_file('tests/example_0.xlsx')
    #    table = self.table_helper.get_single_table(sheet)
        #self.assertEqual(matrix1, table.data)

if __name__ == '__main__':
    unittest.main()
