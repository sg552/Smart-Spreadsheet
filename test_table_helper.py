import unittest

from table_helper import TableHelper

class TestTableHelper(unittest.TestCase):
    def setUp(self):
        self.table_helper = TableHelper('tests/example_0.xlsx')

    def test_is_table_top_left(self):
        sheet = self.table_helper.open_file('tests/example_0.xlsx')

        #=== found table top left: column: 3, row: 1
        #=== found table top left: column: 3, row: 19
        #=== found table top left: column: 3, row: 39
        #=== found table top left: column: 3, row: 51
        #=== found table top left: column: 3, row: 62
        #=== found table top left: column: 3, row: 67
        #=== found table top left: column: 3, row: 73
        #=== found table top left: column: 3, row: 74
        #=== found table top left: column: 3, row: 84
        self.assertEqual(True, self.table_helper.is_table_top_left(sheet, 1-1, 3-1))
        self.assertEqual(True, self.table_helper.is_table_top_left(sheet, 19-1, 3-1))
        self.assertEqual(True, self.table_helper.is_table_top_left(sheet, 39-1, 3-1))
        self.assertEqual(True, self.table_helper.is_table_top_left(sheet, 51-1, 3-1))
        self.assertEqual(True, self.table_helper.is_table_top_left(sheet, 62-1, 3-1))
        self.assertEqual(True, self.table_helper.is_table_top_left(sheet, 67-1, 3-1))
        self.assertEqual(True, self.table_helper.is_table_top_left(sheet, 73-1, 3-1))
        self.assertEqual(False, self.table_helper.is_table_top_left(sheet, 74-1, 3-1))
        self.assertEqual(True, self.table_helper.is_table_top_left(sheet, 84-1, 3-1))

    def test_get_single_table(self):

        sheet = self.table_helper.open_file('tests/example_0.xlsx')
        table = self.table_helper.get_single_table(sheet)
        #self.assertEqual(matrix1, table.data)

if __name__ == '__main__':
    unittest.main()
