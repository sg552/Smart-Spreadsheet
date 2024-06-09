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


        self.assertEqual(False, self.table_helper.is_table_top_left(sheet, 50, 6))
        self.assertEqual(False, self.table_helper.is_table_top_left(sheet, 50, 7))

        sheet = self.table_helper.open_file('tests/example_1.xlsx')
        self.assertEqual(True, self.table_helper.is_table_top_left(sheet, 2, 4))
        self.assertEqual(True, self.table_helper.is_table_top_left(sheet, 21, 2))
        self.assertEqual(True, self.table_helper.is_table_top_left(sheet, 42, 4))
        self.assertEqual(True, self.table_helper.is_table_top_left(sheet, 44, 8))

        sheet = self.table_helper.open_file('tests/example_2.xlsx')
        self.assertEqual(True, self.table_helper.is_table_top_left(sheet, 2, 4))
        self.assertEqual(True, self.table_helper.is_table_top_left(sheet, 20, 2))
        self.assertEqual(True, self.table_helper.is_table_top_left(sheet, 40, 4))
        self.assertEqual(True, self.table_helper.is_table_top_left(sheet, 42, 8))

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

        sheet = self.table_helper.open_file('tests/example_1.xlsx')
        self.assertEqual(True, self.table_helper.is_table_top_right(sheet, 2, 9 ))
        self.assertEqual(True, self.table_helper.is_table_top_right(sheet, 21, 7 ))
        self.assertEqual(True, self.table_helper.is_table_top_right(sheet, 42, 6 ))
        self.assertEqual(True, self.table_helper.is_table_top_right(sheet, 44, 11 ))

        sheet = self.table_helper.open_file('tests/example_2.xlsx')
        self.assertEqual(True, self.table_helper.is_table_top_right(sheet, 2, 9 ))
        self.assertEqual(True, self.table_helper.is_table_top_right(sheet, 20, 7 ))
        self.assertEqual(True, self.table_helper.is_table_top_right(sheet, 40, 6 ))
        self.assertEqual(True, self.table_helper.is_table_top_right(sheet, 42, 11 ))

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

        sheet = self.table_helper.open_file('tests/example_2.xlsx')
        self.assertEqual(True, self.table_helper.is_table_bottom_left(sheet, 18, 4 ))
        self.assertEqual(True, self.table_helper.is_table_bottom_left(sheet, 38, 2 ))
        self.assertEqual(True, self.table_helper.is_table_bottom_left(sheet, 50, 4 ))
        self.assertEqual(True, self.table_helper.is_table_bottom_left(sheet, 52, 8 ))

    def test_get_tables_for_example0(self):

        ## 对于example0, 只看第一个和最后一个table就可以了。
        sheet = self.table_helper.open_file('tests/example_0.xlsx')
        self.table_helper.get_tables(sheet)
        print(f"==  tables: {self.table_helper.tables}")
        self.assertEqual(8, len(self.table_helper.tables))

        table = self.table_helper.tables[0]
        self.assertEqual('BALANCE SHEET', table.name)
        self.assertEqual(0, table.top)
        self.assertEqual(2, table.left)
        self.assertEqual(7, table.right)
        self.assertEqual(16, table.bottom)

        table = self.table_helper.tables[-1]
        self.assertEqual(None, table.name)
        self.assertEqual(83, table.top)
        self.assertEqual(2, table.left)
        self.assertEqual(5, table.right)
        self.assertEqual(93, table.bottom)

    def test_get_tables_for_example1(self):

        # 对于example1, 只看第一个和最后一个table就可以了。
        sheet = self.table_helper.open_file('tests/example_1.xlsx')
        self.table_helper.get_tables(sheet)
        self.assertEqual(4, len(self.table_helper.tables))

        table = self.table_helper.tables[0]
        self.assertEqual('BALANCE SHEET', table.name)
        self.assertEqual(2, table.top)
        self.assertEqual(4, table.left)
        self.assertEqual(9, table.right)
        self.assertEqual(18, table.bottom)

        table = self.table_helper.tables[-1]
        self.assertEqual(None, table.name)
        self.assertEqual(44, table.top)
        self.assertEqual(8, table.left)
        self.assertEqual(11, table.right)
        self.assertEqual(54, table.bottom)


    def test_get_tables_for_example2(self):
        # 对于example2, 查看所有
        sheet = self.table_helper.open_file('tests/example_2.xlsx')
        self.table_helper.get_tables(sheet)
        self.assertEqual(4, len(self.table_helper.tables))

        table = self.table_helper.tables[0]
        self.assertEqual('BALANCE SHEET', table.name)
        self.assertEqual(2, table.top)
        self.assertEqual(4, table.left)
        self.assertEqual(9, table.right)
        self.assertEqual(18, table.bottom)

        table = self.table_helper.tables[1]
        self.assertEqual('INCOME STATEMENT', table.name)
        self.assertEqual(20, table.top)
        self.assertEqual(2, table.left)
        self.assertEqual(7, table.right)
        self.assertEqual(38, table.bottom)

        table = self.table_helper.tables[2]
        self.assertEqual('SHAREHOLDERS (Fully Diluted)', table.name)
        self.assertEqual(40, table.top)
        self.assertEqual(4, table.left)
        self.assertEqual(6, table.right)
        self.assertEqual(50, table.bottom)

        table = self.table_helper.tables[3]
        self.assertEqual(None, table.name)
        self.assertEqual(42, table.top)
        self.assertEqual(8, table.left)
        self.assertEqual(11, table.right)
        self.assertEqual(52, table.bottom)


if __name__ == '__main__':
    unittest.main()
