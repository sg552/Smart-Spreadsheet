import unittest

from table import Table

# 似乎没用到
class TestTable(unittest.TestCase):
    def setUp(self):
        self.table = Table()

    def test_to_text(self):
        self.table.top = 2
        self.table.left = 4
        self.table.right = 9
        self.table.bottom = 18

        result = self.table.to_text()
        self.assertEqual("|date1|date2", result.splitlines()[0])
        self.assertEqual("item1|110|120", result.splitlines()[1])
        self.assertEqual("item2|210|220", result.splitlines()[2])

if __name__ == '__main__':
    unittest.main()
