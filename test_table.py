import unittest

from table import Table

class TestTable(unittest.TestCase):
    def setUp(self):
        data = [
                #[  None, "date1", "date2" ],
                [  '', "date1", "date2" ],
                [  "item1", 110, 120 ],
                [  "item2", 210, 220 ],
                [  "item3", 310, 320 ],
                ]
        self.table = Table(data)

    def test_to_text(self):
        result = self.table.to_text()
        self.assertEqual("|date1|date2", result.splitlines()[0])
        self.assertEqual("item1|110|120", result.splitlines()[1])
        self.assertEqual("item2|210|220", result.splitlines()[2])

if __name__ == '__main__':
    unittest.main()
