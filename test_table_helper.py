import unittest

from table_helper import TableHelper

class TestTableHelper(unittest.TestCase):
    def setUp(self):
        self.table_helper = TableHelper()

    def testHi(self):
        result = self.table_helper.say_hi()
        self.assertEqual("hi", result)

if __name__ == '__main__':
    unittest.main()
