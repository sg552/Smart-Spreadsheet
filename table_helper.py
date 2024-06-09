class TableHelper:
    def __init__(self):
        self.max_blank_column = 1
        self.max_blank_row = 1
        self.tables = []

    def say_hi(self):
        return "hi"

    # 一个最简单的table
    #table1  | column1   | column2
    #row_1   | 1         |   2
    #row_2   | 2         |   3
    def get_single_table(self, matrix):
        one = 1
        table_name = f"table-{one}"
        for i, row in enumerate(matrix):

            for j, cell in enumerate(row):
                # 第一行，应该是： 第一个可有可无。有的话就是table-name
                if i == 0 and j == 0:
                    if cell == None or !(cell.isnumeric()):
                        table_name = cell

                # 看后续有多少个列, 要求列不能断档, 最多断一次
                else if i == 0 and j != 0:
                    if cell is None:
                        j += 1
                        continue


            # 对于其他行，则挨个观察
            else:





