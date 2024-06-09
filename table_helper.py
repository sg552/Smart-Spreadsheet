from table import Table
from openpyxl import load_workbook

class TableHelper:
    HEADER_COLOR_INDEX = 4
    def __init__(self):
        self.max_blank_column = 1
        self.max_blank_row = 1
        self.max_row = 100
        self.max_columns = 20

        self.tables = []

    def is_my_digit(self, cell_value):
        return cell_value.replace(',', '').replace('$', '').strip().isdigit()

    def is_my_string(self, cell_value):
        return type(cell_value) == str

    def is_table_header(self, cell):
        return cell.fill.start_color.index == self.HEADER_COLOR_INDEX

    def is_gray_area(self, cell):
        return cell.fill.start_color.index == 'FF7F7F7F'

    def open_file(self, path):
        workbook = load_workbook(filename = path, data_only = True)
        sheet = workbook.active
        return sheet

    # 左边没有任何空白，
    # 左边是2个空白，本身是 HEADER_COLOR_INDEX的，一定是table起点
    # 左边可以有一个空白，再左边不是空白的话，必须是一个header.
    # 蓝色可以向右有一个空白，然后继续蓝色
    # 可以向下连续一行
    def is_table_top_left(self, sheet, i, j):
        rows = list(sheet.iter_rows())
        #print(f"== cell: {rows[i][j]}, value: {rows[i][j].value}")
        if (j == 0 and self.is_table_header(rows[i][j]) ) or \
                (j == 1 and self.is_table_header(rows[i][j]) and  rows[i][j - 1].value == None)  or \
                (j >= 2 and self.is_table_header(rows[i][j]) and rows[i][j - 2].value == None and rows[i][j - 1].value == None) or \
                (j >= 2 and self.is_table_header(rows[i][j]) and (rows[i][j - 2].value != None and not self.is_table_header(rows[i][j - 2]) ) and rows[i][j - 1].value == None):

            # 如果左边的是一个蓝色的, 则不是  ( 对应example 0 G51, F51)
            if j >= 2 and self.is_table_header(rows[i][j-1]):
                return False

            # 判断一下上一行的cell是否是header, 是的话，本身则不是header了。
            if i == 0:
                return True
            else:
                return not self.is_table_header(rows[i-1][j])

        else:
            return False

    # 右边两个肯定都是空白, 可能一行存在多个。找到立刻return
    def is_table_top_right(self, sheet, i, j):
        rows = list(sheet.iter_rows())
        row = rows[i]

        return self.is_table_header(row[j]) and \
                (not self.is_table_header(row[j + 1])) and \
                (not self.is_table_header(row[j + 2]))

     #必须都是string ,不能是纯数字
    # 从当前cell往下找 , 可以有一个空行, 两个空行就结束
    # 从当前cell往上找 , 可以有一个空行, 两个空行就结束
     #不能出现新的header color
    def is_table_bottom_left(self, sheet, i, j):
        rows = list(sheet.iter_rows())

        #print(f"== checking: {rows[i][j]}, value: {rows[i][j].value}, +1: {rows[i+1][j].value}, +2: {rows[i+2][j].value}")

        # case1: 本身是str, 下面有2个空白行，一定是结束
        # 或者有个 0
        if self.is_my_string(rows[i][j].value)  and \
                (rows[i+1][j].value == None) and \
                (rows[i+2][j].value == None or rows[i+2][j].value == 0) :
            return True

        # case2: 本身是str,下面有一个空白行，外加一个带颜色的cell
        if self.is_my_string(rows[i][j].value)  and \
                (rows[i+1][j].value == None and self.is_table_header(rows[i+2][j])):
            return True

        return False


    # 获得该table top, left, right , bottom
    def get_tables(self, sheet):

        # get all tables top left
        for i, row in enumerate(sheet.iter_rows()):
            for j, cell in enumerate(row):
                if i > self.max_row:
                    break
                #print(f"cell: {cell.value}, i: {i}, j: {j}, {row[j].value}")
                if(self.is_table_top_left(sheet, i, j)):
                    table = Table([[]])
                    print(f"=== found table top left: column: {cell.column}, row: {cell.row}, value: {cell.value}")
                    table.name = cell.value
                    table.left = cell.column - 1
                    table.top = cell.row - 1
                    self.tables.append(table)

        for table in self.tables:
            #print(f"==checking for table: {table}")
            # get all tables top right , 只考察一行即可( table.top)
            target_row = list(sheet.iter_rows())[table.top]
            for cell in target_row:
                #print(f"==checking for top_right: {cell.value}")
                if self.is_table_top_right(sheet, cell.row - 1, cell.column - 1):
                    table.right = cell.column - 1
                    break

            # get all tables left bottom, 只考察一列即可（table.left)
            rows = list(sheet.iter_rows())

            target_column = []
            for m, row in enumerate(rows):
                # 在table.top 上面的，不考虑
                if m <= table.top:
                    continue

                if m > self.max_row:
                    break
                target_column.append(row[table.left])

            #print(f"==checking for bottom_left: {target_column}")
            for cell in target_column:
                if self.is_table_bottom_left(sheet, cell.row - 1, cell.column - 1):
                    #print(f"=== found bottom, cell: #{cell.value}")
                    table.bottom = cell.row - 1
                    break

        print("=== done")












