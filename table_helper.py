from table import Table
from openpyxl import load_workbook

class TableHelper:
    HEADER_COLOR_INDEX = 4
    def __init__(self, excel_file_path):
        self.max_blank_column = 1
        self.max_blank_row = 1
        self.tables = []
        self.excel_file_path = excel_file_path

    def is_my_digit(self, cell_value):
        return cell_value.replace(',', '').replace('$', '').strip().isdigit()
    def is_my_string(self, cell_value):
        return cell_value != '' and (not is_my_digit(cell_value))

    def is_table_header(self, cell):
        return cell.fill.start_color.index == self.HEADER_COLOR_INDEX

    def is_gray_area(self, cell):
        return cell.fill.start_color.index == 'FF7F7F7F'

    def open_file(self, path):
        workbook = load_workbook(filename = self.excel_file_path, data_only = True)
        sheet = workbook.active
        return sheet

    def is_table_top_left(self, sheet, i, j):
        rows = list(sheet.iter_rows())
        #print(f"== cell: {rows[i][j]}")
        if (j == 0 and self.is_table_header(rows[i][j]) ) or \
                (j == 1 and self.is_table_header(rows[i][j]) and  rows[i][j - 1].value == None)  or \
                (j == 2 and self.is_table_header(rows[i][j]) and rows[i][j - 2].value == None and rows[i][j - 1].value == None):
            # 判断一下上一行的cell是否是header, 是的话，本身则不是header了。
            if i == 0:
                return True
            else:
                return not self.is_table_header(rows[i-1][j])
        else:
            return False


    # 获得该table top, left, right , bottom
    def get_single_table(self, sheet):
        table = Table([[]])

        for i, row in enumerate(sheet.iter_rows()):
            for j, cell in enumerate(row):
                if i > 100:
                    break

                print(f"cell: {cell.value}, i: {i}, j: {j}, {row[j].value}")

                # 左边没有任何空白，
                # 左边是2个空白，本身是 HEADER_COLOR_INDEX的，一定是table起点
                # 蓝色可以向右有一个空白，然后继续蓝色
                # 可以向下连续一行

                if (j == 0 and self.is_table_header(cell) ) or \
                        (j == 1 and self.is_table_header(cell) and  row[j - 1].value == None)  or \
                        (j == 2 and self.is_table_header(cell) and row[j - 2].value == None and row[j - 1].value == None):
                    # 判断一下上一行的cell是否是header, 是的话，本身则不是header了。
                    if i == 0:
                        return True
                    else:
                        return not self.is_table_header(sheet.iter_rows()[i-1][j])

                    print(f"=== found table top left: column: {cell.column}, row: {cell.row}")
                    table.left = cell.column
                    table.top = cell.row

                # 判断table header( top left ): 根据颜色吧。
                #if cell = ''
                #    and cell_down != ''
                #    and (not is_my_digit(cell_down))
                #    and cell_left == ''
                #    and (
                #            is_my_string(cell_right)
                #            or ( cell_right == '' and is_my_string(cell_right))
                #            ):
                #    table.left = j
                #    table.top = i

                #else:
                #    pass

                #if cell == '' and cell_right == '' and cell_right_right == '':
                #    table.right = j

                #if cell == '' and



