"""
Author: Steven17D
"""


import os
import Tkinter
import tkFileDialog
import re
import string
import xlrd
import xlwt


TARGET_STRING = 'Used Roll'


class ExcelFile(object):
    """
    Class for representing an Excel file
    """
    def __init__(self, path):
        """Reads data from excel file"""
        self.path = path
        wb = xlrd.open_workbook(self.path)
        self.data = []
        for sheet in wb.sheets():
            for row_index in range(sheet.nrows):
                column_values = [sheet.cell(row_index, column_index).value for column_index in range(sheet.ncols)]
                self.data.append(column_values)

    def save(self, path=None):
        """Saves file into an xls file"""
        if path is None:
            path = tkFileDialog.asksaveasfilename(initialfile=change_extension(change_basename(self.path)))
            if not path:
                return

        path = change_extension(path)
        wbk = xlwt.Workbook()
        sheet = wbk.add_sheet('Data')
        for row, row_data in enumerate(self.data):
            for col, cell_data in enumerate(row_data):
                sheet.write(row, col, cell_data)

        wbk.save(path)
        os.startfile(path)

    def process(self):
        """
        Removes all target strings and sets cell left to it to zero 
        """
        for row_index, row_data in enumerate(self.data):
            for column_index, cell_data in enumerate(row_data):
                if (isinstance(cell_data, str) or isinstance(cell_data, unicode)) and cell_data == TARGET_STRING:
                    self.data[row_index][column_index] = ''
                    self.data[row_index][column_index+1] = 0


def change_basename(path):
    """my_file.xlsx => my_file_NEW.xls"""
    return re.sub("([\w.]+)\.(\w+)", r"\1_NEW.xls", path)


def change_extension(path):
    """my_file.xlsx => my_file.xls"""
    return re.sub("([\w.]+)\.(\w+)", r"\1.xls", path)


def main():
    Tkinter.Tk().withdraw()  # Close the root window
    path = tkFileDialog.askopenfilename(filetypes=[('Excel file', '*')])
    if not path:
        return 0
    excel_file = ExcelFile(path)
    print "Working..."
    excel_file.process()
    print "Saving..."
    excel_file.save()


if __name__ == "__main__":
    main()
