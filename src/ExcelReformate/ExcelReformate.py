import os
import Tkinter
import tkFileDialog
from xlrd import open_workbook
import xlwt


def add_style(path):
    pass


def save_file(information, path):
    path = path.split('.')
    path[-2] += '_NEW'
    path = '.'.join(path)
    path = tkFileDialog.asksaveasfilename(initialfile=path)
    wbk = xlwt.Workbook(style_compression=2)
    sheet = wbk.add_sheet('Data')
    for row, r_data in enumerate(information[1:], 1):
        for col, data in enumerate(r_data):
            if (row - 1) % 18 == 0:
                if col == 3:
                    sheet.write(row, col, data, style=xlwt.easyxf("borders: top double, left thin"))
                else:
                    sheet.write(row, col, data, style=xlwt.easyxf("borders: top double"))
            elif col > 3 and (row - 1) % 3 == 0:
                sheet.write(row, col, data, style=xlwt.easyxf("borders: top thin"))
            elif col == 3 and (row - 1) % 3 == 0:
                sheet.write(row, col, data, style=xlwt.easyxf("borders: top thin, left thin"))
            elif col == 3:
                sheet.write(row, col, data, style=xlwt.easyxf("borders: left thin"))
            else:
                sheet.write(row, col, data)

    for col, f_data in enumerate(information[0]):
        sheet.write(0, col, f_data, style=xlwt.easyxf("font: bold on; borders: bottom thick"))

    for c in range(6):
        sheet.col(c).width = 18 * 256

    wbk.save(path)

    add_style(path)

    os.startfile(path)


def get_data_from_file(in_path):
    wb = open_workbook(in_path)
    values = []
    for s in wb.sheets():
        for row in range(s.nrows):
            col_value = []
            for col in range(s.ncols):
                value = s.cell(row, col).value
                col_value.append(value)
            values.append(col_value)
    return values


def clean_up(d=None):
    if d is None:
        d = []
    for row in d:
        for i in range(4):
            row.pop(0)


def reformat(data, title):
    new_data = []
    title = title[:6]
    title = map(lambda x: x.replace('_1', ''), title)
    title = map(lambda x: x.replace('1', ''), title)
    for row_data in data:
        main_info = row_data[:3]
        row_data = row_data[3:]
        for idOffset in xrange(0, 42, 7):
            current_id = [row_data[idOffset]]
            id_data = row_data[idOffset + 1:idOffset + 7]
            for scrapOffset in xrange(0, 6, 2):
                scrap_info = id_data[scrapOffset:scrapOffset + 2]
                new_data.append(main_info + current_id + scrap_info)
    new_data.insert(0, title)
    return new_data


def main():
    Tkinter.Tk().withdraw()  # Close the root window
    path = tkFileDialog.askopenfilename(filetypes=[('Excel file', '*.xls')])
    data = get_data_from_file(path)
    title = data[0]
    clean_up(data)
    print "Working..."
    data = reformat(data[1:], title)
    print "Saving..."
    save_file(data, path)
    pass


if __name__ == "__main__":
    os.sys.exit(main())
