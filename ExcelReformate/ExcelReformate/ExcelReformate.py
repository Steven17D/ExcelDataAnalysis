import os
import Tkinter
import tkFileDialog
from xlrd import open_workbook
import xlwt

def saveFile(information, path = ''):
    path = path.split('.')
    path[-2] += '_NEW'
    path = '.'.join(path)

    wbk = xlwt.Workbook() 
    sheet = wbk.add_sheet('Data')

    for row, Rdata in enumerate(information):
        for col, data in enumerate(Rdata):
            sheet.write(row,col,data)
    wbk.save(path)
    os.startfile(path)

def getDataFromFile(in_path):
    wb = open_workbook(in_path)
    for s in wb.sheets():
        #print 'Sheet:',s.name
        values = []
        for row in range(s.nrows):
            col_value = []
            for col in range(s.ncols):
                value  = (s.cell(row,col).value)
                col_value.append(value)
            values.append(col_value)
    return values

def cleanUp(d = []):#, columnsToDelete = []):
    for row in d:
        for i in range(4):#columnsToDelete:
            row.pop(0)#i)

def reformate(data,title):
    newData = []
    title = title[:10]
    title = map(lambda x: x.replace('_1','') ,title)
    for row in data:
        newData.append(row[:10]) #Main row + first body
        for x in range(1,6):
            offset = 3 + x*7
            newData.append(['','',''] + row[offset:offset+7])
    newData.insert(0,title)
    return newData

def main():

    Tkinter.Tk().withdraw() # Close the root window
    path = tkFileDialog.askopenfilename(filetypes = [('Excel file','*.xls')])
    data = getDataFromFile(path)
    title = data[0]
    cleanUp(data)#, map(lambda x: title.index(x),['ID','Oprator','StartTime_Run','Shift']))
    data = reformate(data[1:],title)
    saveFile(data,path)
    pass
    

if __name__ == "__main__":
    os.sys.exit(main())