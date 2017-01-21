import os
import Tkinter
import tkFileDialog
from xlrd import open_workbook
import xlwt

def addStyle(path):
    pass

def saveFile(information, path = ''):
    path = path.split('.')
    path[-2] += '_NEW'
    path = '.'.join(path)
    path = tkFileDialog.asksaveasfilename(initialfile = path)
    wbk = xlwt.Workbook(style_compression=2) 
    sheet = wbk.add_sheet('Data')
    for row, Rdata in enumerate(information[1:],1):
        for col, data in enumerate(Rdata):
            if (row - 1) % 18 == 0 :
                if col == 3:
                    sheet.write(row,col,data,style = xlwt.easyxf("borders: top double, left thin"))
                else:
                    sheet.write(row,col,data,style = xlwt.easyxf("borders: top double"))
            elif col > 3 and (row - 1) % 3 == 0:
                sheet.write(row,col,data,style = xlwt.easyxf("borders: top thin"))
            elif col == 3 and (row - 1) % 3 == 0:
                sheet.write(row,col,data,style = xlwt.easyxf("borders: top thin, left thin"))
            elif col == 3:
                sheet.write(row,col,data,style = xlwt.easyxf("borders: left thin"))
            else:
                sheet.write(row,col,data)
    
    for col, Fdata in enumerate(information[0]):
        sheet.write(0,col,Fdata,style = xlwt.easyxf("font: bold on; borders: bottom thick"))
    
    for c in range(6):
        sheet.col(c).width = 18*256

    wbk.save(path)

    addStyle(path)

    os.startfile(path)

def getDataFromFile(in_path):
    wb = open_workbook(in_path)
    for s in wb.sheets():
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
    title = title[:6]
    title = map(lambda x: x.replace('_1','') ,title)
    title = map(lambda x: x.replace('1','') ,title)
    for rowData in data:
        mainInfo = rowData[:3]
        rowData = rowData[3:]
        for idOffset in xrange(0,42,7):
            currentID = [rowData[idOffset]]
            idData = rowData[idOffset+1:idOffset+7]
            for scrapOffset in xrange(0,6,2):
                scrapInfo = idData[scrapOffset:scrapOffset+2]
                newData.append(mainInfo + currentID + scrapInfo)
    newData.insert(0,title)
    return newData

def main():

    Tkinter.Tk().withdraw() # Close the root window
    path = tkFileDialog.askopenfilename(filetypes = [('Excel file','*.xls')])
    data = getDataFromFile(path)
    title = data[0]
    cleanUp(data)#, map(lambda x: title.index(x),['ID','Oprator','StartTime_Run','Shift']))
    print "Working..."
    data = reformate(data[1:],title)
    print "Saving..."
    saveFile(data,path)
    pass
    

if __name__ == "__main__":
    os.sys.exit(main())