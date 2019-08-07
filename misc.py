import time
from datetime import datetime
import os
import pythoncom
import win32api
import win32com.client
import pylab as plt

alph = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
num2col = [i for i in alph] + [i+j for i in alph for j in alph]


class ddict(dict):
    def __init__(self, **kwds):
        self.update(kwds)
        self.__dict__ = self

def getA1(cell):
    return num2col[cell.Column-1]+str(cell.Row)

def getCR(cell):
    return cell.Column, cell.Row

def get_range_locations(cellrange):
    return [getA1(i) for i in cellrange]
    
def is_interactive():
    import __main__ as main
    return not hasattr(main, '__file__')

def spread_iterator():
    for moniker in pythoncom.GetRunningObjectTable():
        try:
            # Workbook implements IOleWindow so only consider objects implementing that
            window = moniker.BindToObject(pythoncom.CreateBindCtx(0), None, pythoncom.IID_IOleWindow)
            disp = window.QueryInterface(pythoncom.IID_IDispatch)


            # Get a win32com Dispatch object from the PyIDispatch object as it's
            # easier to work with.
            book = win32com.client.Dispatch(disp)

        except pythoncom.com_error:
            # Skip any objects we're not interested in
            continue

        try:
            book.Sheets(1) #Object is a book with sheets
        except:
            continue
            
        bookname = moniker.GetDisplayName(pythoncom.CreateBindCtx(0), None)

        yield bookname, book

def get_spreadsheet_with(sheets=[]):
    for bookname, book in spread_iterator():
        print('Test workbook: ', bookname)

        sheetdict = ddict()
        
        did_find = True
        for sheet in sheets:
            sheetdict[sheet] = [i for i in book.Sheets if i.Name.lower() == sheet]
            
            if len(sheetdict[sheet]) == 0:
                did_find=False
                break
            
            sheetdict[sheet] = sheetdict[sheet][0]
                
        if did_find:
            print('We have -->', bookname)
            return book, sheetdict
        
    raise ValueError("Couldn't find spreadsheet.")
        
def get_spreadsheet_by_name(spreadname):
    for bookname, book in spread_iterator():
        print('Test workbook: ', bookname)
        fname = os.path.split(bookname)[-1].lower()

        fexts = ['.xls', '.csv', '.txt']
        for fext in fexts:
            if fext in fname:
                fname = fext.join(fname.split(fext)[:-1])
        if fname == spreadname.lower():
            return book

        
def get_header_structure(ws, firstrow=1):
    from itertools import chain
    def nameify(s):
        return str(s).strip().lower().replace(' ','_').replace('-','_')

    N = ws.UsedRange.Rows.Count
    M = ws.UsedRange.Columns.Count

    myRange = ws.Range('A%d:%s%d'%(firstrow, num2col[M-1], firstrow))

    donezo = set()
    p = ddict()
    for i in myRange:
        if getA1(i) in donezo:
            continue

        merge_area = i.MergeArea        
        if merge_area.Value is None:
            pname = nameify(i.Value)
        else:
            pname = nameify(merge_area[0][0])

        if getA1(i) == 'A1': #hack
            pname = 'UID'

        p[pname] = ddict()

        for j in chain(i, merge_area):
            donezo.add(getA1(j))

            ppname = nameify(ws.Cells(j.Row+1, j.Column).Value)

            p[pname][ppname] = getCR(j)
            
    return p


def clear_images(wb):
    for n, shape in enumerate(ws.Shapes):
        if shape.Name.startswith("Picture"):
            shape.delete()
        
def insert_image(wb, impath, location=(1,1)):

    M = ws.UsedRange.Columns.Count

    left=ws.Cells(*location).Left
    top=ws.Cells(*location).Top

    img = plt.imread(imgpath)


    ws.Shapes.AddPicture(os.path.abspath(imgpath),
                         LinkToFile=False,
                         SaveWithDocument=True,
                         Left=left,
                         Top=top,
                         Width=img.shape[1],
                         Height=img.shape[0])



"""
for n, shape in enumerate(ws.Shapes):
    if shape.Name.startswith('TextBox'):
        shape.delete()

M = ws.UsedRange.Columns.Count

left=ws.Cells(1,M+1).Left
top=ws.Cells(1,1).Top
right=ws.Cells(1,M+7).Left
bottom=ws.Cells(10,1).Top

tb = ws.Shapes.AddTextbox(1,  left      , top,
                              right-left, bottom-top)
tb.TextFrame2.TextRange.Characters.Text = 'This is a great big test.'
tb.TextFrame2.TextRange.Characters.Font.Name = "Consolas"
tb.TextFrame2.TextRange.Characters.Font.Size = 10
"""