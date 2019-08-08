import time
from datetime import datetime
import os
import pythoncom
import win32api
import win32com.client
import pylab as plt
import numpy as np
import sys

#DotMap with two monkeypatch extentions
from dotmap import DotMap as dmap
def __setValues(self, lst): 
    for i, k in enumerate(self):
        self[k] = lst[i]
dmap.setValues = __setValues 
dmap.toList    = lambda self: list(self.values())
dmap.toArray   = lambda self: np.array(list(self.values()))

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
    p = dmap()
    for i in myRange:
        if getA1(i) in donezo:
            continue

        merge_area = i.MergeArea        
        if merge_area.Value is None:
            pname = nameify(i.Value)
        else:
            pname = nameify(merge_area[0][0])

        #if getA1(i) == 'A1': #hack
        #    pname = 'UID'

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


def get_id_location(ws):
    
    cellid = int(ws.Cells(1,1).Value)
    print(cellid)

    for i in range(sys.maxsize):
        has_found=False
        for c in ws.Range('A%d:A%d'%(i*100+2, (i+1)*100+2)):

            try:
                cval = int(c.Value)
            except:
                continue

            if cval==cellid:
                has_found=True
                return getCR(c)    
        
def write_vals(ws, header, vals, rownr=3):
    for i in header:
        for ii, jj in header[i].items():
            #If "vals" is still rc index and not a value, do not write
            try:
                len(vals[i][ii])
            except: pass
            else:
                if len(vals[i][ii]) == 2 and tuple(vals[i][ii]) == tuple(header[i][ii]):
                    continue
                
            #special write for lists and arrays
            if isinstance(vals[i][ii],(list,np.ndarray)):
                ln = len(vals[i][ii])    
                crange = ws.Range(
                    "%s%d:%s%d"%(num2col[header[i][ii][0]-1], rownr+0,
                                 num2col[header[i][ii][0]-1], rownr+ln-1
                                ))

                def is_num(n):
                    try:
                        n+1
                        return True
                    except: return False
                    
                crange.Value = [[float(v) if is_num(v) else v] for v in vals[i][ii][0:ln]]
                    
            #single line write
            else:
                c = ws.Cells(rownr, header[i][ii][0])
                c.Value=vals[i][ii]
            
def clear_rows_from(ws, rownr):
    N = max(rownr, ws.UsedRange.Rows.Count)
    M = num2col[ws.UsedRange.Columns.Count-1]
    
    print("A%d:%s%d"%(rownr,M,N))
    ws.Range("A%d:%s%d"%(rownr,M,N)).ClearContents()
    
def get_row_values(ws, header, rownr):
    vals_in   = header.copy()
    
    for i in vals_in:
        for j in vals_in[i]:
            col = vals_in[i][j][0]
            vals_in[i][j] = ws.Cells(rownr, col).Value
            
    return vals_in    
    
def clear_verbose_sheet(ws):
    N = max(3, ws.UsedRange.Rows.Count)
    ws.Range("B3:J15").ClearContents()
    ws.Range("E3:J%d"%N).ClearContents()
    
    
def read_marimba_params_in_mm(ws, rownr, header=None):
    if not header:
        header = get_header_structure(ws, 2)
    
    values = get_row_values(ws, header, rownr)
    
    values.bar_parameters.depth = values.bar_parameters.pop('thickness')
    for k in ['length', 'width', 'depth']:
        values.bar_parameters[k]/=1000
        
    values.offset_xy = ( values.initials.undercut_x/1000,
                         values.initials.undercut_z/1000 )
    
    values.bar_parameters.E = values.bar_parameters.pop('youngs')
    values.bar_parameters.rho = values.bar_parameters.pop('density')
        
    return values
    