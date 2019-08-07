
# coding: utf-8

# In[10]:


''' >NUL  2>NUL
@echo off
cd /d %~dp0
:loop
python %0 %*
goto loop
'''

from misc import *
if is_interactive():
    get_ipython().run_line_magic('load_ext', 'autoreload')
    get_ipython().run_line_magic('autoreload', '2')
    
#https://docs.microsoft.com/en-us/previous-versions
#/office/developer/office-2003/aa174290(v=office.11)
from marimba import *
import sys
import copy

from pywintypes import com_error
import time 

import win32com
excel = win32com.client.Dispatch("Excel.Application")


# In[11]:


jupyter_name = 'marimba_watchdog'
if is_interactive():
    import os
    import subprocess
    subprocess.call(['jupyter',
                     'nbconvert',
                     '--to',
                     'script',
                     jupyter_name+'.ipynb'])
    try:
        os.remove(jupyter_name+'.bat')
    except: pass
    os.rename(jupyter_name+'.py', jupyter_name+'.bat')


# In[23]:



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
    
    
def headcopy(d):
    dcopy = ddict()
    for i,j in d.items():
        dcopy[i] = ddict(**j)
        
    return dcopy
        
def write_vals(ws, header, vals, rownr=3):
    for i in header:
        for ii, jj in header[i].items():
            
            try:
                len(vals[i][ii])
            except: pass
            else:
                if len(vals[i][ii]) == 2 and tuple(vals[i][ii]) == tuple(header[i][ii]):
                    continue
                
            if isinstance(vals[i][ii],(list,np.ndarray)):
                ln = len(vals[i][ii])    
                crange = ws.Range(
                    "%s%d:%s%d"%(num2col[header[i][ii][0]-1], rownr+0,
                                 num2col[header[i][ii][0]-1], rownr+ln-1
                                ))

                crange.Value = [[float(v)] for v in vals[i][ii][0:ln]]
                    
                    #row = 1
                    #ws.Range(ws.Cells(row,1),
                    #         ws.Cells(row+len(data_array)-1,len(data_array[0]))
                    #         ).Value = data_array
                    #print "Processing time: " + str(time.time() - start) + " seconds."
                    
                    
                #crange.Value = tuple(float(v) for v in vals[i][ii])
                #idx = -1
                #for c in crange:
                #    idx+=1
                #    c.Value = vals[i][ii][idx]
                    
            else:
                c = ws.Cells(rownr, header[i][ii][0])
                c.Value=vals[i][ii]
            
            
def clear_rows_from(ws, rownr):
    N = max(rownr, ws.UsedRange.Rows.Count)
    M = num2col[ws.UsedRange.Columns.Count-1]
    
    print("A%d:%s%d"%(rownr,M,N))
    ws.Range("A%d:%s%d"%(rownr,M,N)).ClearContents()
    


# In[18]:


wb, sheets = get_spreadsheet_with(['python_input', 'python_output', 'verbose'])

while(True):
    try:
        uid_complete = sheets.python_output.Range('B1').Value
        uid_current  = sheets.python_output.Range('A1').Value
        
    except com_error:
        print('.',end='')
        
    else:
        if uid_complete != uid_current:
            sheets.python_output.Range('B1').Value = 0
            break
            
    time.sleep(2)

print('Start all engines')

headers = get_header_structure(sheets.python_output, 2)
    
vals_in   = headcopy(headers)       
vals_prev = headcopy(headers)       
    
row = get_id_location(sheets.python_output)[1]

for i in vals_in:
    for j in vals_in[i]:
        col = vals_in[i][j][0]
        vals_in[i][j] = sheets.python_output.Cells(row,col).Value
        vals_prev[i][j] = sheets.python_output.Cells(row-1,col).Value
        


# In[24]:


n_elements=300

coeffs = [j for i,j in vals_in.prev_coeffs.items()]

final_freq = np.array([j for i,j in vals_in.final_frequency.items()])

raw_to_meas = (np.array([j for i,j in vals_in.prev_raw_fem.items()]),
               np.array([j for i,j in vals_in.prev_measured.items()]))

offset_xy=(vals_in.initials.undercut_x/1000,
           vals_in.initials.undercut_z/1000)

block = ddict(
    width = vals_in.bar_parameters.width/1000,
    depth = vals_in.bar_parameters.thickness/1000, 
    length = vals_in.bar_parameters.length/1000,
    E = vals_in.bar_parameters.youngs,
    rho = vals_in.bar_parameters.density,
    nu = vals_in.bar_parameters.youngs
    )


clear_rows_from(sheets.verbose, 3)

try:
    coeffs_prev = [j for i,j in vals_prev.cur_coeffs.items()]

    offset_xy_prev = (vals_prev.initials.undercut_x/1000,
                      vals_prev.initials.undercut_z/1000)


    block_prev = ddict(
        width = vals_prev.bar_parameters.width/1000,
        depth = vals_prev.bar_parameters.thickness/1000, 
        length = vals_prev.bar_parameters.length/1000,
        E = vals_prev.bar_parameters.youngs,
        rho = vals_prev.bar_parameters.density,
        nu = vals_prev.bar_parameters.youngs
        )

    verbose_header = get_header_structure(sheets.verbose)
    verbose_vals = headcopy(verbose_header)

    (verbose_vals.previous_offset.x,
     verbose_vals.previous_offset.z ) = get_xx_yy(block_prev, coeffs_prev, offset_xy_prev)
    
    verbose_vals.previous_offset.x *= 1000
    verbose_vals.previous_offset.z *= 1000


    write_vals(sheets.verbose, verbose_header, verbose_vals)

except TypeError:
    pass


# In[6]:


out = ddict(coeffs=coeffs,
            err=np.inf)

for i in range(30):
    if out.err < 1:
        break
        
    out = get_bar_shape(block,
                        final_freq,
                        *raw_to_meas,
                        coeffs_initial=out.coeffs)
    
    verbose_vals = headcopy(verbose_header)

    import datetime
    verbose_vals.UID.uid = vals_in[''].uid
    verbose_vals.UID.completed=False
    verbose_vals.UID.date=datetime.date.today().strftime('%d, %b %Y')

    (verbose_vals.current_offset.x,
     verbose_vals.current_offset.z ) = get_xx_yy(out.block, out.coeffs, offset_xy)

    (verbose_vals.current_final.x,
     verbose_vals.current_final.z ) = get_xx_yy(out.block, out.coeffs, (0,0))

    verbose_vals.current_offset.x *= 1000
    verbose_vals.current_offset.z *= 1000
    verbose_vals.current_final.x *= 1000
    verbose_vals.current_final.z *= 1000
    

    write_vals(sheets.verbose, verbose_header, verbose_vals)


# In[6]:


verbose_vals = headcopy(verbose_header)
verbose_vals.completed=True
write_vals(sheets.verbose, verbose_header, verbose_vals)


vals_out = headcopy(headers)


block_offset = ddict(**out.block)

(block_offset.line_xx,
 block_offset.line_yy ) = get_xx_yy(block_offset, out.coeffs, offset_xy)

offset_freqs_uncalib = timoshenko_beam_freqs(block_offset)
offset_freqs_calibrated = calibrate_raw_FEM(offset_freqs_uncalib,
                                            *raw_to_meas)

(vals_out.cur_target_calibrated_fem.f0, 
 vals_out.cur_target_calibrated_fem.f1,
 vals_out.cur_target_calibrated_fem.f2) = offset_freqs_calibrated

(vals_out.cur_raw_fem.f0,
 vals_out.cur_raw_fem.f1,
 vals_out.cur_raw_fem.f2 ) = offset_freqs_uncalib

(vals_out.cur_coeffs.c0,
 vals_out.cur_coeffs.c1,
 vals_out.cur_coeffs.c2 ) = out.coeffs

vals_out.output.cut_point = None
vals_out.output.z_centre  = None
vals_out.output.z_min     = None

row = get_id_location(sheets.python_output)[1]
write_vals(sheets.python_output, headers, vals_out, rownr=row)


# In[7]:


str_out = []

xxr, yyr = get_xx_yy(out.block, out.coeffs, offset_xy, 10000)

xx, yy = get_xx_yy_toolpath(block,
                   coeffs,
                   offset_xy=offset_xy,
                   drill_diam=vals_in.cnc.drill_diam/1000,
                   step_mm=0.4)

to_zero = block.length/2
for x, y in zip(xx, yy):
    str_out.append('X %.2f Z %.2f'%((x+to_zero)*1000, y*1000))

filename = '//Cnc/CNC/marimba FEM/%s_%.3f.txt'%(vals_in.initials.note, vals_in.initials.undercut_z)
    
infstr="""
( filename = %s )
( width = %.2f )
( length = %.2f )
( cutpoint = %.2f )
( depthmin = %.2f )
( depthmid = %.2f )
"""%(
os.path.split(filename)[-1],
out.block.width*1000,
out.block.length*1000,
(xxr[np.argmax(yyr!=yyr[0])]-xxr[0])*1000,
np.min(yyr)*1000,
yyr[int(len(yyr)/2)]*1000
)

CNC_str = open('template_cut.txt', 'r').read().format(
    info=infstr,
    length=block.length*1000,
    width=block.width*1000,
    depth=block.depth*1000,
    forwards='\n'.join(str_out),
    backwards='\n'.join(str_out[::-1])
)


try:
    with open(filename, 'w') as f:
        f.write(CNC_str)
    print('Written to "%s"'%filename.replace('/','\\'))
except:
    print('No output to CNC computer')

sheets.python_output.Range('B1').Value = sheets.python_output.Range('A1').Value

