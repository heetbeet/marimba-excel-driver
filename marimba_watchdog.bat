#!/usr/bin/env python
# coding: utf-8

# In[1]:


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


# In[2]:


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


# In[3]:


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
    


# In[4]:


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

rownr = get_id_location(sheets.python_output)[1]
header = get_header_structure(sheets.python_output, 2)

p = read_marimba_params_in_mm(sheets.python_output, rownr, header)


# In[5]:


v_header = get_header_structure(sheets.verbose)
v_vals   = v_header.copy()

clear_rows_from(sheets.verbose, 3)

try:
    p_prev = read_marimba_params_in_mm(sheets.python_output, rownr-1)

    v_vals.previous_offset  .setValues(
            get_xx_yy(p_prev.block,
                      p_prev.cur_coeffs,
                      p_prev.offset_xy)
    )
    
    v_vals.previous_offset.x*=1000
    v_vals.previous_offset.y*=1000

except TypeError:
    print("Ooops, can find previous block's values")
    
print(v_header)


# In[6]:


#####################################################
# Sculpture a block wiht the correct frequencies
#####################################################
from pprint import pprint


out = dmap(coeffs=p.prev_coeffs.toList(),
           err=np.inf)

for i in range(30):
    if out.err < 1:
        break
        
    out = get_bar_shape(p.bar_parameters,
                        p.final_frequency.toArray(),
                        p.prev_raw_fem.toArray(),
                        p.prev_measured.toArray(),
                        coeffs_initial=out.coeffs)

    prntout = out.copy()
    del prntout.block.line_xx
    del prntout.block.line_yy
    
    print('\n*** Round: ',i)
    pprint(prntout)

    v_vals = v_header.copy()
    v_vals.UID.uid = p[''].uid
    

    v_vals.current_offset. setValues(
                            get_xx_yy(out.block,
                                      out.coeffs, 
                                      p.offset_xy))
        

    v_vals.current_final .setValues(
                            get_xx_yy(out.block, out.coeffs, (0,0))
    )

    
    v_vals.current_offset.x *= 1000
    v_vals.current_offset.z *= 1000
    v_vals.current_final.x  *= 1000
    v_vals.current_final.z  *= 1000

    
    #####################################################
    # What would the freq be if we offset this block?
    #####################################################
    block_offset = out.block.copy()

    (block_offset.line_xx,
     block_offset.line_yy ) = get_xx_yy(out.block, out.coeffs, p.offset_xy)

    offset_freqs_uncalib = timoshenko_beam_freqs(block_offset)
    offset_freqs_calibrated = calibrate_raw_FEM(offset_freqs_uncalib,
                                                p.prev_raw_fem.toArray(),
                                                p.prev_measured.toArray())
    
    v_vals.output.label =["",
                           "f final target",
                           "f final current calib.",
                           "f final current raw",
                           "",
                           "cent err",
                           "",
                           "f undercut calib.",
                           "f undercut raw"]
    
    for ii, line in  enumerate(zip("   ",
                                   out.f_target,
                                   out.freqs_calib,
                                   out.freqs_uncalib,
                                   "   ",
                                   out.cents,
                                   "   ",
                                   offset_freqs_calibrated,
                                   offset_freqs_uncalib)):
    
        v_vals.output[('p0','p1','p2')[ii]] = list(line)
    
    
    
    write_vals(sheets.verbose, v_header, v_vals)


# In[ ]:





# In[7]:


####################################
# Write this to the spreadsheet
####################################
vals_out = header.copy()

vals_out.cur_target_calibrated_fem .setValues(offset_freqs_calibrated)
vals_out.cur_raw_fem               .setValues(offset_freqs_uncalib)
vals_out.cur_coeffs                .setValues(out.coeffs)
 
xxr, yyr = get_xx_yy(out.block,
                     out.coeffs,
                     p.offset_xy,
                     10000)

vals_out.output.cut_point = (xxr[np.argmax(yyr!=yyr[0])]-xxr[0])*1000
vals_out.output.z_centre  = yyr[int(len(yyr)/2)]*1000
vals_out.output.z_min     = np.min(yyr)*1000

write_vals(sheets.python_output, header, vals_out, rownr=rownr)


# In[13]:


##############################
# Write this to the CNC machine
###############################

str_out = []

xx, yy = get_xx_yy_toolpath(
                   out.block,
                   out.coeffs,
                   offset_xy=p.offset_xy,
                   drill_diam=p.cnc.drill_diam/1000,
                   step_mm=0.4)

to_zero = out.block.length/2
for x, y in zip(xx, yy):
    str_out.append('X %.2f Z %.2f'%((x+to_zero)*1000, y*1000))

filename = '//Cnc/CNC/marimba FEM/%s_%.3f.txt'%(p.initials.note, p.offset_xy[1])
    
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
vals_out.output.cut_point,
vals_out.output.z_centre,
vals_out.output.z_min,
)

CNC_str = open('template_cut.txt', 'r').read().format(
    info=infstr,
    length=out.block.length*1000,
    width=out.block.width*1000,
    depth=out.block.depth*1000,
    forwards='\n'.join(str_out),
    backwards='\n'.join(str_out[::-1])
)


try:
    with open(filename, 'w') as f:
        f.write(CNC_str)
    print('Written to "%s"'%filename.replace('/','\\'))
except:
    print('No output to CNC computer')
    import tempfile
    fname = os.path.abspath(tempfile.gettempdir()+'/'+os.path.split(filename)[-1])
    with open(fname, 'w') as f:
        f.write(CNC_str)
    print('Saved to temp directory: ', fname)
        

sheets.python_output.Range('B1').Value = sheets.python_output.Range('A1').Value

