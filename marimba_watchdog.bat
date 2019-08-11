
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
from pprint import pprint


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


# In[ ]:


##############################################################
# Find spreadsheet and wait until event is triggered
##############################################################
wb, sheets = get_spreadsheet_with(['python_output', 'verbose'])

while(True):
    try:
        uid_complete = sheets.python_output.Range('B1').Value
        uid_current  = sheets.python_output.Range('A1').Value
        uid_allowrun = sheets.python_output.Range('C1').Value
        
    except com_error:
        print('.',end='')
        
    else:
        if (uid_complete != uid_current) and uid_allowrun:
            sheets.python_output.Range('B1').Value = 0
            break
            
    time.sleep(2)
print('Start all engines')

##############################################################
# Get the main sheet's headers
##############################################################

rownr = get_id_location(sheets.python_output)[1]
header = get_header_structure(sheets.python_output, 2)

p = read_marimba_params_in_mm(sheets.python_output, rownr, header)


# In[ ]:


##############################################################
# Get the verbose sheet's headers and try to write previous line_xx, line_yy
##############################################################

v_header = get_header_structure(sheets.verbose)
v_vals   = v_header.copy()

clear_verbose_sheet(sheets.verbose)

try:
    p_prev = read_marimba_params_in_mm(sheets.python_output, rownr-1)

    v_vals.previous_offset  .setValues(
            get_xx_yy(p_prev.bar_parameters,
                      p_prev.cur_coeffs.toList(),
                      p_prev.offset_xy)
    )
    
    v_vals.previous_offset.x*=1000
    v_vals.previous_offset.z*=1000
    
    #Add a zero crossing
    d = v_vals.previous_offset
    d.x = np.r_[d.x[0], d.x, d.x[-1]]
    d.z = np.r_[0     , d.z, 0      ]
    
    
    write_vals(sheets.verbose, v_header, v_vals)

except:
    print("Ooops, can find previous block's values")


# In[ ]:


#####################################################
# Sculpture a block with the correct frequencies
#####################################################
from pprint import pprint


out = dmap(coeffs=p.prev_coeffs.toList(),
           err=np.inf)

for ii in range(30):
    if out.err < 0.5:
        break
        
    out = get_bar_shape(p.bar_parameters,
                        p.final_frequency.toArray(),
                        p.prev_raw_fem.toArray(),
                        p.prev_measured.toArray(),
                        coeffs_initial=out.coeffs)

    prntout = out.copy()
    del prntout.block.line_xx
    del prntout.block.line_yy
    
    print('\n*** Round: ', ii)
    pprint(prntout)

    #####################################################
    # Write verbose output
    #####################################################
    v_vals = v_header.copy()
    v_vals.UID.uid = p[''].uid
    
    
    block_offset = out.block.copy()
    
    (block_offset.line_xx,
     block_offset.line_yy ) = get_xx_yy(out.block, out.coeffs, p.offset_xy)
    
    v_vals.current_offset. setValues(
                              get_xx_yy(out.block, out.coeffs, p.offset_xy)
    )

    v_vals.current_final .setValues(
                            get_xx_yy(out.block, out.coeffs, (0,0))
    )
    
    #Times 1000 to mm
    v_vals.current_offset.x *= 1000
    v_vals.current_offset.z *= 1000
    v_vals.current_final.x  *= 1000
    v_vals.current_final.z  *= 1000

    for d in (v_vals.current_offset, 
              v_vals.current_final):
        
        d.x = np.r_[d.x[0], d.x, d.x[-1]]
        d.z = np.r_[0     , d.z, 0      ]

    
    #####################################################
    # What would the freq be if we offset this block?
    #####################################################
    offset_freqs_uncalib = timoshenko_beam_freqs(block_offset)
    offset_freqs_calibrated = calibrate_raw_FEM(offset_freqs_uncalib,
                                                p.prev_raw_fem.toArray(),
                                                p.prev_measured.toArray())
    #This is a absolute mess:
    v_vals.output.uid = p[''].uid
    
    for ii, line in  enumerate(zip([ii+1, '', ''],
                                   ['','',''],
                                   out.f_target,
                                   out.freqs_calib,
                                   out.freqs_uncalib,
                                   ['','',''],
                                   out.cents,
                                   ['','',''],
                                   offset_freqs_calibrated,
                                   offset_freqs_uncalib)):
    
        v_vals.output[('p0','p1','p2')[ii]] = list(line)
    
    
    write_vals(sheets.verbose, v_header, v_vals)


# In[ ]:


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


# In[ ]:


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

filename = '//Cnc/CNC/marimba FEM/%s_%.3f.txt'%(p.initials.note, p.offset_xy[1]*1000)
    
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
    sheets.verbose.Range("A14").Value = ""
    
except:
    print('No output to CNC computer')
    import tempfile
    fname = os.path.abspath(tempfile.gettempdir()+'/'+os.path.split(filename)[-1])
    with open(fname, 'w') as f:
        f.write(CNC_str)
    print('Saved to temp directory: ', fname)
    
    sheets.verbose.Range("A14").Value = "ERROR: not written to CNC"

    
sheets.python_output.Range('B1').Value = sheets.python_output.Range('A1').Value

