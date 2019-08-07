
# coding: utf-8

# In[ ]:


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


# In[ ]:


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


wb, sheets = get_spreadsheet_with(['python_input', 'python_output', 'verbose'])
#get_header_structure(sheets.python_input)


# In[ ]:


coeffs=[6,0.3000,0.0030]
n_elements=300
    
raw_to_meas = ([100      , 1000,      10000],
               [90.716097, 889.13575, 8566.6892])

block = ddict(
    width=0.0600, 
    depth=0.0192, 
    length=0.3000,
    E=1.25e+10,
    rho=650,
    nu = 0.3
    )

out = ddict(coeffs=coeffs,
            err=np.inf)

for i in range(10):
    print(out)
    if out.err < 1:
        break
        
    out = get_bar_shape(block,
                        [120,  500, 1200],
                        *raw_to_meas,
                        coeffs_initial=out.coeffs)
    print(out.cents)
    print(out.freqs_calib)
    
    plt.figure(figsize=[12,4])
    plt.plot(out.block.line_xx*1000, out.block.line_yy*1000)
    plt.show()

