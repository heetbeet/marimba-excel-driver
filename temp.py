# -*- coding: utf-8 -*-
"""
Created on Thu Oct 17 16:24:14 2019

@author: Laptop
"""
from misc import ddict
import marimba

#%%|
block= ddict(**{
        'E': 15000000000.0,
        'depth': 0.02,
        'length': 0.4689,
        'node_x': 85.80580620469274,
        'rho': 700.0,
        'width': 0.067})

cents = np.array([0.00382612, 0.00094661, 0.39387347])
coeffs = np.array([ 9.04361628e+00, -3.81259087e-01,  6.72109633e-03])
err = 0.10061808108866899
f_target = np.array([ 110.5   ,  442.    , 1113.7295])
freqs_calib = np.array([ 110.49975579,  441.99975832, 1113.47614396])
freqs_uncalib = np.array([ 128.30773257,  518.11324962, 1372.62930603])

#%%
line_xx, line_yy = marimba.get_xx_yy(block, coeffs)
plt.plot(0,0)
plt.plot(line_xx, line_yy)

block2 = ddict(**block)
block2.line_yy = line_yy

marimba.timoshenko_beam_freqs_fast(block2.line_yy, block2.width, block2.length, block2.E)


#%%
import marimba
import pylab as plt
bla = marimba.get_shaving_positions(block2)
plt.plot(bla[:,0], 'r')
plt.plot(bla[:,1], 'g')
plt.plot(bla[:,2], 'b')