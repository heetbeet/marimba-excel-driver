from misc import dmap, ddict
import six

import warnings
import numpy as np
import pylab as plt
from scipy import optimize
from scipy import interpolate

from collections import OrderedDict

#Find where a signal goes from positive to negative axis
def find_up_to_down_idx(arr):
    return np.where((arr[:-1]>=0)*(arr[1:]<0))[0]

def hertz_to_cents(x):
    return 1200*np.log2(x)

def cents_to_hertz(x):
    return 2**(x/1200)

def get_notes_dict():
    C1 =  0.5*65.703693103900300
    namelist = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B']
    notes_dict = OrderedDict()
    for scale in range(1,11):
        for name in namelist:
            if notes_dict:
                notes_dict[name+str(scale)] = notes_dict[next(reversed(notes_dict))]*(2**(1./12))
            else:
                notes_dict['C1']=C1
    return notes_dict

notes_dict = get_notes_dict()

def freq_to_note(freq):
    idx = np.argmin(np.abs(np.array(list(notes_dict.values()))-freq))
    note = list(notes_dict.keys())[idx]
    deltacents = hertz_to_cents(notes_dict[note])-hertz_to_cents(freq)
    return note, deltacents

def note_to_freq(note):
    return notes_dict[note.upper()]


def get_xx_yy(block,
              coeffs,
              offset_xy=None,
              n_elements=None):
    
    if n_elements is None:
        n_elements = 300
    if offset_xy is None:
        offset_xy = (0,0)
    
    n_points = n_elements + 1

    delx = block.length/n_points

    def f(line_xx):
        return (coeffs[0]*(np.abs(line_xx))**3 +
                coeffs[1]*(np.abs(line_xx))**2 +
                coeffs[2])
    
    line_xx = np.linspace(-block.length/2, block.length/2,  n_points)
    line_yy = f(line_xx)

    if offset_xy[0]!=0 or offset_xy[1]!=0:
        ox, oy = offset_xy
        
        line_yy = line_yy+oy
        line_yy = np.max(np.r_[[line_yy],
                               [f(line_xx+ox)+oy],
                               [f(line_xx-ox)+oy] ], axis=0)
        
        
    line_yy[line_yy>block.depth] = block.depth
    
    return line_xx, line_yy

def get_xx_yy_toolpath(block,
                       coeffs,
                       offset_xy=None,
                       drill_diam=0.01,
                       step_mm=0.5):
    
        rad = drill_diam/2
        steplength = step_mm/1000
        xx, yy = get_xx_yy(block, coeffs, offset_xy, n_elements=100000)

        rad_in_samples = int(np.round(rad/(xx[1]-xx[0])))
        if rad_in_samples > 0:
            yy_new = np.max(np.r_[[yy[rad_in_samples:-rad_in_samples]],
                                  [yy[rad_in_samples*2:]],
                                  [yy[:-rad_in_samples*2]]], axis=0)
            xx_new = xx[rad_in_samples:-rad_in_samples]
        else:
            yy_new = yy;
            xx_new = xx;

        try:
            str_idx = find_up_to_down_idx(yy_new-block.depth+1e-8)[0]
        except IndexError: 
            str_idx = 0

        str_point = xx_new[str_idx]
        end_point = xx_new[-str_idx]

        xx_out = np.linspace(str_point, end_point, 
                            (end_point-str_point)/steplength)
        
        yy_out = interpolate.interp1d(xx_new, yy_new)(xx_out)

        return xx_out, yy_out
    
#%%
def get_shaving_positions(beam, nfreqs=3):
    try:
        beam_copy = ddict(**beam.toDict())
    except:
        beam_copy = ddict(**beam)
        
    cut_depth = np.array([beam.width*0.01, beam.width*0.02, beam.width*0.03, beam.width*0.02, beam.width*0.01])*0.1
    n = 2
    
    freqs_real = timoshenko_beam_freqs(beam)
    
    freqs = []
    for i in range(n,len(beam.line_yy)-n):
        beam_copy.line_yy = np.array(beam.line_yy)
        beam_copy.line_yy[i-n:i+n+1] -=cut_depth
        freqs.append(timoshenko_beam_freqs(beam_copy))
        

    freqs = np.array(([freqs[0]]*(n+1))+freqs+([freqs[-1]]*(n+1)))

    return hertz_to_cents(freqs) - hertz_to_cents(freqs_real)


#%%
def timoshenko_beam_freqs(beam, n_freqs=3):
    #print(beam)
    from scipy import sparse
    nu = 0.3
    line_yy = (beam.line_yy +
               np.random.normal(scale=1e-16, size=beam.line_yy.shape))
    
    yy_midpnts = (line_yy[1:] + line_yy[:-1])/2
    
    I_yy    = beam.width*(yy_midpnts**3)/12
    area_yy = beam.width*yy_midpnts
    N_elem  = len(yy_midpnts)
    l       = (beam.length)/N_elem
    phi_yy  = 24*(6/5)*(1+nu)*I_yy/(area_yy*l**2)
    
    ME_yy = [
        (beam.rho*area*l/420)*
        np.array(
            [[ 156  ,  22*l   ,  54   , -13*l   ], 
             [ 22*l ,  4*l**2 ,  13*l , -3*l**2 ], 
             [ 54   ,  13*l   ,  156  , -22*l   ], 
             [-13*l , -3*l**2 , -22*l ,  4*l**2 ]])
        for area in area_yy
    ]
    
    KE_yy = [
        (beam.E*I/((1+phi)*l**3))*
        np.array(
            [[ 12  ,  6*l         , -12  ,  6*l        ],
             [ 6*l , (4+phi)*l**2 , -6*l , (2-phi)*l**2],
             [-12  , -6*l         ,  12  , -6*l        ],
             [ 6*l , (2-phi)*l**2 , -6*l , (4+phi)*l**2]])
        
        for phi, I, in zip(phi_yy, I_yy)
    ]
    
    M = np.zeros((2*N_elem+2, 2*N_elem+2))
    K = np.zeros((2*N_elem+2, 2*N_elem+2))

    for ii in range(0,M.shape[0]-2,2):
        i = int(ii//2)
        M[ii:ii+4,ii:ii+4] += ME_yy[i]
        K[ii:ii+4,ii:ii+4] += KE_yy[i]

    #for a free-free beam, no contraints is set, but the first two
    #eigenvalues are discarded
    omega_pow2, _ = sparse.linalg.eigs(sparse.csc_matrix(K),
                                       n_freqs+2,
                                       sparse.csc_matrix(M),
                                       sigma=0)
        
    f = np.sort(np.sqrt(np.abs(omega_pow2.real))/(2*np.pi))[2:]
    return f


#%%
def calibrate_raw_FEM(raw_freqs,
                      samples_raw_freqs,
                      samples_measured):
    N = len(raw_freqs)
    samples_raw_freqs = samples_raw_freqs[:N]
    samples_measured = samples_measured[:N]
    
    ratios = np.array(samples_measured)/np.array(samples_raw_freqs)
    return raw_freqs*ratios

    
def uncalibrate_FEM(freqs,
                    samples_raw_freqs,
                    samples_measured):
    N = len(freqs)
    samples_raw_freqs = samples_raw_freqs[:N]
    samples_measured = samples_measured[:N]
    
    ratios = np.array(samples_measured)/np.array(samples_raw_freqs)

    return freqs/ratios


def get_bar_shape(block,
                  f_target,
                  
                  samples_raw,
                  samples_measured,

                  coeffs_initial=[6, 0.3000, 0.0030],
                  n_elements=None,
                  **kwargs):

    set_kwargs = dict(xtol=0.001,
                      ftol=0.5,
                      maxfun=1000, 
                      disp=False)
    for i in kwargs: set_kwargs[i] = kwargs[i]

    block_copy = block.copy()

    #hack to keep the coefficient in a decent metric
    #currently (c0>>c1>>c2) and it causes issues
    coeffs_scaling = np.array([1,10,1000])
    coeffs_ = np.array(coeffs_initial)*coeffs_scaling
    
    #hack to make f0 more sensitive than f1 and f2
    err_ratio=[1, 0.5, 0.5]
    err_ratio = np.array(err_ratio)/np.sum(err_ratio)

    #the return parameters
    keeps = dmap()
    keeps.f_target = f_target
    
    def minimize_f(coeffs_):
        coeffs=coeffs_/coeffs_scaling


        (block_copy.line_xx,
         block_copy.line_yy) = get_xx_yy(block_copy,
                                         coeffs,
                                         offset_xy=None,
                                         n_elements=n_elements)

        freqs_uncalib = timoshenko_beam_freqs(block_copy, n_freqs=3)
        freqs_calib = calibrate_raw_FEM(freqs_uncalib,
                                        samples_raw,
                                        samples_measured)

        delta = (hertz_to_cents(f_target)-
                 hertz_to_cents(freqs_calib))

        keeps.block = block_copy
        keeps.coeffs = coeffs
        keeps.freqs_calib = freqs_calib
        keeps.freqs_uncalib = freqs_uncalib
        keeps.cents = delta
        keeps.err = np.sum(np.abs(delta)*err_ratio)
        
        return keeps.err

    optimize.fmin(minimize_f, coeffs_, **set_kwargs)
    
    return keeps