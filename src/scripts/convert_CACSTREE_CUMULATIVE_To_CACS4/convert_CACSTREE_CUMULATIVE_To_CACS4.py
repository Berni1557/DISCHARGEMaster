# -*- coding: utf-8 -*-

import sys,os
sys.path.append('H:/cloud/cloud_data/Projects/DL/Code/src')
from glob import glob
from ct.CTDataStruct import CTPatient, CTImage, CTRef

fp = 'H:/cloud/cloud_data/Projects/DL/Code/src/tmp/References_FB_segments'
files = glob(fp + '/*.nrrd')
class_dict=dict({1: 1,
                 2: 1,
                 23: 1,
                 24: 1,
                 25: 1,
                 26: 1,
                 27: 1,
                 28: 1,
                 29: 1,
                 30: 1,
                 31: 1,
                 32: 1,
                 33: 1,
                 34: 1,
                 35: 1,
                 8: 5,
                 9: 5,
                 10: 5,
                 11: 5,
                 12: 5,
                 13: 2,
                 14: 2,
                 15: 2,
                 16: 2,
                 17: 2,
                 18: 3,
                 19: 3,
                 20: 3,
                 21: 3,
                 22: 3,
                 3: 4,
                 4: 4,
                 5: 4,
                 6: 4,
                 7: 4})

for i,f in enumerate(files):
    print('i: ', i)
    ref = CTRef()
    ref.load(f)
    arr_org = ref.ref()
    arr = arr_org.copy()
    for c0 in class_dict:
        arr[arr_org==c0]=class_dict[c0]
    ref.setRef(arr)
    ref.save(f)
