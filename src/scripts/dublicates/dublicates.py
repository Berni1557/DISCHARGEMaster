# -*- coding: utf-8 -*-
import sys, os
sys.path.append('H:/cloud/cloud_data/Projects/DL/Code/src')
sys.path.append('H:/cloud/cloud_data/Projects/DL/Code/src/ct')
import pandas as pd
from tqdm import tqdm

filepath_hist = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/data/discharge_master/discharge_master_01092020/discharge_sources_01092020/discharge_hist_01092020.pkl'
filepath_dublicates = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/data/discharge_master/discharge_master_01092020/tmp/discharge_dublicates.pkl'

# Load data
df_hist = pd.read_pickle(filepath_hist)
columns_hist = [str(i) for i in range(100)]

# Filter  empty arrays
df_col = df_hist[columns_hist]
idx_empty = df_col['0']>-1
df_filt0 = df_col[idx_empty]
df_hist0 = df_hist[idx_empty]

# Filter dublicates
idx_dub = df_filt0.duplicated(keep=False)
df_hist1 = df_hist0[idx_dub]

# Extract dublicates
pbar = tqdm(total=len(df_hist1))
pbar.set_description("Extract dublicates")
df_dublicates = pd.DataFrame(columns=['SeriesInstanceUID', 'dublicates'])
for index0, row0 in df_hist1.iterrows():
    pbar.update(1)
    c0 = list(row0[columns])
    dub_list=[]
    for index1, row1 in df_hist1.iterrows():
        if row0['SeriesInstanceUID']!=row1['SeriesInstanceUID']:
            c1 = list(row1[columns])
            if c0==c1:
                dub_list.append(row1['SeriesInstanceUID'])
    if len(dub_list)>0:
        df_dublicates = df_dublicates.append(dict({'SeriesInstanceUID': row0['SeriesInstanceUID'], 'dublicates': dub_list}), ignore_index=True)
pbar.close()
df_dublicates.to_pickle(filepath_dublicates)
            
            
            