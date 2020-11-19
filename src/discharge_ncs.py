#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Apr  9 15:40:52 2020

@author: lukass

pip3.8 install xlrd==1.1.0

KM


ContrastBolusAgent 
ContrastBolusStartTime
ContrastBolusIngredientConcentration
ContrastBolusStopTime
ContrastBolusTotalDose
ContrastFlowDuration
ContrastFlowRate


"""
import time
import numpy as np
import pandas as pd


def autosize_excel_columns(worksheet, df):
  autosize_excel_columns_df(worksheet, df.index.to_frame())
  autosize_excel_columns_df(worksheet, df, offset=df.index.nlevels)

def autosize_excel_columns_df(worksheet, df, offset=0):
  for idx, col in enumerate(df):
    series = df[col]
    
    max_len = max((  series.astype(str).map(len).max(),      len(str(series.name)) )) + 1        
    max_len = min(max_len, 100)
    worksheet.set_column(idx+offset, idx+offset, max_len)
    
def df_to_excel(writer, sheetname, df):
     
    df.to_excel(writer, sheet_name = sheetname, freeze_panes = (df.columns.nlevels, df.index.nlevels))
    autosize_excel_columns(writer.sheets[sheetname], df)
    
def apply_format(worksheet, levels, df_ref, df_bool, format_highlighted):
      
    for i in range(df_ref.shape[0]):
        for j in range(df_ref.shape[1]):
            if df_bool.iloc[i,j]:
                v  = df_ref.iloc[i,j]
                try:
                    if v!=v:
                        worksheet.write_blank(i+1, j+levels, None, format_highlighted)
                    else:                                        
                        worksheet.write(i+1, j+levels, v, format_highlighted)
                except:
                    print("An exception occurred "+type(v))                    
    return

def isNaN(num):
    return num != num

def isNumber(value):
  try:
    float(value)
    return True
  except ValueError:
    return False

def discharge_ncs(df_data):
    print('Apply discharge_ncs')
    df = df_data.copy()
    for i in df.index:
        
        if i % 1000 == 0:
            print('index:', i, '/', len(df))
        
        size = 0.0 if not isNumber(df.loc[i,'ReconstructionDiameter']) else float(df.loc[i,'ReconstructionDiameter'])
        thick = 0.0 if not isNumber(df.loc[i,'SliceThickness']) else float(df.loc[i,'SliceThickness'])
        id_split = str(df.loc[i,'PatientID']).split('-')        
        site = 'P'+ id_split[0]
        
        sizeok = (size >= 320.0) 
        
        hasge = (site=='P10' or site=='P13' or site=='P29')
        if hasge: 
            thickok = (thick == 0.625)
        else:
            thickok = (thick == 1.0)
            
        islarge = (sizeok and thickok)
        
        desc =  df.loc[i,'SeriesDescription']        
        kernel =  df.loc[i,'ConvolutionKernel']
        comment = df.loc[i,'ImageComments']        

        ir = ['aidr','id', 'asir','imr']
        fbp = ['org','fbp']
        lung = ['b60','b70', 'bi57','br64','fc51','lung']
        body = ['b30','b31', 'br36','fc17','body']
        
        isir = any(x in str(desc).lower() for x in ir)
        isfbp = any(x in str(desc).lower() for x in fbp)
     
        islung1 = any(x in str(kernel).lower() for x in lung)        
        isbody1 = any(x in str(kernel).lower() for x in body)
        
        islung2 = any(x in str(desc).lower() for x in lung)        
        isbody2 = any(x in str(desc).lower() for x in body)
        
        islung3 = any(x in str(comment).lower() for x in lung)        
        isbody3 = any(x in str(comment).lower() for x in body)
        
        islung = any([islung1,islung2,islung3])
        isbody = any([isbody1,isbody2,isbody3])
        
        haskm = False if pd.isna(df.loc[i,'ContrastBolusAgent']) else True
                   
        df.loc[i,'Site'] = site
        df.loc[i,'HasGE'] = hasge
        df.loc[i,'SizeOK'] = sizeok    
        df.loc[i,'ThickOK'] = thickok
        df.loc[i,'IsLarge'] = islarge
     
        df.loc[i,'IsIR'] = isir        
        df.loc[i,'IsFBP'] = isfbp
        df.loc[i,'FoundFBPorIR'] = isfbp or isir
        
        df.loc[i,'IsLung'] = islung and islarge
        df.loc[i,'IsBody'] = isbody and islarge
        
        df.loc[i,'HasKM'] = haskm
     
        
    #per patent, count is large
    for i in df.index:
        if i % 1000 == 0:
            print('index:', i, '/', len(df))
            
        pid =  df.loc[i,'PatientID']  
        #dfpid = df[(df['PatientID']==pid) & (df['IsLarge']==True)]   
        dfpid = df[df['PatientID']==pid]
        
        ncs = dfpid['IsLarge']
        nncs = sum(ncs)
        
        fbp = dfpid['IsFBP']
        ir = dfpid['IsIR']
        
        nir = sum(ir & ncs)
        nfbp = sum(fbp & ncs)
        
        km = dfpid['HasKM']
        nokm = ~dfpid['HasKM']
        
        nirkm = sum(ir & ncs & km)
        nfbpkm = sum(fbp & ncs & km)
        
        nirnokm = sum(ir & ncs & nokm)
        nfbpnokm = sum(fbp & ncs & nokm)
    
        df.loc[i,'NrNCS'] = nncs        
        df.loc[i,'NrNCSIR'] = nir
        df.loc[i,'NrNCSFBP'] = nfbp

        df.loc[i,'NrNCSIRKM'] = nirkm
        df.loc[i,'NrNCSFBPKM'] = nfbpkm
        
        df.loc[i,'NrNCSIRNoKM'] = nirnokm
        df.loc[i,'NrNCSFBPNoKM'] = nfbpnokm
        
    df_lung = df['IsLung']
    df_body = df['IsBody']
        
    return df_lung, df_body

    
