# -*- coding: utf-8 -*-
import numpy as np
import pandas as pd
from discharge_ncs import discharge_ncs
from computeCTA import computeCTA

def isNaN(num):
    return num != num

def filter_CACS(df_data):
    print('Apply filter_CACS_')
    df_CACS = pd.DataFrame(columns=['CACS'])
    for index, row in df_data.iterrows():
        if index % 1000 == 0:
            print('index:', index, '/', len(df_data))
        criteria1 = row['ReconstructionDiameter'] <= 300
        criteria2 = (row['SliceThickness']==3.0) or (row['SliceThickness']==2.5 and row['Site'] in ['P10', 'P13', 'P29'])
        criteria3 = row['Modality'] == 'CT'
        criteria4 = isNaN(row['ContrastBolusAgent'])
        criteria5 = row['Count']>=30 and row['Count']<=90
        result = criteria1 and criteria2 and criteria3 and criteria4 and criteria5
        df_CACS = df_CACS.append({'CACS': result}, ignore_index=True)
    return df_CACS

def filter_NCS(df_data):
    df = pd.DataFrame()
    df_lung, df_body = discharge_ncs(df_data)
    df['NCS_CACS'] = df_lung
    df['NCS_CTA'] = df_body
    return df

def filterReconstruction(df_data, settings):
    print('Apply filterReconstruction123')
    ir_description = ['aidr','id', 'asir','imr']
    fbp_description = ['org','fbp']
    ir_kernel = ['I20f','I26f','I30f', 'I31f','I50f', 'I70f']
    fbp_kernel = []
    df_reco = pd.DataFrame(columns=['CACSExtended'])
    for index, row in df_data.iterrows():
        if index % 1000 == 0:
            print('index:', index, '/', len(df_data))
        row['Modality'] == 'CT'
        desc =  row['SeriesDescription']  
        kernel =  row['ConvolutionKernel'] 
        if isNaN(kernel):kernel=''
       
        # Check Cernel
        isir = any(x.lower() in str(desc).lower() for x in ir_description) or any(x.lower() in str(kernel).lower() for x in ir_kernel)
        isfbp = any(x.lower() in str(desc).lower() for x in fbp_description) or any(x.lower() in str(kernel).lower() for x in fbp_kernel)
        if isfbp:
            reco = settings['recoClasses'][0]
        elif isir:
            reco = settings['recoClasses'][1]
        else:
            reco = settings['recoClasses'][2]
        df_reco = df_reco.append({'RECO': reco}, ignore_index=True)
    return df_reco



def filter_CTA(settings):
    df_cta = computeCTA(settings)
    df = pd.DataFrame()
    df['phase'] = df_cta['CTA_phase']
    df['arteries'] = df_cta['CTA_arteries']
    df['source'] = df_cta['CTA_source']
    df['CTA'] = df_cta['CTA']
    df.fillna(value=np.nan, inplace=True)   
    return df    