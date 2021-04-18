# -*- coding: utf-8 -*-
import os, sys
import pandas as pd
from glob import glob
import pydicom
import numpy as np
from settings import initSettings, saveSettings, loadSettings, fillSettingsTags

# Load settings
filepath_settings = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/data/settings.json'
settings=initSettings()
saveSettings(settings, filepath_settings)
settings = fillSettingsTags(loadSettings(filepath_settings))
    
# Read patients
root = settings['folderpath_discharge']
fip_cacs_table = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/src/scripts/scanlength/data/cacs_table_1.xlsx'
fip_scanlength = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/src/scripts/scanlength/data/scan_length.xlsx'
df_cacs = pd.read_excel(fip_cacs_table)
patients = list(df_cacs['Unnamed: 0'].unique())

# Read master
fip_master = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/src/scripts/scanlength/data/discharge_master_01092020.xlsx'
df_master = pd.read_excel(fip_master, sheet_name='MASTER_01092020')

# Extract CTA and CACS for patient
df_scans = pd.DataFrame()
for patient in patients:
    print('PatientID', patient)
    df_patient = df_master[df_master['PatientID']==patient]
    if len(df_patient)>0:
        # Extract CACS
        df_patient_cacs = df_patient[df_patient['RFCLabel']=='CACS']
        df_patient_cacs.reset_index(inplace=True)
        df_patient_cta = df_patient[df_patient['RFCLabel']=='CTA']
        df_patient_cta.reset_index(inplace=True)
        if len(df_patient_cacs)>0:
            StudyInstanceUID_cacs = df_patient_cacs.loc[0,'StudyInstanceUID']
            SeriesInstanceUID_cacs = df_patient_cacs.loc[0,'SeriesInstanceUID']
        else:
            StudyInstanceUID_cacs = ''
            SeriesInstanceUID_cacs = ''
            
        if len(df_patient_cta)>0:
            StudyInstanceUID_cta = df_patient_cta.loc[0,'StudyInstanceUID']
            SeriesInstanceUID_cta = df_patient_cta.loc[0,'SeriesInstanceUID']
        else:
            StudyInstanceUID_cta = ''
            SeriesInstanceUID_cta = ''
            
        row=dict({'PatientID': patient,
                  'StudyInstanceUID_cacs': StudyInstanceUID_cacs,
                  'SeriesInstanceUID_cacs': SeriesInstanceUID_cacs,
                  'StudyInstanceUID_cta': StudyInstanceUID_cta,
                  'SeriesInstanceUID_cta': SeriesInstanceUID_cta})
                  
        df_scans = df_scans.append(row, ignore_index=True)


# Extract scan length   
df_scans['ScanLength_cacs'] = ''   
df_scans['ScanLength_cta'] = ''   
for index, row in df_scans.iterrows():
    print('index', index)
    if index>72:
        StudyInstanceUID_cacs = row['StudyInstanceUID_cacs']
        SeriesInstanceUID_cacs = row['SeriesInstanceUID_cacs']
        if StudyInstanceUID_cacs is not '' and SeriesInstanceUID_cacs is not '':
            path_series = os.path.join(root, StudyInstanceUID_cacs, SeriesInstanceUID_cacs)   
            alldcm = glob(path_series + '/*.dcm')
            # Check if multi slice or single slice format
            ds = pydicom.dcmread(alldcm[0], force = False, defer_size = 256, specific_tags = ['NumberOfFrames'], stop_before_pixels = True)
            try:        
                NumberOfFrames = ds.data_element('NumberOfFrames').value
                MultiSlice = True                              
            except: 
                NumberOfFrames=''
                MultiSlice = False
            if MultiSlice:
                sys.exit() 
            else:
                try:
                    llist=[]
                    for dcm in alldcm:
                        ds = pydicom.dcmread(dcm, force = False, defer_size = 256, specific_tags = ['SliceLocation'], stop_before_pixels = True)
                        l = float(ds.data_element('SliceLocation').value)
                        llist.append(l)
                    ScanLength = max(llist)-min(llist)
                    df_scans.loc[index,'ScanLength_cacs'] = ScanLength
                except Exception as why:
                    print('Exception')
                    continue
        StudyInstanceUID_cta = row['StudyInstanceUID_cta']
        SeriesInstanceUID_cta = row['SeriesInstanceUID_cta']
        if StudyInstanceUID_cta is not '' and SeriesInstanceUID_cta is not '':
            path_series = os.path.join(root, StudyInstanceUID_cta, SeriesInstanceUID_cta)   
            alldcm = glob(path_series + '/*.dcm')
            # Check if multi slice or single slice format
            ds = pydicom.dcmread(alldcm[0], force = False, defer_size = 256, specific_tags = ['NumberOfFrames'], stop_before_pixels = True)
            try:        
                NumberOfFrames = ds.data_element('NumberOfFrames').value
                MultiSlice = True                              
            except: 
                NumberOfFrames=''
                MultiSlice = False
            if MultiSlice:
                ds = pydicom.dcmread(alldcm[0], force = False, defer_size = 256, specific_tags = ['SliceLocation'], stop_before_pixels = True)
                v=ds.data_element('SliceLocation').value
            else:
                try:
                    llist=[]
                    for dcm in alldcm:
                        ds = pydicom.dcmread(dcm, force = False, defer_size = 256, specific_tags = ['SliceLocation'], stop_before_pixels = True)
                        l = float(ds.data_element('SliceLocation').value)
                        llist.append(l)
                    ScanLength = max(llist)-min(llist)
                    df_scans.loc[index,'ScanLength_cta'] = ScanLength
                except Exception as why:
                    print('Exception')
                    continue            

# Update df_cacs
df_save = pd.merge(df_cacs, df_scans, left_on='Unnamed: 0', right_on='PatientID')
df_save = df_save.drop(['PatientID'], axis=1)
writer = pd.ExcelWriter(fip_scanlength)            
df_save.to_excel(writer, sheet_name = "linear")
writer.save()



##########################################################
import pydicom
dcm = 'G:/discharge/1.2.392.200036.9116.2.6.1.3268.2047366984.1447034187.883735/1.2.392.200036.9116.2.6.1.3268.2047366984.1447034922.125669/1.2.392.200036.9116.2.6.1.3268.2047366984.1447034922.124088.dcm'
ds = pydicom.dcmread(dcm, force = False, defer_size = 256, specific_tags = ['TablePosition'], stop_before_pixels = True)
ds = pydicom.dcmread(dcm, force = False, defer_size = 256, stop_before_pixels = True)
v=ds.data_element('StudyDate').value



