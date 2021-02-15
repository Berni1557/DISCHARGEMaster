# -*- coding: utf-8 -*-

import sys, os
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import numpy as np
from collections import defaultdict
from glob import glob
import pydicom
#from scanner_map import searchKey, CertifiedManufacturerModelNameCTDict, CertifiedManufacturerCTDict, TrueManufacturerModelNameCTDict, TrueManufacturerCTDict
#from scanner_map import ScannerType, CertifiedManufacturerModelNameICADict, CertifiedManufacturerICADict, TrueManufacturerModelNameICADict, TrueManufacturerICADict


filepath_scanner_study = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/src/scanner/data/scanner_correspondence_V06.xlsx'
filepath_scanner_pre = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/src/scanner/data/CT_scanner_11012021_VW_pre.xlsx'
filepath_scanner_post = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/src/scanner/data/CT_scanner_11012021_VW_BF.xlsx'

df_scanner_study = pd.read_excel(filepath_scanner_study, 'linear', index_col=None)
df_scanner = pd.read_excel(filepath_scanner_pre, 'Tabelle1', index_col=None)
df_row = pd.read_excel(filepath_scanner_pre, 'Tabelle2', index_col=0)

# Update scanner
StudyDateKeys = ['StudyDate_00', 'StudyDate_01', 'StudyDate_02', 'StudyDate_03', 'StudyDate_04', 'StudyDate_05', 'StudyDate_06']
ModalityKey = ['Modality_00', 'Modality_01', 'Modality_02', 'Modality_03', 'Modality_04', 'Modality_5', 'Modality_06']
TrueManufacturerKey = ['TrueManufacturer_00', 'TrueManufacturer_01', 'TrueManufacturer_02', 'TrueManufacturer_03', 'TrueManufacturer_04', 'TrueManufacturer_05', 'TrueManufacturer_06']
TrueManufacturerModelNameKey = ['TrueManufacturerModelName_00', 'TrueManufacturerModelName_01', 'TrueManufacturerModelName_02', 'TrueManufacturerModelName_03', 'TrueManufacturerModelName_04', 'TrueManufacturerModelName_05', 'TrueManufacturerModelName_06']

for index, row in df_scanner.iterrows():
    patient = row['Patient identifier']
    print('patient', patient)
    #if patient=='01-BER-0014':
    #    sys.exit()
    date = row['Date of CT DICOM']
    study = df_scanner_study[df_scanner_study['PatientID']==patient]
    study.reset_index(inplace=True)
    for studyKey, modKey, manKey, manModKey in zip(StudyDateKeys, ModalityKey, TrueManufacturerKey, TrueManufacturerModelNameKey):
        if (study[studyKey][0] == date) and ('CT' in study[modKey][0]):
            df_scanner.loc[index,'TrueManufacturer_00'] = study[manKey][0]
            df_scanner.loc[index,'TrueManufacturerModelName_00'] = study[manModKey][0]
        
    #sys.exit()
    

# Update rows
df_scanner_new = df_scanner.copy()
df_scanner_new['Rows'] = ''
modelList = list(df_row['Unnamed: 1'])
rowlist = list(df_row['Unnamed: 2'])
row_dict = dict(zip(modelList, rowlist))
row_dict[np.nan] = ''

for index, row in df_scanner.iterrows():
    model = row['TrueManufacturerModelName_00']
    if model == 'SOMATOM Definition':
        model = 'Definition DS dual-source'
    r = row_dict[model]
    df_scanner_new.loc[index,'Rows']=r
    df_scanner_new.loc[index,'TrueManufacturerModelName_00']=model
    
    
sheet_name='Tabelle1'
df_scanner_new.to_excel(filepath_scanner_post, sheet_name=sheet_name)
# sheet_name='Tabelle2'
# df_row.to_excel(filepath_scanner_post, sheet_name=sheet_name)

def scanner_check():
    # Check scanner
    folderpath_discharge = 'G:/discharge'
    for index, row in df_scanner.iterrows():
        # if index==960:
        #     sys.exit()
        model = row['TrueManufacturerModelName_00']
        ITT = row['ITT (CT=1; ICA=0; 2=exluded']
        StudyInstanceUID = row['StudyInstanceUID_00']
        if ITT==1 and pd.isnull(model) and not pd.isnull(StudyInstanceUID):
            print(index)
            # if index==50:
            #     sys.exit()
            
            if not pd.isnull(StudyInstanceUID):
                folders = glob(os.path.join(folderpath_discharge, StudyInstanceUID + '/*'))
                for f in folders:
                    alldcm = glob(f + '/*.dcm')
                    ds = pydicom.dcmread(alldcm[0], force = False, defer_size = 256, specific_tags = ['Modality'], stop_before_pixels = True)
                    Modality = ds.data_element('Modality').value
                    if 'CT' in Modality:
                        ds = pydicom.dcmread(alldcm[0], force = False, defer_size = 256, specific_tags = ['Modality', 'Manufacturer', 'ManufacturerModelName'], stop_before_pixels = True)
                        ManufacturerModelName = ds.data_element('ManufacturerModelName').value
                        Manufacturer = ds.data_element('Manufacturer').value
                        print('Index:', index, Manufacturer, ManufacturerModelName)
                        break

