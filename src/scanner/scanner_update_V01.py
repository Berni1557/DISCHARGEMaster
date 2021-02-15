# -*- coding: utf-8 -*-

import sys, os
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import numpy as np
from collections import defaultdict
from scanner_map import searchKey, CertifiedManufacturerModelNameCTDict, CertifiedManufacturerCTDict, TrueManufacturerModelNameCTDict, TrueManufacturerCTDict
from scanner_map import ScannerType, CertifiedManufacturerModelNameICADict, CertifiedManufacturerICADict, TrueManufacturerModelNameICADict, TrueManufacturerICADict

filepath_scanner_pre = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/src/scanner/data/CT_scanner_11012021_VW_pre.xlsx'
filepath_scanner_post = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/src/scanner/data/CT_scanner_11012021_VW_BF.xlsx'

df_scanner = pd.read_excel(filepath_scanner_pre, 'Tabelle1', index_col=None)
df_row = pd.read_excel(filepath_scanner_pre, 'Tabelle2', index_col=0)


for index, row in df_scanner.iterrows():
    patient = row['Patient identifier']
    print('patient', patient)
    #if patient=='01-BER-0014':
    #    sys.exit()
    date = row['Date of CT DICOM']
    StudyInstanceUID = row['StudyInstanceUID_00']
    
    study.reset_index(inplace=True)
    for studyKey, modKey, manKey, manModKey in zip(StudyDateKeys, ModalityKey, TrueManufacturerKey, TrueManufacturerModelNameKey):
        if (study[studyKey][0] == date) and ('CT' in study[modKey][0]):
            df_scanner.loc[index,'TrueManufacturer_00'] = study[manKey][0]
            df_scanner.loc[index,'TrueManufacturerModelName_00'] = study[manModKey][0]
            print('set')
            


