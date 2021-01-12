# -*- coding: utf-8 -*-

import sys, os
sys.path.append('H:/cloud/cloud_data/Projects/DL/Code/src')
sys.path.append('H:/cloud/cloud_data/Projects/DL/Code/src/ct')
import pandas as pd
import ntpath
import datetime
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.styles import Font, Color, Border, Side
from openpyxl.styles import Protection
from openpyxl.styles import PatternFill
from glob import glob
from shutil import copyfile
import numpy as np
from collections import defaultdict
from openpyxl.utils import get_column_letter
from CTDataStruct import CTPatient
import keyboard
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting import Rule
from settings import initSettings, saveSettings, loadSettings, fillSettingsTags
from classification import createRFClassification, initRFClassification, classifieRFClassification
from filterTenStepsGuide import filter_CACS_10StepsGuide, filter_CACS, filter_NCS, filterReconstruction, filter_CTA, filer10StepsGuide, filterReconstructionRF
from CTDataStruct import CTPatient, CTImage
import SimpleITK as sitk

def createCACS(settings, name, folderpath, createPreview, createDatasetFromPreview, NumSamples=None):
    
    # Create dataset folder
    folderpath_data = os.path.join(folderpath, name)
    os.makedirs(folderpath_data, exist_ok=True)
    folderpath_preview = os.path.join(folderpath_data, 'preview')
    os.makedirs(folderpath_preview, exist_ok=True)
    folderpath_dataset = os.path.join(folderpath_data, 'Images')
    os.makedirs(folderpath_dataset, exist_ok=True)   
    filepath_preview = os.path.join(folderpath_preview, 'preview.xlsx')
    filepath_dataset = os.path.join(folderpath_dataset, name +'.xlsx')
    cols = ['ID_CACS', 'PatientID', 'StudyInstanceUID', 'SeriesInstanceUID', 'SeriesNumber', 'Count', 'NumberOfFrames',
            'KHK', 'RECO', 'SliceThickness', 'ReconstructionDiameter', 'ConvolutionKernel', 'CACSSelection', 
            'StudyDate', 'ITT', 'Comment', 'KVP']
    cols_master = ['PatientID', 'StudyInstanceUID', 'SeriesInstanceUID', 'SeriesNumber', 'Count', 'NumberOfFrames',
            'RECO', 'SliceThickness', 'ReconstructionDiameter', 'ConvolutionKernel', 'StudyDate', 'ITT', 'Comment']
    cols_first = ['ID_CACS','CACSSelection', 'PatientID', 'SeriesNumber', 'StudyInstanceUID', 'SeriesInstanceUID']
    
    if createPreview:
        
        # Read master
        df_master = pd.read_excel(settings['filepath_master'], sheet_name='MASTER_01092020')
        df_preview = pd.DataFrame(columns=cols)
        df_preview[cols_master] = df_master[cols_master]
        df_preview['KHK'] = 'UNDEFINED'
        df_preview['KVP'] = 'UNDEFINED'
        df_preview['CACSSelection'] = (df_master['RFCLabel']=='CACS')*1
        
        # Create preview excel
        df_preview.reset_index(drop=True, inplace=True)
        constrain = list(df_master['RFCLabel']=='CACS')
        k=0
        ID_CACS = [-1 for i in range(len(constrain))]
        for i in range(len(constrain)):
            if constrain[i]==True:
                ID_CACS[i]="{:04n}".format(k)
                k = k + 1
        df_preview['ID_CACS'] = ID_CACS
        cols_new = cols_first + [x for x in cols if x not in cols_first]
        df_preview = df_preview[cols_new]
        df_preview.reset_index(inplace=True, drop=True)
        df_preview.to_excel(filepath_preview)
        
        # Create preview mhd
        filepath_preview_mhd = os.path.join(folderpath_preview, 'preview.mhd')
        k_max = k-1
        image_preview = np.zeros((k_max,512,512), np.int16)
        
        if NumSamples is None:
            NumSamples = len(df_preview)
        
        for index, row in df_preview[0:NumSamples].iterrows():
            if int(row['ID_CACS'])>-1:
                try:
                    if index % 100==0:
                        print('Index:', index)
                    print('Index:', index)
                    patient = CTPatient(row['StudyInstanceUID'], row['PatientID'])
                    series = patient.loadSeries(settings['folderpath_discharge'], row['SeriesInstanceUID'], None)
                    image = series.image.image()
                    if image.shape[1]==512:
                        SliceNum = int(np.round(image.shape[0]*0.7))
                        image_preview[int(row['ID_CACS']),:,:] = image[SliceNum,:,:]
                    else:
                        print('Image size is not 512x512')
                        print('SeriesInstanceUID', row['SeriesInstanceUID'])
                except:
                    print('Coud not open image:', row['SeriesInstanceUID'])
    
        image_preview_mhd = CTImage()
        image_preview_mhd.setImage(image_preview)
        image_preview_mhd.save(filepath_preview_mhd)
    
    # Create dataset
    if createDatasetFromPreview:
        df_preview = pd.read_excel(filepath_preview)
        df_cacs = df_preview[df_preview['ManualCorrection']==1]
        df_cacs.to_excel(filepath_dataset)
        for index, row in df_preview[0:10].iterrows():
            if index % 100==0:
                print('Index:', index)
            patient=CTPatient(row['StudyInstanceUID'], row['PatientID'])
            series = patient.loadSeries(settings['folderpath_discharge'], row['SeriesInstanceUID'], None)
            image = series.image
            filepath_image = os.path.join(folderpath_dataset, row['PatientID'] + '_' + row['SeriesInstanceUID'] + '.mhd')
            image.save(filepath_image)
    
    
    
def createDataset():
    """ Create master file
    """
    
    # Load settings
    filepath_settings = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/data/settings.json'
    settings=initSettings()
    saveSettings(settings, filepath_settings)
    settings = fillSettingsTags(loadSettings(filepath_settings))
    
    # Create CACS preview
    name = 'CACS_20210801'
    folderpath = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/datasets'
    createPreview = True
    createDatasetFromPreview = False
    createCACS(settings, name, folderpath, createPreview, createDatasetFromPreview, NumSamples=None)
    
    
    
    
    
    