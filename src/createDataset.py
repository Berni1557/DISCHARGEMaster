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
from CTDataStruct import CTPatient, CTImage, CTRef
import SimpleITK as sitk
import matplotlib.pyplot as plt
from glob import glob
   
    
def splitFilePath(filepath):
    """ Split filepath into folderpath, filename and file extension
    
    :param filepath: Filepath
    :type filepath: str
    """
    folderpath, _ = ntpath.split(filepath)
    head, file_extension = os.path.splitext(filepath)
    folderpath, filename = ntpath.split(head)
    return folderpath, filename, file_extension

def checkRefereencesAL():
    fp = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/datasets/CACS_20210801_XA/References/'
    files = glob(fp + '*-label.nrrd')
    filenameList=[]
    ratioList=[]
    for i,fip in enumerate(files):
        print(i)
        _, filename, _ = splitFilePath(fip)
        ref = CTRef()
        ref.load(fip)
        N=ref.ref().shape[0]*ref.ref().shape[1]*ref.ref().shape[2]
        Nr=(ref.ref()!=0).sum()
        ratio = Nr / N
        filenameList.append(filename)
        ratioList.append(ratio)


def createCACS(settings, name, folderpath, createPreview, createDatasetFromPreview, NumSamples=None):
    
    # Create dataset folder
    folderpath_data = os.path.join(folderpath, name)
    os.makedirs(folderpath_data, exist_ok=True)
    folderpath_preview = os.path.join(folderpath_data, 'preview')
    os.makedirs(folderpath_preview, exist_ok=True)
    folderpath_dataset = os.path.join(folderpath_data, 'Images')
    os.makedirs(folderpath_dataset, exist_ok=True)   
    filepath_preview = os.path.join(folderpath_preview, 'preview.xlsx')
    filepath_preview_refine = os.path.join(folderpath_preview, 'preview_refine.xlsx')
    filepath_dataset = os.path.join(folderpath_dataset, name +'.xlsx')
    cols = ['ID_CACS', 'PatientID', 'StudyInstanceUID', 'SeriesInstanceUID', 'SeriesNumber', 'Count', 'NumberOfFrames',
            'KHK', 'RECO', 'SliceThickness', 'ReconstructionDiameter', 'ConvolutionKernel', 'CACSSelection', 
            'StudyDate', 'ITT', 'Comment', 'KVP']
    cols_master = ['PatientID', 'StudyInstanceUID', 'SeriesInstanceUID', 'SeriesNumber', 'Count', 'NumberOfFrames',
            'RECO', 'SliceThickness', 'ReconstructionDiameter', 'ConvolutionKernel', 'StudyDate', 'ITT', 'Comment']
    cols_first = ['ID_CACS','CACSSelection', 'PatientID', 'SeriesNumber', 'StudyInstanceUID', 'SeriesInstanceUID']
    
    if createPreview:
        
        # Read master
        df_master = pd.read_excel(settings['filepath_master_preview'], sheet_name='MASTER_01092020')
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
        df_cacs = pd.read_excel(filepath_preview_refine)
        #df_cacs = df_preview[df_preview['CACSSelection']==1]
        df_cacs.to_excel(filepath_dataset)
        for index, row in df_cacs.iterrows():
            if index % 100==0:
                print('Index:', index)
            if row['CACSSelection']==1:
                patient = CTPatient(row['StudyInstanceUID'], row['PatientID'])
                series = patient.loadSeries(settings['folderpath_discharge'], row['SeriesInstanceUID'], None)
                image = series.image
                filepath_image = os.path.join(folderpath_dataset, row['PatientID'] + '_' + row['SeriesInstanceUID'] + '.mhd')
                image.save(filepath_image)
            
    
def renameReferences(settings, folderpath):
  
    # Rename label
    folderpath = 'H:/tmp/CACS/References'
    files = glob(folderpath + '/*-label.nrrd')
    df_cacs = pd.read_excel('H:/tmp/CACS_20210801.xlsx')
    for file in files:
        folderpath, filename, file_extension = splitFilePath(file)
        SeriesInstanceUID = filename.split('-')[0]
        df_tmp = df_cacs[df_cacs['SeriesInstanceUID']==SeriesInstanceUID]
        for index, row in df_tmp.iterrows():
            if  row['SeriesInstanceUID'] == SeriesInstanceUID:
                print(SeriesInstanceUID)
                file_rename = row['PatientID'] + '_' + row['SeriesInstanceUID'] + '-label.nrrd'
                filepath_rename = os.path.join(folderpath, file_rename)                
                os.rename(file, filepath_rename)
                
    # Rename label-lesion
    folderpath = 'H:/tmp/CACS/References'
    files = glob(folderpath + '/*-label-lesion.nrrd')
    df_cacs = pd.read_excel('H:/tmp/CACS_20210801.xlsx')
    for file in files:
        folderpath, filename, file_extension = splitFilePath(file)
        SeriesInstanceUID = filename.split('-')[0]
        df_tmp = df_cacs[df_cacs['SeriesInstanceUID']==SeriesInstanceUID]
        for index, row in df_tmp.iterrows():
            if  row['SeriesInstanceUID'] == SeriesInstanceUID:
                print(SeriesInstanceUID)
                file_rename = row['PatientID'] + '_' + row['SeriesInstanceUID'] + '-label-lesion.nrrd'
                filepath_rename = os.path.join(folderpath, file_rename)                
                os.rename(file, filepath_rename)
    
    # Rename label_pred
    folderpath = 'H:/tmp/CACS/References'
    files = glob(folderpath + '/*-label_pred.nrrd')
    df_cacs = pd.read_excel('H:/tmp/CACS_20210801.xlsx')
    for file in files:
        folderpath, filename, file_extension = splitFilePath(file)
        SeriesInstanceUID = filename.split('-')[0]
        df_tmp = df_cacs[df_cacs['SeriesInstanceUID']==SeriesInstanceUID]
        for index, row in df_tmp.iterrows():
            if  row['SeriesInstanceUID'] == SeriesInstanceUID:
                print(SeriesInstanceUID)
                file_rename = row['PatientID'] + '_' + row['SeriesInstanceUID'] + '-label_pred.nrrd'
                filepath_rename = os.path.join(folderpath, file_rename)                
                os.rename(file, filepath_rename)
    
    # Rename label_pred
    folderpath = 'H:/tmp/CACS/References'
    files = glob(folderpath + '/*-label-lesion_pred.nrrd')
    df_cacs = pd.read_excel('H:/tmp/CACS_20210801.xlsx')
    for file in files:
        folderpath, filename, file_extension = splitFilePath(file)
        SeriesInstanceUID = filename.split('-')[0]
        df_tmp = df_cacs[df_cacs['SeriesInstanceUID']==SeriesInstanceUID]
        for index, row in df_tmp.iterrows():
            if  row['SeriesInstanceUID'] == SeriesInstanceUID:
                print(SeriesInstanceUID)
                file_rename = row['PatientID'] + '_' + row['SeriesInstanceUID'] + '-label-lesion_pred.nrrd'
                filepath_rename = os.path.join(folderpath, file_rename)                
                os.rename(file, filepath_rename)
                
    # name = 'CACS_20210801'
    # folderpath_dataset = os.path.join(folderpath_data, 'Images')
    # filepath_dataset = os.path.join(folderpath_dataset, name +'.xlsx')
    
    # files = glob(folderpath + '/*' + ext)
    
    # df_cacs = pd.read_excel(filepath_dataset)
    
    # for file in files:
    #     folderpath, filename, file_extension = splitFilePath(file)
    #     for index, row in df_cacs.iterrows():
    #         if  row['SeriesInstanceUID'] == filename:
    #             file_rename = row['PatientID'] + '_' + row['SeriesInstanceUID']
    #             filepath_rename = os.path.join(folderpath, file_rename + file_extension)
    #             #sys.exit()
    #             #patient = CTPatient(row['StudyInstanceUID'], row['PatientID'])
    #             #series = patient.loadSeries(settings['folderpath_discharge'], row['SeriesInstanceUID'], None)
    #             #image = series.image
    #             image = CTImage()
    #             image.load(file)
    #             image.save(filepath_rename)
    #             #os.remove(file)
    #             #os.rename(file, filepath_rename)
    #             print('Renaming:', file)

    
def createDataset():
    """ Create master file
    """
    
    # Load settings
    filepath_settings = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/data/settings.json'
    settings=initSettings()
    saveSettings(settings, filepath_settings)
    settings = fillSettingsTags(loadSettings(filepath_settings))
    settings['filepath_master_preview'] = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/datasets/CACS_20210801/preview/discharge_master_01092020_preview.xlsx'
    
    # Create CACS preview
    name = 'CACS_20210801'
    folderpath = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/datasets'
    createPreview = True
    createDatasetFromPreview = False
    createCACS(settings, name, folderpath, createPreview, createDatasetFromPreview, NumSamples=None)
    
    # Create CACS dataset
    name = 'CACS_20210801'
    folderpath = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/datasets'
    createPreview = False
    createDatasetFromPreview = True
    createCACS(settings, name, folderpath, createPreview, createDatasetFromPreview, NumSamples=None)
    
    # Rename files
    ext='.mhd'
    folderpath = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/datasets/CACS_20210801/tmp/Images'
    renameReferences(settings, folderpath, ext=ext)
    
    
    # #H:/cloud/cloud_data/Projects/DISCHARGEMaster/datasets/CACS_20210801/preview_refine/preview_V01.xlsx
    
    # Rename label
    folderpath = '//192.168.1.150/cloud_data/Projects/tmp/ref'
    files = glob(folderpath + '/*-label.nrrd')
    df_cacs = pd.read_excel('H:/tmp/CACS_20210801.xlsx')
    for file in files:
        folderpath, filename, file_extension = splitFilePath(file)
        SeriesInstanceUID = filename.split('-')[0]
        df_tmp = df_cacs[df_cacs['SeriesInstanceUID']==SeriesInstanceUID]
        for index, row in df_tmp.iterrows():
            if  row['SeriesInstanceUID'] == SeriesInstanceUID:
                print(SeriesInstanceUID)
                file_rename = row['PatientID'] + '_' + row['SeriesInstanceUID'] + '-label.nrrd'
                filepath_rename = os.path.join(folderpath, file_rename)                
                os.rename(file, filepath_rename)
    
    
    # patient = CTPatient('1.2.392.200036.9116.2.6.1.3268.2051314220.1494313218.277559', '22-GLA-0015')
    # series = patient.loadSeries(settings['folderpath_discharge'], '1.2.392.200036.9116.2.6.1.3268.2051314220.1494313218.430652', None)
    # image = series.image
    # image.save('H:/cloud/cloud_data/Projects/DISCHARGEMaster/tmp/test.mhd')
    
    # im = image.image_sitk
    
    # sitk.WriteImage(im, 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/tmp/test.mhd')
    
    
    # image=CTImage()
    # image.load('G:/discharge/1.2.392.200036.9116.2.6.1.3268.2051314220.1494313218.277559/1.2.392.200036.9116.2.6.1.3268.2051314220.1494313218.430652')    
    # image.save('H:/cloud/cloud_data/Projects/DISCHARGEMaster/tmp/test.mhd')
    
    
    # arr = image.image()
    # imageS=CTImage()
    # imageS.setImage(arr)
    # imageS.copyInformationFrom(image)
    # imageS.save('H:/cloud/cloud_data/Projects/DISCHARGEMaster/tmp/test.mhd')
    
