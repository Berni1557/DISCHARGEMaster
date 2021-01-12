import os, sys
import numpy as np
import pandas as pd
import time
import pydicom
from glob import glob  

def computeSliceSpacing(alldcm):
    
    try:
        if len(alldcm)>1:
            ds0 = pydicom.dcmread(alldcm[0], force = False, defer_size = 256, specific_tags = ['SliceLocation'], stop_before_pixels = True)
            value0 = float(ds0.data_element('SliceLocation').value)
            ds1 = pydicom.dcmread(alldcm[1], force = False, defer_size = 256, specific_tags = ['SliceLocation'], stop_before_pixels = True)
            value1 = float(ds1.data_element('SliceLocation').value)
            SliceSpacing = abs(value1-value0)
        else:
            SliceSpacing = -1.0
            
    except Exception as why:
        SliceSpacing = -1.0
    
    return SliceSpacing

def extractDICOMTags(settings, NumSamples=None):

    root = settings['folderpath_discharge']
    fout = settings['filepath_dicom']
    specific_tags = settings['dicom_tags']
    cols_first=[]

    study_uids = os.listdir(root)
    
    df = pd.DataFrame(columns=specific_tags)
    i = 0
    
    if NumSamples is None:
        NumSamples = len(study_uids)
    
    specific_tags_dcm = specific_tags.copy()
    if 'Site' in specific_tags_dcm: specific_tags_dcm.remove('Site')
    if 'Count' in specific_tags_dcm: specific_tags_dcm.remove('Count')
    if 'SliceSpacing' in specific_tags_dcm: specific_tags_dcm.remove('SliceSpacing')
    
    for istudy, study_uid in enumerate(study_uids[0:NumSamples]):              
        
        print(istudy, study_uid)
        if not os.path.exists(os.path.join(root, study_uid)): continue

        series_uids = os.listdir(os.path.join(root, study_uid))
        
        if True:
            for series_uid in series_uids:
                
                #print('series_uid', series_uid)
                
                path_series = os.path.join(root, study_uid, series_uid)   
                alldcm = glob(path_series + '/*.dcm')
                
                # Check if multi slice or single slice format
                #print('x0')
                ds = pydicom.dcmread(alldcm[0], force = False, defer_size = 256, specific_tags = ['NumberOfFrames'], stop_before_pixels = True)
                #print('x1')
                try:        
                    NumberOfFrames = ds.data_element('NumberOfFrames').value
                    MultiSlice = True                              
                except: 
                    NumberOfFrames=''
                    MultiSlice = False
                    #print('except0:')
    
                if MultiSlice:
                    for dcm in alldcm[0:1]:
                        try:
                            ds = pydicom.dcmread(dcm, force = False, defer_size = 256, specific_tags = specific_tags_dcm, stop_before_pixels = True)
                        except Exception as why:          
                            #print('Exception:', why)
                            print('StudyInstanceUID:', study_uid)
                            print('SeriesInstanceUID:', series_uid)
                            df.loc[i,'StudyInstanceUID'] = study_uid
                            df.loc[i,'SeriesInstanceUID'] = series_uid
                            continue
                        if 'Site' in specific_tags:
                            df.loc[i,'Site'] = 'P'+ str(ds.PatientID).split('-')[0]
                        if 'Count' in specific_tags:
                            df.loc[i,'Count'] = len(alldcm)
                        if 'SliceSpacing' in specific_tags:
                            df.loc[i,'SliceSpacing'] = -1
                        
                        for tag in specific_tags:
                            try:        
                                data_element = ds.data_element(tag)                                
                            except:   
                                #print('except1')
                                continue                
                            if data_element is None:
                                continue
                            df.loc[i,tag] = str(data_element.value)
                else:
                    try:
                        ds = pydicom.dcmread(alldcm[0], force = False, defer_size = 256, specific_tags = specific_tags_dcm, stop_before_pixels = True)
                    except Exception as why: 
                        #print('except3')
                        #print('Exception:', why)
                        print('StudyInstanceUID:', study_uid)
                        print('SeriesInstanceUID:', series_uid)
                        df.loc[i,'StudyInstanceUID'] = study_uid
                        df.loc[i,'SeriesInstanceUID'] = series_uid
                        continue
                    if 'Site' in specific_tags:
                        df.loc[i,'Site'] = 'P'+ str(ds.PatientID).split('-')[0]
                    if 'Count' in specific_tags:
                        df.loc[i,'Count'] = len(alldcm)
                    if 'SliceSpacing' in specific_tags:
                        #print('test02')
                        SliceSpacing = computeSliceSpacing(alldcm)
                        df.loc[i,'SliceSpacing'] = SliceSpacing
                    
                    #print('test01')
                    # Extract tags if exist
                    for tag in specific_tags:
                        #print('found0', tag)
                        try:        
                            data_element = ds.data_element(tag)                                
                        except: 
                            #print('found1', tag)
                            continue                
                        if data_element is None:
                            continue
                        df.loc[i,tag] = str(data_element.value)
                i += 1 #series

    # Reorder datafame
    cols = df.columns.tolist()
    cols_new = cols_first + [x for x in cols if x not in cols_first]
    df = df[cols_new]
    
    # Convert strings to numbers in df
    tags_str = ['ReconstructionDiameter', 'Count', 'SeriesNumber', 'SeriesNumber', 'NumberOfFrames', 'Rows',
                'Columns', 'InstanceNumber', 'SliceThickness', 'SliceThickness', 'ReconstructionDiameter']
    df.replace(to_replace=['None'], value=np.nan, inplace=True)
    for tag in tags_str:
        df[tag] = pd.to_numeric(df[tag])
        
    df.sort_values('PatientID', inplace=True)
    df.reset_index(drop=True, inplace=True)        

    writer = pd.ExcelWriter(fout)            
    df.to_excel(writer, sheet_name = "linear")
    writer.save()

