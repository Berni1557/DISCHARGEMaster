"""
Created on Thu Sep 26 12:28:35 2019

@author: lukass

pip install pydicom --proxy=http://proxy.charite.de:8080
 
PANDA

https://towardsdatascience.com/replacing-excel-with-python-30aa060d35e
  
  
"""

#from __future__ import print_function
#https://pydicom.github.io/pydicom/dev/auto_examples/input_output/plot_read_dicom_directory.html
import os, sys
import numpy as np
import pandas as pd
import time
import pydicom
from glob import glob

def autosize_excel_columns(worksheet, df):
  autosize_excel_columns_df(worksheet, df.index.to_frame())
  autosize_excel_columns_df(worksheet, df, offset=df.index.nlevels)

def autosize_excel_columns_df(worksheet, df, offset=0):
  for idx, col in enumerate(df):
    series = df[col]
    
    #import sys
    #reload(sys)  # Reload does the trick!
    #sys.setdefaultencoding('UTF8')
    
    max_len = max((  series.astype(str).map(len).max(),      len(str(series.name)) )) + 1        
    max_len = min(max_len, 100)
    worksheet.set_column(idx+offset, idx+offset, max_len)
    
def df_to_excel(writer, sheetname, df):
     
    df.to_excel(writer, sheet_name = sheetname, freeze_panes = (df.columns.nlevels, df.index.nlevels))
    autosize_excel_columns(writer.sheets[sheetname], df)
    

def start_log(p):
    
    old_stdout = sys.stdout
    log_file = open("message.log","w")    
    sys.stdout = log_file
    print ("this will be written to message.log")
    return log_file
    
def end_log(p):
    sys.stdout = old_stdout    
    log_file.close()
       
'''

https://www.dicomlibrary.com/dicom/dicom-tags/
https://pydicom.github.io/pydicom/stable/base_element.html
https://github.com/pydicom/pydicom/blob/master/examples/input_output/plot_read_dicom_directory.py
https://pydicom.github.io/pydicom/stable/ref_guide.html
'''
def extract_tags(p):

    print(p)
    patients = list()
    
    y = find_dicom_files(directory = p, pattern = "*.dcm", directory_exclude_pattern = ".*", recursive = True)
     
    print (len(y))
    deferSize = 16383  # 128**2-1             deferSize = None
    force=True
    
    n = len(y)
    df = pd.DataFrame(index=range(n), columns=('PatientID', 'StudyInstanceUID', 'SeriesInstanceUID','SOPInstanceUID'))
   
    df['Modality']=np.nan    
    df['Study Description']=np.nan   
    df['Series Description']=np.nan 
    df['Image Count']=0
    
    s={}
    
    for i,x in enumerate(y):
        print( i)#,x 
     
        
        try:
            #stop_before_pixels=False
            dcm = pydicom.read_file(x, deferSize, force=force)
        except pydicom.filereader.InvalidDicomError:
            continue  # skip non-dicom file
        except Exception as why:
          
            print('Warning:', why)
            continue

        try:
            PatientID = dcm.PatientID
            StudyInstanceUID = dcm.StudyInstanceUID
            SeriesInstanceUID = dcm.SeriesInstanceUID         
            SOPInstanceUID = dcm.SOPInstanceUID
            
            print( PatientID, StudyInstanceUID, SeriesInstanceUID, SOPInstanceUID)
         
            df.loc[i,'PatientID'] = PatientID
            df.loc[i,'StudyInstanceUID'] = StudyInstanceUID
            df.loc[i,'SeriesInstanceUID'] = SeriesInstanceUID
            df.loc[i,'SOPInstanceUID'] = SOPInstanceUID
            df.loc[i,'Modality'] = dcm.Modality
            
            if SeriesInstanceUID in s: 
                s[SeriesInstanceUID] += 1 
            else: 
                s[SeriesInstanceUID] = 1

            #df[df['SeriesInstanceUID']==SeriesInstanceUID].loc[i,'Image Count'] += 1
            #df.loc[df[df['SeriesInstanceUID']==SeriesInstanceUID].index,'Image Count']=+1
            #l = df[df['SeriesInstanceUID']==SeriesInstanceUID].index
            #print l
            #df.loc[l,'Image Count']+=1
            
            
            df.loc[i,'Series Description'] = dcm[0x08103E].value
            df.loc[i,'Study Description'] = dcm[0x081030].value
            
        except AttributeError:
            continue  # some other kind of dicom file
            
    if 0:
 #Index.unique(self, level=None)[source]¶
        SeriesInstanceUID='1.3.12.2.1107.5.4.5.35449.30000017113015391426500000005'
        print (df[df['SeriesInstanceUID']==SeriesInstanceUID])
       
        l = df[df['SeriesInstanceUID']==SeriesInstanceUID].index
        print (l)
        df.loc[l,'Image Count']=2
        
        print (df[df['SeriesInstanceUID']==SeriesInstanceUID])
        
        return
   
    for SeriesInstanceUID in s: 
        l = df[df['SeriesInstanceUID']==SeriesInstanceUID].index
        df.loc[l,'Image Count']=s[SeriesInstanceUID]
        
            
    
    df2 = df.set_index(['PatientID','StudyInstanceUID','SeriesInstanceUID'])
    #df2.drop('SOPInstanceUID', axis=1, inplace=True)
        
    fout = 'discharge_tags.xlsx'

    writer = pd.ExcelWriter(fout)
    
    df_to_excel(writer, "summary", df)    
    df_to_excel(writer, "ordered", df2)    
      
    #workbook  = writer.book
    #sheetname="summary"
    #df.to_excel(writer, sheet_name=sheetname, freeze_panes=(df.columns.nlevels, df.index.nlevels))
    #worksheet = writer.sheets[sheetname]
    #autosize_excel_columns(worksheet, df)    
    writer.save()
    

def extract_series(p):
    df = pd.DataFrame()    
    d1 = os.listdir(p)
    i = 0
    for d in d1:
        a = os.path.join(p, d)
        d2 = os.listdir(a)
        
        #print(p, d, d2)
        for j in d2:
            df.loc[i,'StudyInstanceUID'] = d
            df.loc[i,'SeriesInstanceUID'] = j
            i += 1
            print(p, d, j)
        
    
    #df = pd.DataFrame(index=range(n), columns=('PatientID', 'StudyInstanceUID', 'SeriesInstanceUID','SOPInstanceUID'))   
    #df['Modality']=np.nan    
    
    print (df.head())
    
    writer = pd.ExcelWriter('series.xlsx')    
    df_to_excel(writer, "summary", df)    
    writer.save()


def extract_all_tags(root, fout):
    
    #fout = 'tags_xa.xlsx'
    #fout = 'tags_ct.xlsx'
    #fout = 'discharge_tags.xlsx'
    deferSize = 16383  # 128**2-1             deferSize = None
    force = True
    
    exclude_tags =['PixelData','ReferencedPerformedProcedureStepSequence', 'DerivationCodeSequence','ReferencedImageSequence',
              'CTDIPhantomTypeCodeSequence','RequestAttributesSequence','SourceImageSequence','ReferencedStudySequence',
              'ProcedureCodeSequence', 'RequestedProcedureCodeSequence','ConceptNameCodeSequence','ContentSequence','ContentTemplateSequence',
              'CurrentRequestedProcedureEvidenceSequence','DimensionIndexSequence','DimensionOrganizationSequence',
              'PerFrameFunctionalGroupsSequence','ReferencedSeriesSequence','SegmentSequence',
              'SharedFunctionalGroupsSequence','DataSetTrailingPadding','IconImageSequence', 
              'OtherPatientIDsSequence','PatientInsurancePlanCodeSequence','PerformedProtocolCodeSequence',
              'ReferencedPatientSequence','ContributingEquipmentSequence',
              'DeidentificationMethodCodeSequence', 'ScheduledProtocolCodeSequence','DisplayedAreaSelectionSequence','GraphicLayerSequence',
              'SoftcopyVOILUTSequence','ReferencedRequestSequence','EnergyWindowRangeSequence',
              'PatientOrientationCodeSequence','PatientOrientationModifierCodeSequence','RadiopharmaceuticalInformationSequence',
              'AcquisitionContextSequence','WaveformSequence','IssuerOfAccessionNumberSequence',
              'IssuerOfPatientIDQualifiersSequence','VerifyingObserverSequence',
              'AnatomicRegionSequence','AngularViewVector','DetectorInformationSequence','DetectorVector',
              'EnergyWindowInformationSequence','EnergyWindowVector','InterventionDrugInformationSequence',
              'RotationInformationSequence','RotationVector','GatedInformationSequence','RRIntervalVector','TimeSlotVector',
              'RelatedSeriesSequence','SliceVector','IdenticalDocumentsSequence','CTExposureSequence','GraphicAnnotationSequence',
              'PurposeOfReferenceCodeSequence','RegistrationSequence','OriginalAttributesSequence']
    
    #which
    
    study_uids = os.listdir(root)
    #print(len(study_uids))
    #08-BUD-0256	1.2.124.113532.80.22205.16961.20170512.160053.188470747	1.3.46.670589.33.1.63630211589265829600001.5603026839217296534

    #study_uids=['2.25.102321763836695836514179156254521767846']    
    #study_uids = study_uids[:5]    
    #study_uids = study_uids[4409:]
    
    df = pd.DataFrame()    
    i = 0
      
    
    for istudy, study_uid in enumerate(study_uids):              
        
        print(istudy, study_uid)
        series_uids = os.listdir(os.path.join(root, study_uid))
        
                
        #print ('\n{} {} {} {} {}'.format(study_uid, studyuid, pat, modality,site),  )       
        #continue
        
        
        for series_uid in series_uids:
            
            print(i, series_uid)
            path_series = os.path.join(root, study_uid, series_uid)
            alldcm = [fn for fn in os.listdir(path_series) if fn.endswith('dcm')]            
            
            count = len(alldcm)
                        
            try: 
                fdcm = os.path.join(path_series, alldcm[0])
                dcm = pydicom.read_file(fdcm, deferSize, force=force)
                #dcm = pydicom.read_file(fdcm)             #stop_before_pixels=False
            except pydicom.filereader.InvalidDicomError:
                continue  # skip non-dicom file
            except Exception as why:          
                print('Warning:', why)
                continue
            
            id_split = str(dcm.PatientID).split('-')
            if id_split[0] == 'DP':
                site = 'P'+ id_split[1]
            else:
                site = 'P'+ id_split[0]
            
            df.loc[i,'PatientID'] = dcm.PatientID
            df.loc[i,'StudyInstanceUID'] = dcm.StudyInstanceUID            
            df.loc[i,'SeriesInstanceUID'] = dcm.SeriesInstanceUID         
            df.loc[i,'Modality'] = dcm.Modality
            df.loc[i,'Count'] = count
            df.loc[i,'Site'] = site
            
            df.loc[i,'SeriesDescription'] = ''
            df.loc[i,'SeriesNumber'] = 0
            
            if 0:
                if 'SeriesDescription' in dcm:          
                    #data_element = dcm.data_element('SeriesDescription')
                    #print (data_element.value, data_element.VR, data_element.VM,str(data_element.value))                                    
                    df.loc[i,'SeriesDescription'] = str(dcm.SeriesDescription)
                if 'SeriesNumber' in dcm:                
                    df.loc[i,'SeriesNumber'] = str(dcm.SeriesNumber)
                
            
            #SOPInstanceUID = dcm.SOPInstanceUID
    
            tags = dcm.dir()
            for tag in tags:
                
                try:
        
                    data_element = dcm.data_element(tag)
                    #print(data_element)
                except NotImplementedError:
                    continue  # skip non-dicom file
                except Exception as why:          
                    print('Warning:', why)
                    continue
                
                if data_element is None:
                    continue
                
                #data_element.tag (0018, 1151)
                
                #print (tag)
                if (tag not in exclude_tags):
                    #print (tag, data_element.value, data_element.VR, data_element.VM,)                    
                    #print()
                
                    df.loc[i,tag] = str(data_element.value)
                    
                    '''
                    if data_element.VM > 1:                       
                        #print ('bigher1')
                        #df.at[i, tag] = str(data_element.value)   
                        df.loc[i,tag] = str(data_element.value)
                    else:
                        #print ('one')
                        df.loc[i,tag] = data_element.value
                        #df.at[i, tag] = data_element.value
                    ''' 
            
            i += 1 #series
        
        print() #study

    #linear
    df.sort_values('PatientID', inplace=True)
    df.reset_index(drop=True, inplace=True)
    #df.sort_index(inplace=True)
    
    #ORDER
    #df = df.reindex(columns=sorted(df.columns))


    #ordered
    #https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.DataFrame.set_index.html
    #DataFrame.set_index(self, keys, drop=True, append=False, inplace=False, verify_integrity=False)[source]¶
    df2 = df.set_index(['PatientID','StudyInstanceUID','SeriesInstanceUID'])
    #df2.drop('SOPInstanceUID', axis=1, inplace=True)    
    
    #https://pandas.pydata.org/pandas-docs/version/0.23.3/generated/pandas.DataFrame.sort_index.html
    #DataFrame.sort_index(axis=0, level=None, ascending=True, inplace=False, kind='quicksort', na_position='last', sort_remaining=True, by=None)[source]
    df2.sort_index(inplace=True)
    
    
    writer = pd.ExcelWriter(fout)            
    df_to_excel(writer, "ordered", df2)    
    df_to_excel(writer, "linear", df)    
    
    writer.save()


def extract_specific_tags_df(root, suids=[], specific_tags=['PatientID'], NumSamples=None, cols_first=[]):

    # specific_tags = ['Site', 'PatientID', 'StudyInstanceUID', 'SeriesInstanceUID', 'SOPInstanceUID', 'AcquisitionDate',
    #                  'SeriesNumber', 'Count', 'SeriesDescription', 'Modality', 'AcquisitionTime', 'NumberOfFrames',
    #                  'Rows', 'Columns', 'InstanceNumber', 'PatientSex', 'PatientAge', 'ProtocolName',
    #                  'ContrastBolusAgent', 'ImageComments', 'PixelSpacing', 'SliceThickness', 'FilterType',
    #                  'ConvolutionKernel', 'ReconstructionDiameter', 'RequestedProcedureDescription',
    #                  'ContrastBolusStartTime', 'NominalPercentageOfCardiacPhase', 'CardiacRRIntervalSpecified', 'StudyDate']

    specific_tags = ['Site', 'PatientID', 'StudyInstanceUID', 'SeriesInstanceUID', 'AcquisitionDate',
                     'SeriesNumber', 'Count', 'SeriesDescription', 'Modality', 'AcquisitionTime', 'NumberOfFrames',
                     'Rows', 'Columns', 'InstanceNumber', 'PatientSex', 'PatientAge', 'ProtocolName',
                     'ContrastBolusAgent', 'ImageComments', 'PixelSpacing', 'SliceThickness', 'FilterType',
                     'ConvolutionKernel', 'ReconstructionDiameter', 'RequestedProcedureDescription',
                     'ContrastBolusStartTime', 'NominalPercentageOfCardiacPhase', 'CardiacRRIntervalSpecified', 'StudyDate']


    if not suids:
        study_uids = os.listdir(root)
    else:
        study_uids = suids
    
    df = pd.DataFrame(columns=specific_tags)
    i = 0
    
    if NumSamples is None:
        NumSamples = len(study_uids)
    
    specific_tags_dcm = specific_tags.copy()
    if 'Site' in specific_tags_dcm: specific_tags_dcm.remove('Site')
    if 'Count' in specific_tags_dcm: specific_tags_dcm.remove('Count')
    
    for istudy, study_uid in enumerate(study_uids[0:NumSamples]):              
        
        print(istudy, study_uid)
        if not os.path.exists(os.path.join(root, study_uid)): continue

        series_uids = os.listdir(os.path.join(root, study_uid))
        
        for series_uid in series_uids:
            
            path_series = os.path.join(root, study_uid, series_uid)
            #alldcm = [fn for fn in os.listdir(path_series) if fn.endswith('dcm')]    
            alldcm = glob(path_series + '/*.dcm')
            
            # Check if multi slice or sinle slice format
            #fdcm = os.path.join(path_series, alldcm[0])
            ds = pydicom.dcmread(alldcm[0], force = False, defer_size = 256, specific_tags = ['NumberOfFrames'], stop_before_pixels = True)
            try:        
                NumberOfFrames = ds.data_element('NumberOfFrames').value
                MultiSlice = True                              
            except: 
                NumberOfFrames=''
                MultiSlice = False
                
            #print('MultiSlice', MultiSlice)
            #print('NumberOfFrames', NumberOfFrames)
                
            if MultiSlice:
                for dcm in alldcm[0:1]:
                    try:
                        ds = pydicom.dcmread(dcm, force = False, defer_size = 256, specific_tags = specific_tags_dcm, stop_before_pixels = True)
                    except Exception as why:          
                        print('Exception:', why)
                        continue
                    if 'Site' in specific_tags:
                        df.loc[i,'Site'] = 'P'+ str(ds.PatientID).split('-')[0]
                    if 'Count' in specific_tags:
                        df.loc[i,'Count'] = len(alldcm)
                    
                    for tag in specific_tags:
                        try:        
                            data_element = ds.data_element(tag)                                
                        except:                          
                            continue                
                        if data_element is None:
                            continue
                        df.loc[i,tag] = str(data_element.value)
                    i += 1 #series
            else:
                try:
                    ds = pydicom.dcmread(alldcm[0], force = False, defer_size = 256, specific_tags = specific_tags_dcm, stop_before_pixels = True)
                except Exception as why:          
                    print('Exception:', why)
                    continue
                if 'Site' in specific_tags:
                    df.loc[i,'Site'] = 'P'+ str(ds.PatientID).split('-')[0]
                if 'Count' in specific_tags:
                    df.loc[i,'Count'] = len(alldcm)
                
                
                for tag in specific_tags:
                    try:        
                        data_element = ds.data_element(tag)                                
                    except:                          
                        continue                
                    if data_element is None:
                        continue
                    df.loc[i,tag] = str(data_element.value)
                #if 'SOPInstanceUID' in specific_tags:
                #    df.loc[i,'SOPInstanceUID'] = ''
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
        
    return df
           
def extract_specific_tags(root, fout, NumSamples=None):

    df =  extract_specific_tags_df(root, NumSamples=NumSamples)

    
    df.sort_values('PatientID', inplace=True)
    df.reset_index(drop=True, inplace=True)        
    
    #df2 = df.set_index(['PatientID','StudyInstanceUID','SeriesInstanceUID'])    
    #df2.sort_index(inplace=True)    
    
    writer = pd.ExcelWriter(fout)            
    #df_to_excel(writer, "ordered", df2)    
    df_to_excel(writer, "linear", df)    
    writer.save()
 
def compare_specific_tags(root, root2, fout):

    df =  extract_specific_tags_df(root)
    #df = df.head()
    suids = df['StudyInstanceUID'].unique().tolist()
    #print(suids)
    
    df2 =  extract_specific_tags_df(root2, suids)
    
    
    df = df.set_index(['PatientID','StudyInstanceUID','SeriesInstanceUID'])    
    df.sort_index(inplace=True) 
    
    df2 = df2.set_index(['PatientID','StudyInstanceUID','SeriesInstanceUID'])    
    df2.sort_index(inplace=True) 
    
    writer = pd.ExcelWriter(fout)            
    df_to_excel(writer, "new", df)    
    df_to_excel(writer, "old", df2)    
    writer.save()
    

if __name__=='__main__':      
    
    start_time = time.time()    
    # root = 'G:/discharge'
    # fout = 'H:/cloud/cloud_data/Projects/CACSFilter/data/tmp/discharge_dicom.xlsx'
    # extract_specific_tags(root, fout)

    print("--- %s seconds ---" % (time.time() - start_time))
    

    
