# -*- coding: utf-8 -*-
"""
Created on Wed May 13 13:59:31 2020

@author: bernifoellmer
"""
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
from filterTenStepsGuide import filter_CACS_10StepsGuide, filter_CACS, filter_NCS, filterReconstruction, filter_CTA, filer10StepsGuide

tags_dicom=['Count', 'Site', 'PatientID', 'StudyInstanceUID', 'SeriesInstanceUID',
        'Modality', 'SeriesNumber', 'SeriesDescription', 'AcquisitionDate',
        'AcquisitionTime', 'NumberOfFrames', 'Rows', 'Columns',
        'InstanceNumber', 'PatientSex', 'PatientAge', 'ProtocolName',
        'ContrastBolusAgent', 'ImageComments', 'PixelSpacing', 'SliceThickness',
        'FilterType', 'ConvolutionKernel', 'ReconstructionDiameter',
        'RequestedProcedureDescription', 'ContrastBolusStartTime','NominalPercentageOfCardiacPhase',
        'CardiacRRIntervalSpecified', 'StudyDate']

cols_first = ['Site', 'PatientID', 'StudyInstanceUID', 'SeriesInstanceUID', 
              'AcquisitionDate', 'SeriesNumber', 'Count', 'SeriesDescription']

patient_status = ['OK', 'EXCLUDED', 'MISSING_CACS', 'MISSING_CTA', 'MISSING_NC_CACS', 'MISSING_NC_CTA']
patient_status_manual = ['OK', 'EXCLUDED', 'UNDEFINED', 'INPROGRESS']
patient_status_manualStr = '"' + 'OK,' + 'EXCLUDED,' + 'UNDEFINED,' + 'INPROGRESS,' + '"'


#scanClasses = defaultdict(lambda:None,{0: 'UNDEFINED', 1: 'CACS', 2: 'CTA', 3: 'NCS_CACS', 4: 'NCS_CTA', 5: 'ICA'})
scanClasses = defaultdict(lambda:None,{'UNDEFINED': 0, 'CACS': 1, 'CTA': 2, 'NCS_CACS': 3, 'NCS_CTA': 4, 'ICA': 5, 'OTHER': 6})
scanClassesInv = defaultdict(lambda:None,{0: 'UNDEFINED', 1: 'CACS', 2: 'CTA', 3: 'NCS_CACS', 4: 'NCS_CTA', 5: 'ICA', 6: 'OTHER'})
scanClassesStr = '"' + 'UNDEFINED,' + 'CACS,' + 'CTA,' + 'NCS_CACS,' + 'NCS_CTA,' + 'ICA,' + 'OTHER' +'"'
scanClassesManualStr = '"' + 'UNDEFINED,' + 'CACS,' + 'CTA,' + 'NCS_CACS,' + 'NCS_CTA,' + 'ICA,' + 'OTHER,' + 'PROBLEM,' + 'QUESTION,' +'"'
# recoClasses = ['FBP', 'IR', 'UNDEFINED']
changeClasses = ['NO_CHANGE', 'SOURCE_CHANGE', 'MASTER_CHANGE', 'MASTER_SOURCE_CHANGE']

def setColor(workbook, sheet, rows, NumColumns, color):
    for r in rows:
        if r % 100 == 0:
            print('index:', r, '/', max(rows))
        for c in range(1,NumColumns):
            cell = sheet.cell(r, c)
            cell.fill = PatternFill(start_color=color, end_color=color, fill_type = 'solid')

def setColorFormula(sheet, formula, color, colorrange):
    dxf = DifferentialStyle(font=Font(color=color))
    r = Rule(type="expression", dxf=dxf, stopIfTrue=True)
    r.formula = [formula]
    sheet.conditional_formatting.add(colorrange, r)

def setBorderFormula(sheet, formula, NumRows, NumColumns):
    column_letter = get_column_letter(NumColumns+1)
    colorrange="B1:" + str(column_letter) + str(NumRows)
    thin = Side(border_style="thin", color="000000")
    border = Border(bottom=thin)
    dxf = DifferentialStyle(border=border)
    r = Rule(type="expression", dxf=dxf, stopIfTrue=True)
    r.formula = [formula]
    sheet.conditional_formatting.add(colorrange, r)
        
    # Set border for index
    for i in range(1, NumRows + 1):
        cell = sheet.cell(i, 1)
        cell.border = Border()
    return sheet
        
def sortFilepath(filepathList):
    filenameList=[]
    folderpathList=[]
    for filepath in filepathList:
        folderpath, filename, _ = splitFilePath(filepath)
        filenameList.append(filename)
        folderpathList.append(folderpath)
        
    dates_str = [x.split('_')[-1] for x in filenameList]
    dates = [datetime.datetime(int(x[4:8]), int(x[2:4]), int(x[0:2])) for x in dates_str]
    idx = list(np.argsort(dates))
    filepathlistsort=[]
    for i in idx:
        filepathlistsort.append(folderpathList[i] + '/' + '_'.join(filenameList[i].split('_')[0:-1]) + '_' + dates[i].strftime("%d%m%Y") + '.xlsx')
    
    return filepathlistsort

def sortFolderpath(folderpath, folderpathList):
    dates_str = [x.split('_')[-1] for x in folderpathList]
    dates = [datetime(int(x[4:8]), int(x[2:4]), int(x[0:2])) for x in dates_str]
    date_str = folderpath.split('_')[-1]
    date = datetime(int(date_str[4:8]), int(date_str[2:4]), int(date_str[0:2]))
    idx = list(np.argsort(dates))
    folderpathSort=[]
    for i in idx:
        folderpathSort.append(folderpathList[i])
        if dates[i] == date:
            break
    return folderpathSort

def isNaN(num):
    return num != num

def splitFilePath(filepath):
    """ Split filepath into folderpath, filename and file extension
    
    :param filepath: Filepath
    :type filepath: str
    """
    folderpath, _ = ntpath.split(filepath)
    head, file_extension = os.path.splitext(filepath)
    folderpath, filename = ntpath.split(head)
    return folderpath, filename, file_extension

 

def update_CACS_10StepsGuide(df_CACS, sheet):
    for index, row in df_CACS.iterrows():
        cell_str = 'AB' + str(index+2)
        cell = sheet[cell_str]
        cell.value = row['CACS10StepsGuide']
        #cell.protection = Protection(locked=False)
    return sheet    

def mergeITT(df_ITT, df_data):
    # Merge ITT table   
    print('Merge ITT table')
    for i in range(len(df_data)):
        patient = df_ITT[df_ITT['ID']==df_data.loc[i, 'PatientID']]
        if len(patient)==1:
            df_data.loc[i, 'ITT'] = patient.iloc[0]['ITT']
            df_data.loc[i, 'Date CT'] = patient.iloc[0]['Date CT']
            df_data.loc[i, 'Date ICA'] = patient.iloc[0]['Date ICA']
    return df_data

def mergeDicom(df_dicom, df_data_old=None):
    print('Merge dicom table')
    if df_data_old is None:
        df_data = df_dicom.copy()
    else:
        idx = df_dicom['SeriesInstanceUID'].isin(df_data_old['SeriesInstanceUID'])
        df_data = pd.concat([df_data_old, df_dicom[idx==False]], axis=0)     
    return df_data

def mergeTracking(df_tracking, df_data, df_data_old=None):
    
    if df_data_old is None:
        
        df_data = df_data.copy()
        df_tracking = df_tracking.copy()
        df_data.replace(to_replace=[np.nan], value='', inplace=True)
        df_tracking.replace(to_replace=[np.nan], value='', inplace=True)
        
        # Merge tracking table
        print('Merge tracking table')
        df_data['Responsible Person Problem'] = ''
        df_data['Date Query'] = ''
        df_data['Date Answer'] = ''
        df_data['Problem Summary'] = ''
        df_data['Results'] = ''
        for index, row in df_tracking.iterrows():
            patient = row['PatientID']
            df_patient = df_data[df_data['PatientID']==patient]
            for indexP, rowP in df_patient.iterrows(): 
                # Update 'Problem Summary'
                if df_data.loc[indexP, 'Problem Summary']=='':
                    df_data.loc[indexP, 'Problem Summary'] = row['Problem Summary']
                else:
                    df_data.loc[indexP, 'Problem Summary'] = df_data.loc[indexP, 'Problem Summary'] + ' | ' + row['Problem Summary']
                # Update 'results'
                if df_data.loc[indexP, 'Results']=='':
                    df_data.loc[indexP, 'Results'] = row['results']
                else:
                    df_data.loc[indexP, 'Results'] = df_data.loc[indexP, 'Results'] + ' | ' + row['results']
    else:
        
        df_data = df_data.copy()
        df_data_old = df_data_old.copy()
        df_tracking = df_tracking.copy()
        df_data.replace(to_replace=[np.nan], value='', inplace=True)
        df_data_old.replace(to_replace=[np.nan], value='', inplace=True)
        df_tracking.replace(to_replace=[np.nan], value='', inplace=True)
    
        l = len(df_data_old)
        df_data['Responsible Person Problem'] = ''
        df_data['Date Query'] = ''
        df_data['Date Answer'] = ''
        df_data['Problem Summary'] = ''
        df_data['Results'] = ''
        df_data['Responsible Person Problem'][0:l] = df_data_old['Responsible Person Problem']
        df_data['Date Query'][0:l] = df_data_old['Date Query']
        df_data['Date Answer'][0:l] = df_data_old['Date Answer']
        df_data['Problem Summary'][0:l] = df_data_old['Problem Summary']
        df_data['Results'][0:l] = df_data_old['Results']
        
        for index, row in df_tracking.iterrows():
            patient = row['PatientID']
            df_patient = df_data[df_data['PatientID']==patient]
            for indexP, rowP in df_patient.iterrows():
                # Update 'Problem Summary'
                if df_data.loc[indexP, 'Problem Summary']=='':
                    df_data.loc[indexP, 'Problem Summary'] = row['Problem Summary']
                else:
                    if not row['Problem Summary'] in df_data.loc[indexP, 'Problem Summary']:
                        df_data.loc[indexP, 'Problem Summary'] = df_data.loc[indexP, 'Problem Summary'] + ' | ' + row['Problem Summary']
                # Update 'results'
                if df_data.loc[indexP, 'Results']=='':
                    df_data.loc[indexP, 'Results'] = row['results']
                else:
                    if not row['results'] in df_data.loc[indexP, 'Results']:
                        df_data.loc[indexP, 'Results'] = df_data.loc[indexP, 'Results'] + ' | ' + row['results']
    return df_data

def mergeEcrf(df_ecrf, df_data):
    
    # Merge ecrf table
    print('Merge ecrf table')
    df_data['1. Date of CT scan'] = ''
    for index, row in df_ecrf.iterrows():
        patient = row['Patient identifier']
        df_patient = df_data[df_data['PatientID']==patient]
        for indexP, rowP in df_patient.iterrows(): 
            # Update '1. Date of CT scan'
            df_data.loc[indexP, '1. Date of CT scan'] = row['1. Date of CT scan']

    return df_data

def mergePhase_exclude_stenosis(df_phase_exclude_stenosis, df_data):
    # Merge phase_exclude_stenosis
    print('Merge phase_exclude_stenosis table')
    df_data['phase_i0011'] = ''
    df_data['phase_i0012'] = ''
    for index, row in df_phase_exclude_stenosis.iterrows():
        patient = row['mnpaid']
        df_patient = df_data[df_data['PatientID']==patient]
        for indexP, rowP in df_patient.iterrows(): 
            # Update tags
            if df_data.loc[indexP, 'phase_i0011']=='':
                df_data.loc[indexP, 'phase_i0011'] = str(row['phase_i0011'])
            else:
                df_data.loc[indexP, 'phase_i0011'] = str(df_data.loc[indexP, 'phase_i0011']) + ', ' + str(row['phase_i0011'])
                
            if df_data.loc[indexP, 'phase_i0012']=='':
                df_data.loc[indexP, 'phase_i0012'] = str(row['phase_i0012'])
            else:
                df_data.loc[indexP, 'phase_i0012'] = str(df_data.loc[indexP, 'phase_i0011']) + ', ' + str(row['phase_i0011'])

    return df_data

def mergePrct(df_prct, df_data):
    # Merge phase_exclude_stenosis
    print('Merge prct table')
    df_data['other_best_phase'] = ''
    df_data['rca_best_phase'] = ''
    df_data['lad_best_phase'] = ''
    df_data['lcx_best_phase'] = ''
    for index, row in df_prct.iterrows():
        patient = row['PatientId']
        df_patient = df_data[df_data['PatientID']==patient]
        for indexP, rowP in df_patient.iterrows(): 
            # Update tags
            df_data.loc[indexP, 'other_best_phase'] = row['other_best_phase']
            df_data.loc[indexP, 'rca_best_phase'] = row['rca_best_phase']
            df_data.loc[indexP, 'lad_best_phase'] = row['lad_best_phase']
            df_data.loc[indexP, 'lcx_best_phase'] = row['lcx_best_phase']
    return df_data


def mergeStenosis_bigger_20_phase(df_stenosis_bigger_20_phases, df_data):
    # Merge phase_exclude_stenosis
    print('Merge Stenosis_bigger_20_phase table')
    df_data['STENOSIS'] = ''
    
    patientnames = df_stenosis_bigger_20_phases['mnpaid'].unique()
    df_stenosis_bigger_20_phases.replace(to_replace=[np.nan], value='', inplace=True)
    for patient in patientnames:
        patientStenose = df_stenosis_bigger_20_phases[df_stenosis_bigger_20_phases['mnpaid']==patient]
        sten = ''
        for index, row in patientStenose.iterrows(): 
            art=''
            if row['LAD']==1:
                art = 'LAD'
            if row['RCA']==1:
                art = 'RCA'
            if row['LMA']==1:
                art = 'LMA'
            if row['LCX']==1:
                art = 'LCX'
            if sten =='':
                if not art=='':
                    sten = art + ':' + str(row['sten_i0231 (Phase #1)']) + ':' + str(row['sten_i0241']) + ':' + str(row['sten_i0251'])
            else:
                if not art=='':
                    sten = sten + ', ' + art + ':' + str(row['sten_i0231 (Phase #1)']) + ':' + str(row['sten_i0241']) + ':' + str(row['sten_i0251'])
        
        df_patient = df_data[df_data['PatientID']==patient]
        for indexP, rowP in df_patient.iterrows(): 
            df_data.loc[indexP, 'STENOSIS'] = sten

    return df_data


#########################################################################
def freeze(writer, sheetname, df):
    NumRows=1
    NumCols=1
    df.to_excel(writer, sheet_name = sheetname, freeze_panes = (NumCols, NumRows))
 

def highlight_columns(sheet, columns=[], color='A5A5A5', offset=2):
    for col in columns:
        cell = sheet.cell(1, col+offset)
        cell.fill = PatternFill(start_color=color, end_color=color, fill_type = 'solid')
    return sheet  

def setAccessRights(sheet, columns=[], promt='', promptTitle='', formula1='"Dog,Cat,Bat"'):
    for column in columns:
        column_letter = get_column_letter(column+2)
        dv = DataValidation(type="list", formula1=formula1, allow_blank=True)
        dv.prompt = promt
        dv.promptTitle = promptTitle
        column_str = column_letter + str(1) + ':' + column_letter + str(1048576)
        dv.add(column_str)
        sheet.add_data_validation(dv)
    return sheet 

def setComment(sheet, columns=[], comment=''):
    for column in columns:
        column_letter = get_column_letter(column+2)
        dv = DataValidation()
        dv.prompt = comment
        column_str = column_letter + str(1) + ':' + column_letter + str(1048576)
        dv.add(column_str)
        sheet.add_data_validation(dv)
    return sheet  

def setLock(sheet, column='BF'):
    
    sheet.protection.sheet = True
    
    # Unlock all cells
    for column in sheet.columns:
        for cell in column:
            if cell.protection.locked:
                cell.protection = Protection(locked=False)

    # Lock selected column                               
    cell.protection = Protection(locked=False)
    for column in sheet.columns:
        for cell in column:
            if cell.column_letter == column:
                cell.protection = Protection(locked=True)

    return sheet  

def checkTables(settings):

    # Check if requird tables exist
    tables=['filepath_dicom', 'filepath_ITT', 'filepath_ecrf', 'filepath_prct',
            'filepath_phase_exclude_stenosis', 'filepath_stenosis_bigger_20_phases', 'filepath_tracking']
    for table in tables:
        if not os.path.isfile(settings[table]):
            raise ValueError("Source file " + settings[table] + ' does not exist. Please copy file in the correct directory!')
    return True

def createData(settings, NumSamples=None):
        
    dicom_tag_order = settings['dicom_tag_order']
    columns_first = settings['columns_first']

    # Extract dicom data
    df_dicom = pd.read_excel(settings['filepath_dicom'], index_col=0)
    df_dicom = df_dicom[dicom_tag_order]
    df_dicom = df_dicom[(df_dicom['Modality']=='CT') | (df_dicom['Modality']=='OT')]
    df_dicom = df_dicom.reset_index(drop=True)
    cols = df_dicom.columns.tolist()
    cols_new = columns_first + [x for x in cols if x not in columns_first]
    df_dicom = df_dicom[cols_new]
    df_data = df_dicom.copy()
    df_data = df_data.reset_index(drop=True)
    # Select subset
    if NumSamples is not None:
        df_data = df_data[NumSamples[0]:NumSamples[1]]
    # Extract ecrf data
    df_ecrf = pd.read_excel(settings['filepath_ecrf'])
    df_data = mergeEcrf(df_ecrf, df_data)
    # Extract ITT 
    df_ITT = pd.read_excel(settings['filepath_ITT'], 'Tabelle1')
    df_data = mergeITT(df_ITT, df_data)
    # Extract phase_exclude_stenosis 
    df_phase_exclude_stenosis = pd.read_excel(settings['filepath_phase_exclude_stenosis'])
    df_data = mergePhase_exclude_stenosis(df_phase_exclude_stenosis, df_data)
    # Extract prct
    df_prct = pd.read_excel(settings['filepath_prct'])
    df_data = mergePrct(df_prct, df_data)
    # Extract stenosis_bigger_20_phases
    df_stenosis_bigger_20_phases = pd.read_excel(settings['filepath_stenosis_bigger_20_phases'])
    df_data = mergeStenosis_bigger_20_phase(df_stenosis_bigger_20_phases, df_data)  
    # Reoder columns
    cols = df_data.columns.tolist()
    cols_new = columns_first + [x for x in cols if x not in columns_first]
    df_data.to_excel(settings['filepath_data'])


def updateData(folderpath_master, folderpath_master_before):
    
    date = folderpath_master.split('_')[-1]
    filepathMaster = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')
    
    folderpath_master_before_list = glob(folderpath_master_before + '/*master*')
    folderpath_master_before_list = sortFolderpath(folderpath_master, folderpath_master_before_list)
    
    #filepathMaster = glob(folderpath_master + '/*master*.xlsx')[0]
    folderpathComponent = glob(folderpath_master + '/*components*')[0]
    filepathSource = glob(folderpathComponent + '/*discharge_data*')[0]
    filepath_data = os.path.join(folderpathComponent, 'discharge_master_data_' + date + '.xlsx')
    filepath_change = os.path.join(folderpathComponent, 'discharge_change_' + date + '.xlsx')
    
    # Filer folderpath_master_before_list by process file existance
    folderpath_master_before_list_tmp = folderpath_master_before_list
    folderpath_master_before_list=[]
    for folder in folderpath_master_before_list_tmp:
        if len(glob(folder + '/*process*.xlsx'))>0:
            folderpath_master_before_list.append(folder)
    
    filepathMasters = [glob(folder + '/*master*.xlsx')[0] for folder in folderpath_master_before_list]
    filepathMasters = filepathMasters + glob(folderpath_master_before_list[-1] + '/*process*.xlsx')
    folderpathComponents = [glob(folder + '/*components*')[0] for folder in folderpath_master_before_list] + [folderpathComponent]
    filepathSources = [glob(folder + '/*data*.xlsx')[0] for folder in folderpathComponents]

    changeClasses = ['NO_CHANGE', 'SOURCE_CHANGE', 'MASTER_CHANGE', 'MASTER_SOURCE_CHANGE', 'MASTER_SOURCE_CHANGE_CRITICAL', 'NEW_DATA']
    
    shape = pd.read_excel(filepathSources[-1], index_col=0).shape
    DM = []
    for file in filepathMasters:
        df = pd.read_excel(file, index_col=0)
        df.replace(to_replace=[np.nan], value='', inplace=True)
        DM.append(df.take(range(0,shape[1]), axis=1).copy())

    #shape = pd.read_excel(filepathSources[-1], index_col=0).shape
    DS = []
    for file in filepathSources:
        df = pd.read_excel(file, index_col=0)
        df.replace(to_replace=[np.nan], value='', inplace=True)
        DS.append(df.take(range(0,shape[1]), axis=1).copy())
        
    shape = [max(DM[-1].shape[0], DS[-1].shape[0]), DM[-1].shape[1]]
    
    # Harmonize DM
    dfM = DM[-1]
    row_new_data = pd.DataFrame('NEW_DATA', index=np.arange(1), columns=dfM.columns)
    for k in range(0,len(DM)):
        df = pd.DataFrame('NEW_DATA', index=np.arange(shape[0]), columns=dfM.columns)
        df_old = DM[k]
        for index_new, row_new in dfM.iterrows():
            index_old = df_old.index[df_old['SeriesInstanceUID']== row_new['SeriesInstanceUID']]
            if len(index_old)==0:
                df.loc[index_new] = row_new_data.loc[0].copy()
            else:
                df.loc[[index_new]] = df_old.loc[index_old]
        DM[k] = df.copy()
        
    # Harmonize DS
    dfS = DS[-1]
    row_new_data = pd.DataFrame('NEW_DATA', index=np.arange(1), columns=dfS.columns)
    for k in range(0,len(DS)):
        df = pd.DataFrame('NEW_DATA', index=np.arange(shape[0]), columns=dfS.columns)
        df_old = DS[k]
        for index_new, row_new in dfS.iterrows():
            index_old = df_old.index[df_old['SeriesInstanceUID']== row_new['SeriesInstanceUID']]
            if len(index_old)==0:
                df.loc[index_new] = row_new_data.loc[0].copy()
            else:
                df.loc[[index_new]] = df_old.loc[index_old]
        DS[k] = df.copy()
            
    # Master Change
    #shape = DM[-1].shape
    DMC = pd.DataFrame('NO_CHANGE', index=np.arange(shape[0]), columns=DM[0].columns)
    for i in range(0, shape[0]):
        for j in range(0, shape[1]):
            values=[]
            for k in range(0,len(DM)):
                values.append(DM[k].iloc[i,j])
                if DM[k].iloc[i,j]=='NEW_DATA':
                    values=[]
            if len(values)<2:
                DMC.iloc[i,j] = 'NEW_DATA'
            else:
                if not values.count(values[0]) == len(values):
                    DMC.iloc[i,j] = 'MASTER_CHANGE'
            
                
    DMC['PatientID'] = DM[-1]['PatientID']
    DMC['StudyInstanceUID'] = DM[-1]['StudyInstanceUID']
    DMC['SeriesInstanceUID'] = DM[-1]['SeriesInstanceUID']

    # Source Change
    shape = DS[-1].shape
    DSC = pd.DataFrame('NO_CHANGE', index=np.arange(len(DS[0])), columns=DS[0].columns)
    for i in range(0, shape[0]):
        for j in range(0, shape[1]):
            values=[]
            for k in range(0,len(DS)):
                values.append(DS[k].iloc[i,j])
                if DS[k].iloc[i,j]=='NEW_DATA':
                    values=[]
            if len(values)<2:
                DSC.iloc[i,j] = 'NEW_DATA'
            else:
                if not values.count(values[0]) == len(values):
                    DSC.iloc[i,j] = 'SOURCE_CHANGE'

    DSC['PatientID'] = DS[-1]['PatientID']
    DSC['StudyInstanceUID'] = DS[-1]['StudyInstanceUID']
    DSC['SeriesInstanceUID'] = DS[-1]['SeriesInstanceUID']

    # Datset Changes
    #shape = [max(DM[-1].shape[0], DS[-1].shape[0]), DM[-1].shape[1]]
    DC = pd.DataFrame('NO_CHANGE', index=np.arange(shape[0]), columns=DM[-1].columns)
    for i in range(0, shape[0]):
        for j in range(0, shape[1]):
            if DMC.iloc[i,j]=='MASTER_CHANGE' and DSC.iloc[i,j]=='NO_CHANGE':
                DC.iloc[i,j] = 'MASTER_CHANGE'
            if DMC.iloc[i,j]=='NO_CHANGE' and DSC.iloc[i,j]=='SOURCE_CHANGE':
                DC.iloc[i,j] = 'SOURCE_CHANGE'
            if DMC.iloc[i,j]=='MASTER_CHANGE' and DSC.iloc[i,j]=='SOURCE_CHANGE' and not DS[-1].iloc[i,j]==DM[-1].iloc[i,j]:
                DC.iloc[i,j] = 'MASTER_SOURCE_CHANGE_CRITICAL'
            if DMC.iloc[i,j]=='MASTER_CHANGE' and DSC.iloc[i,j]=='SOURCE_CHANGE' and DS[-1].iloc[i,j]==DM[-1].iloc[i,j]:
                DC.iloc[i,j] = 'MASTER_SOURCE_CHANGE'
            if DMC.iloc[i,j]=='NEW_DATA' or DSC.iloc[i,j]=='NEW_DATA':
                DC.iloc[i,j] = 'NEW_DATA'
    DC['PatientID'] = DSC['PatientID']
    DC['StudyInstanceUID'] = DSC['StudyInstanceUID']
    DC['SeriesInstanceUID'] = DSC['SeriesInstanceUID']

    #df_master_process =pd.read_excel(filepathMasters[-1], index_col=0)
    df_data = pd.DataFrame('NO_CHANGE', index=np.arange(shape[0]), columns=DM[-1].columns)
    for i in range(0, shape[0]):
        for j in range(0, shape[1]):
            if DC.iloc[i,j]=='NO_CHANGE':
                df_data.iloc[i,j] = DM[-1].iloc[i,j]
            if DC.iloc[i,j]=='MASTER_CHANGE':
                df_data.iloc[i,j] = DM[-1].iloc[i,j]
            if DC.iloc[i,j]=='SOURCE_CHANGE':
                df_data.iloc[i,j] = DS[-1].iloc[i,j]
            if DC.iloc[i,j]=='MASTER_SOURCE_CHANGE_CRITICAL':
                df_data.iloc[i,j] = DM[-1].iloc[i,j]
            if DC.iloc[i,j]=='MASTER_SOURCE_CHANGE':
                df_data.iloc[i,j] = DM[-1].iloc[i,j]
            if DC.iloc[i,j]=='NEW_DATA':
                if DM[-1].iloc[i,j]=='NEW_DATA':
                    df_data.iloc[i,j] = DS[-1].iloc[i,j]
                else:
                    df_data.iloc[i,j] = DS[-1].iloc[i,j]
                    
    df_data['PatientID'] = DSC['PatientID']
    df_data['StudyInstanceUID'] = DSC['StudyInstanceUID']
    df_data['SeriesInstanceUID'] = DSC['SeriesInstanceUID']

    df_data.to_excel(filepath_data)
    DC.to_excel(filepath_change)
    
def updatePredictions(folderpath_master):
    dataname='discharge_master_data_'
    createPredictions(folderpath_master, dataname)
    

def createPredictions(settings):

    df_data = pd.read_excel(settings['filepath_data'])
    df_pred = pd.DataFrame()
    
    # Filter by CACS based on 10-Steps-Guide
    df = filter_CACS_10StepsGuide(df_data)
    df_pred['CACS10StepsGuide'] = df['CACS10StepsGuide']
    
    # Filter by CACS based  selection
    df = filter_CACS(df_data)
    df_pred['CACS'] = df['CACS']
    
    # Filter by NCS_CACS and NCS_CTA based on  criteria
    df = filter_NCS(df_data)
    df_pred['NCS_CTA'] = df['NCS_CTA']
    df_pred['NCS_CACS'] = df['NCS_CACS']
    
    # Filter by CTA
    df = filter_CTA(settings)
    df_pred['CTA'] = df['CTA'].astype('bool')
    df_pred['CTA_phase'] = df['phase']
    df_pred['CTA_arteries'] = df['arteries']
    df_pred['CTA_source'] = df['source']
   
    # Filter by ICA
    df = pd.DataFrame('', index=np.arange(len(df_pred)), columns=['ICA'])
    df_pred['ICA'] = df['ICA']
 
    # Filter by reconstruction
    df = filterReconstruction(df_data, settings)
    df_pred['RECO'] = df['RECO']
    
    # Filter by reconstruction
    #df_pred['RFRECO'] = ''
    
    # Predict CLASS
    classes = ['CACS', 'CTA', 'NCS_CTA', 'NCS_CACS']
    for i in range(len(df_pred)):
        if i % 1000 == 0:
            print('index:', i, '/', len(df_pred))
        value=''
        for c in classes:
            if df_pred.loc[i, c]:
                if value=='':
                    value = value + c
                else:
                    value = value + '+' + c
        if value == '':
            value = 'UNDEFINED'
        df_pred.loc[i, 'CLASS'] = value
        
    # Save predictions    
    df_pred.to_excel(settings['filepath_prediction'])

def updateRFClassification(folderpath_master, folderpath_master_before):

    date = folderpath_master.split('_')[-1]
    folderpath_components = os.path.join(folderpath_master, 'discharge_components_' + date)
    filepath_rfc = os.path.join(folderpath_components, 'discharge_rfc_' + date + '.xlsx')
    
    folderpath_master_before_list = glob(folderpath_master_before + '/*master*')
    folderpath_master_before_list = sortFolderpath(folderpath_master, folderpath_master_before_list)
    filepathMasters = glob(folderpath_master_before_list[-2] + '/*process*.xlsx')
    
    
    date_before = folderpath_master_before_list[-2].split('_')[-1]
    df_master = pd.read_excel(filepathMasters[0], sheet_name='MASTER_' + date_before)
    columns = ['RFCLabel', 'RFCClass', 'RFCConfidence']
    df_rfc = pd.DataFrame('UNDEFINED', index=np.arange(len(df_master)), columns=columns)
    df_rfc[columns] = df_master[columns]

    df_rfc.to_excel(filepath_rfc)



        

def updateManualSelection(folderpath_master, folderpath_master_before):
    print('Update manual selection')
    
    date = folderpath_master.split('_')[-1]
    folderpath_components = os.path.join(folderpath_master, 'discharge_components_' + date)
    filepath_manual = os.path.join(folderpath_components, 'discharge_manual_' + date + '.xlsx')
    
    folderpath_master_before_list = glob(folderpath_master_before + '/*master*')
    folderpath_master_before_list = sortFolderpath(folderpath_master, folderpath_master_before_list)
    filepathMasters = glob(folderpath_master_before_list[-2] + '/*process*.xlsx')
    
    date_before = folderpath_master_before_list[-2].split('_')[-1]
    df_master = pd.read_excel(filepathMasters[0], sheet_name='MASTER_' + date_before)
    columns=['ClassManualCorrection', 'Comment']
    df_manual = pd.DataFrame('UNDEFINED', index=np.arange(len(df_master)), columns=columns)
    df_manual[columns] = df_master[columns]
    df_manual.to_excel(filepath_manual)

    
def createManualSelection(settings):
    print('Create manual selection')
    df_data = pd.read_excel(settings['filepath_data'], index_col=0)
    df_manual0 = pd.DataFrame('UNDEFINED', index=np.arange(len(df_data)), columns=['ClassManualCorrection'])
    df_manual1 = pd.DataFrame('', index=np.arange(len(df_data)), columns=['Comment'])
    df_manual2 = pd.DataFrame('', index=np.arange(len(df_data)), columns=['Responsible Person'])
    df_manual = pd.concat([df_manual0, df_manual1, df_manual2], axis=1)
    df_manual.to_excel(settings['filepath_manual'])

    
def updateTracking(folderpath_master, folderpath_master_before):
    print('Update tracking')
    date = folderpath_master.split('_')[-1]
    folderpath_components = os.path.join(folderpath_master, 'discharge_components_' + date)
    if not os.path.isdir(folderpath_components):
        os.mkdir(folderpath_components)
    folderpath_sources = os.path.join(folderpath_master, 'discharge_sources_' + date)
    filepath_data = os.path.join(folderpath_components, 'discharge_data_' + date + '.xlsx')
    filepath_tracking = glob(folderpath_sources + '/*tracking*.xlsx')[0]
    filepath_track = os.path.join(folderpath_components, 'discharge_track_' + date + '.xlsx')
    filepath_master_track = os.path.join(folderpath_components, 'discharge_master_track_' + date + '.xlsx')
    
    folderpath_master_before_list = glob(folderpath_master_before + '/*master*')
    folderpath_master_before_list = sortFolderpath(folderpath_master, folderpath_master_before_list)
    folderpathComponents = [glob(folder + '/*components*')[0] for folder in folderpath_master_before_list]
    date_before = folderpathComponents[-2].split('_')[-1]
    filepath_track_before = os.path.join(folderpathComponents[-2], 'discharge_master_track_' + date_before + '.xlsx')
    
    df_tracking_before = pd.read_excel(filepath_track_before)
    
    df_tracking = pd.read_excel(filepath_tracking, 'tracking table')
    df_tracking.replace(to_replace=[np.nan], value='', inplace=True)
    
    createTracking(folderpath_master, create_master_track=False)
    
    filepath_master_data = os.path.join(folderpath_components, 'discharge_master_data_' + date + '.xlsx')
    df_data = pd.read_excel(filepath_master_data, index_col=0)
    
    df_data = mergeTracking(df_tracking, df_data, df_tracking_before)
    
    df_track = pd.DataFrame()
    df_track = df_data[['Problem Summary', 'Date Query', 'Date Answer', 'Results', 'Responsible Person Problem']]
    df_track.to_excel(filepath_master_track)
    
def createTracking(folderpath_master, create_master_track=True):
    print('Create tracking')
    date = folderpath_master.split('_')[-1]
    folderpath_components = os.path.join(folderpath_master, 'discharge_components_' + date)
    if not os.path.isdir(folderpath_components):
        os.mkdir(folderpath_components)
    folderpath_sources = os.path.join(folderpath_master, 'discharge_sources_' + date)
    #filepath_data = os.path.join(folderpath_components, 'discharge_data_' + date + '.xlsx')
    filepath_tracking = glob(folderpath_sources + '/*tracking*.xlsx')[0]
    filepath_track = os.path.join(folderpath_components, 'discharge_track_' + date + '.xlsx')
    filepath_master_track = os.path.join(folderpath_components, 'discharge_master_track_' + date + '.xlsx')
    
    df_tracking = pd.read_excel(filepath_tracking, 'tracking table')
    df_tracking.replace(to_replace=[np.nan], value='', inplace=True)
    
    filepath_master_data = os.path.join(folderpath_components, 'discharge_master_data_' + date + '.xlsx')
    df_data = pd.read_excel(filepath_master_data, index_col=0)

    
    #df_data = pd.read_excel(filepath_data)
    df_data = mergeTracking(df_tracking, df_data, None)
    
    df_track = pd.DataFrame()
    df_track = df_data[['Problem Summary', 'Date Query', 'Date Answer', 'Results', 'Responsible Person Problem']]
    df_track.to_excel(filepath_track)
    
    if create_master_track:
        copyfile(filepath_track, filepath_master_track)

def createTrackingTable(settings):
    print('Create tracking table')
    df_track = pd.DataFrame(columns=settings['columns_tracking'])
    df_track.to_excel(settings['filepath_master_track'])
    # Update master
    writer = pd.ExcelWriter(settings['filepath_master'], engine="openpyxl", mode="a")
    # Remove sheet if already exist
    sheet_name = 'TRACKING' + '_' + settings['date']
    workbook  = writer.book
    sheetnames = workbook.sheetnames
    if sheet_name in sheetnames:
        sheet = workbook[sheet_name]
        workbook.remove(sheet)
        
    # Add patient ro master
    df_track.to_excel(writer, sheet_name=sheet_name)
    writer.save()
    
def updateTrackingTableFromMaster(folderpath_master, master_process=False):
    print('Update tracking table')

    date = folderpath_master.split('_')[-1]
    folderpath_tracking = os.path.join(folderpath_master, 'discharge_tracking')
    filepath_tracking = glob(folderpath_tracking + '/*tracking*.xlsx')[0]
    if master_process==False:
        filepath_master = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')
    else:
        filepath_master_tmp = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')
        folderpath, filename, file_extension = splitFilePath(filepath_master_tmp)
        filepath_master = os.path.join(folderpath, filename + '_process' + file_extension)
    
    # Read tracking table
    df_tracking = pd.read_excel(filepath_tracking, 'tracking table')
    df_tracking.replace(to_replace=[np.nan], value='', inplace=True)
    df_track = pd.read_excel(filepath_master, 'TRACKING_' + date, index_col=0)
    
    columns_track = df_track.columns
    columns_tracking = df_tracking.columns
    columns_union = columns_track
    
    # ProblemID from int to string
    for index, row in df_track.iterrows():
        df_track.loc[index,'ProblemID'] = str(row['ProblemID']).zfill(6)
    for index, row in df_tracking.iterrows():
        df_tracking.loc[index,'ProblemID'] = str(row['ProblemID']).zfill(6)
        
    ProblemID = 0
    for index, row in df_track.iterrows():
        ProblemID = row['ProblemID']
        problem = df_tracking[df_tracking['ProblemID']==ProblemID]
        if len(problem)==1:
            index = df_tracking['ProblemID'][df_tracking['ProblemID'] == ProblemID].index[0]
            for col in columns_union:
                df_tracking.loc[index, col] = row[col]
        else:
            df_tracking = df_tracking.append(row, ignore_index=True)


    ProblemSummary='"Date of the CT images are wrong,Date of the ICA images are wrong,Missing CT Images,Missing ICA Images,ICA problem,Problem with images"'
   
    # Update master
    writer = pd.ExcelWriter(filepath_tracking, engine="openpyxl", mode="a")
    # Remove sheet if already exist
    sheet_name = 'tracking table'
    workbook  = writer.book
    sheetnames = workbook.sheetnames
    if sheet_name in sheetnames:
        sheet = workbook[sheet_name]
        workbook.remove(sheet)
        
    # Add patient to master
    df_tracking.to_excel(writer, sheet_name=sheet_name, index=False)
    sheet = workbook[sheet_name]
    df_tracking_cols = list(df_tracking.columns)
    sheet = highlight_columns(sheet, columns=[df_tracking_cols.index(col) for col in df_tracking_cols] , color='5B95F9', offset=1)
    sheet = setAccessRights(sheet, columns=[df_tracking_cols.index(col) for col in df_tracking_cols], promt='', promptTitle='', formula1=ProblemSummary)
    writer.save()
    
    
def updateMasterFromTrackingTable(settings):
    print('Create tracking table')
    # Read tracking table
    df_tracking = pd.read_excel(settings['filepath_tracking'], 'tracking table')
    df_tracking.replace(to_replace=[np.nan], value='', inplace=True)
    df_track = pd.read_excel(settings['filepath_master'], 'TRACKING_' + settings['date'], index_col=0)
    
    columns_track = df_track.columns
    columns_tracking = df_tracking.columns
    #columns_union = ['ProblemID', 'PatientID', 'Problem Summary', 'Problem']
    columns_union = columns_track


    if len(df_track)==0:
        ProblemIDMax=-1
        df_track = df_tracking[columns_union]
    else:
        ProblemIDMax = max([int(x) for x in list(df_track['ProblemID'])])
    
    ProblemIDInt = 0
    for index, row in df_tracking.iterrows():
        ProblemID = row['ProblemID']
        if not ProblemID == '':
            index = df_track['ProblemID'][df_track['ProblemID'] == ProblemID].index[0]
            for col in columns_union:
                df_track.loc[index,col] = row[col]
        else:
            ProblemIDInt = ProblemIDMax + 1
            ProblemIDMax = ProblemIDInt
            row['ProblemID'] = str(ProblemIDInt).zfill(6)
            row_new = pd.DataFrame('', index=[0], columns=columns_union)        
            for col in columns_union:
                row_new.loc[0,col] = row[col]
            df_track = df_track.append(row_new, ignore_index=True)
            df_tracking.loc[index,'ProblemID'] = str(ProblemIDInt).zfill(6)
    
    
    # Update master
    writer = pd.ExcelWriter(settings['filepath_master'], engine="openpyxl", mode="a")
    # Remove sheet if already exist
    sheet_name = 'TRACKING' + '_' + settings['date']
    workbook  = writer.book
    sheetnames = workbook.sheetnames
    if sheet_name in sheetnames:
        sheet = workbook[sheet_name]
        workbook.remove(sheet)
        
    # Add patient ro master
    df_track.to_excel(writer, sheet_name=sheet_name)
    writer.save()
    
    # Update tracking
    writer = pd.ExcelWriter(settings['filepath_tracking'], engine="openpyxl", mode="a")
    # Remove sheet if already exist
    sheet_name = 'tracking table'
    workbook  = writer.book
    sheetnames = workbook.sheetnames
    if sheet_name in sheetnames:
        sheet = workbook[sheet_name]
        workbook.remove(sheet)
        
    # Add patient to master
    df_tracking.to_excel(writer, sheet_name=sheet_name, index=False)
    writer.save()

def orderMasterData(df_master, settings):

    # Reoder columns
    # cols_first=['Site', 'PatientID', 'StudyInstanceUID', 'SeriesInstanceUID', 'CLASSExtended', 'RFCClass', 'RFCLabel', 'RFCConfidence',
    #             'ClassManualCorrection', 'Comment', 'Responsible Person', 'Count', 'SeriesDescription', 'SeriesNumber', 'AcquisitionDate']
    
    cols = df_master.columns.tolist()
    cols_new = settings['columns_first'] + [x for x in cols if x not in settings['columns_first']]
    df_master = df_master[cols_new]
    df_master = df_master.sort_values(['PatientID', 'StudyInstanceUID', 'SeriesInstanceUID'], ascending = (True, True, True))
    df_master.reset_index(inplace=True, drop=True)
    return df_master
    
def mergeMaster(settings):
    print('Create master')
    # Read tables
    print('Read discharge_master_data')
    #df_master_data = pd.read_excel(filepath_master_data, index_col=0)
    print('Read discharge_data')
    df_data = pd.read_excel(settings['filepath_data'], index_col=0)
    print('Read discharge_pred')
    df_pred = pd.read_excel(settings['filepath_prediction'], index_col=0)
    df_pred['CTA'] = df_pred['CTA'].astype('bool')
    print('Read discharge_reco')
    #df_reco_load = pd.read_excel(filepath_reco, index_col=0)
    df_reco = pd.DataFrame()
    #df_reco['RFRECO'] = df_reco_load['RFRECO']
    df_reco['RFRECO'] = ''
    print('Read discharge_guide')
    #df_guide = pd.read_excel(filepath_guide, index_col=0)
    #print('Read discharge_patient')
    #df_patient = pd.read_excel(filepath_patient, index_col=0)
    print('Read discharge_rfc')
    df_rfc = pd.read_excel(settings['filepath_rfc'], index_col=0)
    print('Read discharge_manual')
    df_manual = pd.read_excel(settings['filepath_manual'], index_col=0)
    print('Read discharge_track')
    print('Create discharge_master')
    df_master = pd.concat([df_data, df_pred, df_rfc, df_manual, df_reco], axis=1)
    #df_master = orderMasterData(df_master, settings)

    writer = pd.ExcelWriter(settings['filepath_master'], engine="openpyxl", mode="w")
    df_master.to_excel(writer, sheet_name = 'MASTER' + '_' + settings['date'])
    # Add patient data
    writer.save()

def creatMaster(folderpath_master):
    print('Create master')
    date = folderpath_master.split('_')[-1]
    folderpath_components = os.path.join(folderpath_master, 'discharge_components_' + date)
    if not os.path.isdir(folderpath_components):
        os.mkdir(folderpath_components)
    folderpath_sources = os.path.join(folderpath_master, 'discharge_sources_' + date)
    filepath_master = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')
    filepath_data = os.path.join(folderpath_components, 'discharge_data_' + date + '.xlsx')
    filepath_pred = os.path.join(folderpath_components, 'discharge_pred_' + date + '.xlsx')
    filepath_rfc = os.path.join(folderpath_components, 'discharge_rfc_' + date + '.xlsx')
    filepath_manual = os.path.join(folderpath_components, 'discharge_manual_' + date + '.xlsx')
    filepath_track = os.path.join(folderpath_components, 'discharge_track_' + date + '.xlsx')
    
    # Read tables
    print('Read discharge_data')
    df_data = pd.read_excel(filepath_data, index_col=0)
    print('Read discharge_pred')
    df_pred = pd.read_excel(filepath_pred, index_col=0)
    print('Read discharge_rfc')
    df_rfc = pd.read_excel(filepath_rfc, index_col=0)
    print('Read discharge_manual')
    df_manual = pd.read_excel(filepath_manual, index_col=0)
    print('Read discharge_track')
    #df_track = pd.read_excel(filepath_track, index_col=0)
    print('Create discharge_master')
    #df_master = pd.concat([df_data, df_pred, df_rfc, df_manual, df_track], axis=1)
    df_master = pd.concat([df_data, df_pred, df_rfc, df_manual], axis=1)
    return df_master 


def createMasterProcess(folderpath_master):
    # Create master_process
    date = folderpath_master.split('_')[-1]
    filepath_master = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')
    folderpath, filename, file_extension = splitFilePath(filepath_master)
    filepath_process = os.path.join(folderpath, filename + '_process' + file_extension)
    copyfile(filepath_master, filepath_process)


def formatMaster(settings, format='ALL'):
    print('Format master')
    # Read tables
    print('Read discharge_data')
    df_data = pd.read_excel(settings['filepath_data'], index_col=0)
    print('Read discharge_pred')
    df_pred = pd.read_excel(settings['filepath_prediction'], index_col=0)
    print('Read discharge_rfc')
    df_rfc = pd.read_excel(settings['filepath_rfc'], index_col=0)
    print('Read discharge_manual')
    df_manual = pd.read_excel(settings['filepath_manual'], index_col=0)
    print('Read discharge_track')
    df_track = pd.read_excel(settings['filepath_master_track'], index_col=0)
    print('Read patient_data')
    df_patient = pd.read_excel(settings['filepath_patient'], index_col=0)
    print('Create discharge_master')
    #df_master = pd.concat([df_data, df_pred, df_rfc, df_manual, df_track], axis=1)
    #df_master = pd.concat([df_data, df_pred, df_rfc, df_manual], axis=1)
    df_master = pd.read_excel(settings['filepath_master'], sheet_name='MASTER_' + settings['date'], index_col=0)
    df_master = orderMasterData(df_master, settings)
    #filepath_master = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')

    writer = pd.ExcelWriter(settings['filepath_master'], engine="openpyxl", mode="a")
    workbook  = writer.book
    
    colors=['A5A5A5', 'FFFF00', '70AD47', 'FFC000', '5B95F9']
    sheetnames = workbook.sheetnames
    
    for sheetname in sheetnames:
        sheet = workbook[sheetname]
        if 'DATA' in sheetname and (format=='ALL' or format=='DATA'):
            # Clear existing conditional_formatting list
            sheet.conditional_formatting = ConditionalFormattingList()
            sheet.data_validations.dataValidation = []
            # Highlight master
            columns = [x for x in range(0,df_data.shape[1])]
            sheet = highlight_columns(sheet, columns=columns , color=colors[0])
            # Freeze worksheet
            workbook[sheetname].freeze_panes = "D2"
            # Comment data
            sheet = setComment(sheet, columns=columns, comment='Please add a comment why the data has been changed!')
            # Add filter
            sheet.auto_filter.ref = sheet.dimensions
            # Draw border
            sheet = setBorderFormula(sheet, formula='=$C1<>$C2', NumRows=df_data.shape[0], NumColumns=df_data.shape[1])
            
        if 'MASTER' in sheetname and (format=='ALL' or format=='MASTER'):

            # Clear existing conditional_formatting list
            sheet.conditional_formatting = ConditionalFormattingList()
            sheet.data_validations.dataValidation = []
            
            # Highlight master
            df_master_cols = list(df_master.columns)
            df_data_cols = list(df_data.columns)
            df_pred_cols = list(df_pred.columns)
            df_rfc_cols = list(df_rfc.columns)
            df_manual_cols = list(df_manual.columns)
            #df_track_cols = list(df_track.columns)
            sheet = highlight_columns(sheet, columns=[df_master_cols.index(col) for col in df_data_cols] , color=colors[0])
            sheet = highlight_columns(sheet, columns=[df_master_cols.index(col) for col in df_pred_cols] , color=colors[1])
            sheet = highlight_columns(sheet, columns=[df_master_cols.index(col) for col in df_rfc_cols] , color=colors[2])
            sheet = highlight_columns(sheet, columns=[df_master_cols.index(col) for col in df_manual_cols] , color=colors[3])

            #sheet = highlight_columns(sheet, columns=[df_master_cols.index(col) for col in df_track_cols] , color=colors[4])
            
            # Comment  and access rights
            sheet = setComment(sheet, columns=[df_master_cols.index(col) for col in df_data_cols], comment='Please add a comment why the data has been changed!')
            sheet = setComment(sheet, columns=[df_master_cols.index(col) for col in df_pred_cols], comment='Do not change this data!')
            sheet = setAccessRights(sheet, columns=[df_master_cols.index(col) for col in ['RFCLabel']], promt='RFLabel!', promptTitle='DISCHARGE Scan Classes', formula1=scanClassesStr)
            sheet = setComment(sheet, columns=[df_master_cols.index(col) for col in ['RFCClass', 'RFCConfidence']], comment='Do not change this data!')
            sheet = setAccessRights(sheet, columns=[df_master_cols.index(col) for col in ['ClassManualCorrection']], promt='RFLabel!', promptTitle='DISCHARGE Scan Classes', formula1=scanClassesManualStr)
            sheet = setComment(sheet, columns=[df_master_cols.index(col) for col in ['Comment']], comment='Thank you for adding a comment!')
            #sheet = setComment(sheet, columns=[df_master_cols.index(col) for col in df_track_cols], comment='Do not change this data!')
            
            # Highlight based on modality
            setColorFormula(sheet, formula='$G1="CACSExtended"', color="EE1111", colorrange="A1:AI45000")
            setColorFormula(sheet, formula='$G1="CTAExtended"', color="00B050", colorrange="A1:AI45000")
            setColorFormula(sheet, formula='$G1="NCS_CACSExtended"', color="0070C0", colorrange="A1:AI45000")
            setColorFormula(sheet, formula='$G1="NCS_CTAExtended"', color="FFC000", colorrange="A1:AI45000")
            setColorFormula(sheet, formula='$G1="UNDEFINED"', color="000000", colorrange="A1:AI45000")
            
            #setColorFormula(sheet, formula='ISBLANK($CC1)', color="000000", colorrange="A1:AI45000")
            
            # Highligt based on confidence score
            setColorFormula(sheet, formula='$I1<0.5', color="EE1111", colorrange="A2:AI45000")
            setColorFormula(sheet, formula='AND($I1<0.9, $I1>0.3)', color="FFFF00", colorrange="A2:AI45000")
            setColorFormula(sheet, formula='$I1>0.9', color="00B050", colorrange="A2:AI45000")
            
            # Freeze worksheet
            workbook[sheetname].freeze_panes = "D2"
            # Draw border
            sheet = setBorderFormula(sheet, formula='=$C1<>$C2', NumRows=df_master.shape[0], NumColumns=df_master.shape[1])
            # Add filter
            sheet.auto_filter.ref = sheet.dimensions

        if 'PATIENT' in sheetname and (format=='ALL' or format=='PATIENT'):
            # Clear existing conditional_formatting list
            sheet.conditional_formatting = ConditionalFormattingList()
            sheet.data_validations.dataValidation = []
            df_patient_cols = list(df_patient.columns)
            sheet = highlight_columns(sheet, columns=[df_patient_cols.index(col) for col in df_patient_cols] , color=colors[4])
            patient_status_manual = ['OK', 'EXCLUDED', 'UNDEFINED']
            sheet = setComment(sheet, columns=[df_patient_cols.index(col) for col in df_patient_cols[0:-2]], comment='Do not change this data!')
            sheet = setAccessRights(sheet, columns=[df_patient_cols.index(col) for col in ['STATUS_MANUAL_CORRECTION']], promt='PatientLabel', promptTitle='DISCHARGE Patient label', formula1=patient_status_manualStr)
            sheet = setComment(sheet, columns=[df_patient_cols.index(col) for col in ['COMMENT']], comment='Thank you for adding a comment!')
            # Freeze worksheet
            workbook[sheetname].freeze_panes = "D2"
            # Draw border
            sheet = setBorderFormula(sheet, formula='=$C1<>$C2', NumRows=df_master.shape[0], NumColumns=df_master.shape[1])
            # Add filter
            sheet.auto_filter.ref = sheet.dimensions
            
        if 'TRACKING' in sheetname and (format=='ALL' or format=='TRACKING'):
            
            dft = pd.read_excel(settings['filepath_master'], sheet_name=sheetname, index_col=0)
            
            # Clear existing conditional_formatting list
            sheet.conditional_formatting = ConditionalFormattingList()
            sheet.data_validations.dataValidation = []
            #df_track_cols = list(df_track.columns)
            df_track_cols = list(dft.columns)
            sheet = highlight_columns(sheet, columns=[df_track_cols.index(col) for col in df_track_cols] , color=colors[4])
            Problem_Summary = ['UNDEFINED',
                                'Big problem',
                                'Date of the CT images are wrong',
                                'Missing ICA Images',
                                'ICA Problem']
            

            Problem_SummaryStr = '"' + 'UNDEFINED,' + 'Big problem,' + 'Date of the CT images are wrong,' + 'Missing ICA Images,' + 'ICA Problem,' + '"'

            sheet = setAccessRights(sheet, columns=[df_track_cols.index(col) for col in ['Problem Summary']], promt='Problem!', promptTitle='DISCHARGE Problem Summary', formula1=Problem_SummaryStr)
            # Freeze worksheet
            workbook[sheetname].freeze_panes = "D2"
            # Draw border
            #sheet = setBorderFormula(sheet, formula='=$C1<>$C2', NumRows=df_master.shape[0], NumColumns=df_master.shape[1])
            # Add filter
            sheet.auto_filter.ref = sheet.dimensions

    writer.save()  

def updatePatient(folderpath_master, folderpath_master_before):
    
    print('Update patient table')
    
    createPatient(folderpath_master)
    
    date = folderpath_master.split('_')[-1]
    folderpath_components = os.path.join(folderpath_master, 'discharge_components_' + date)
    filepath_patient = os.path.join(folderpath_components, 'discharge_patient_' + date + '.xlsx')
    
    folderpath_master_before_list = glob(folderpath_master_before + '/*master*')
    folderpath_master_before_list = sortFolderpath(folderpath_master, folderpath_master_before_list)
    filepathMasters = glob(folderpath_master_before_list[-2] + '/*process*.xlsx')
    filepath_master = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')
    
    date_before = folderpath_master_before_list[-2].split('_')[-1]
    df_master = pd.read_excel(filepathMasters[0], sheet_name='PATIENT_STATUS_' + date_before, index_col=0)
    columns=['STATUS_MANUAL_CORRECTION', 'COMMENT']
    df_patient = pd.read_excel(filepath_patient, index_col=0)
    df_patient[columns] = df_master[columns]
    
    df_patient.to_excel(filepath_patient)
    
    
    
    # Add patient ro master
    sheet_name='PATIENT_STATUS' + '_' + date
    writer = pd.ExcelWriter(filepath_master, engine="openpyxl", mode="a")
    workbook  = writer.book
    sheet = workbook[sheet_name]
    workbook.remove(sheet)
    df_patient.to_excel(writer, sheet_name=sheet_name)
    writer.save()
    
# def createPatient(folderpath_master):
#     print('Create patient table.')

#     date = folderpath_master.split('_')[-1]
#     folderpath_components = os.path.join(folderpath_master, 'discharge_components_' + date)
#     filepath_pred = os.path.join(folderpath_components, 'discharge_pred_' + date + '.xlsx')
#     filepath_master_data = os.path.join(folderpath_components, 'discharge_master_data_' + date + '.xlsx')
#     filepath_patient = os.path.join(folderpath_components, 'discharge_patient_' + date + '.xlsx')
#     filepath_master = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')
    
#     df_master = pd.read_excel(filepath_master, sheet_name='MASTER_'+ date, index_col=0)
#     df_master_data = pd.read_excel(filepath_master_data, index_col=0)
#     #df_pred = pd.read_excel(filepath_pred, index_col=0)
#     #df_pred['CTAExtended'] = df_pred['CTAExtended'].astype('bool')
#     df_patient = pd.DataFrame(columns=['Site', 'PatientID', 'Modality', 'AcquisitionDate',
#                                        'CACS_NUM', 'CACS_FBP_NUM', 'CACS_IR_NUM', 
#                                        'CTA_NUM', 'CTA_FBP_NUM', 'CTA_IR_NUM', 
#                                        'NCS_CACS_NUM', 'NCS_CACS_FBP_NUM', 'NCS_CACS_IR_NUM', 
#                                        'NCS_CTA_NUM', 'NCS_CTA_FBP_NUM', 'NCS_CTA_IR_NUM',
#                                        'ITT', 'STATUS', 'STATUS_MANUAL_CORRECTION', 'COMMENT'])

#     patient_list = df_master_data['PatientID'].unique()
#     for patientID in patient_list:
#         #data = df_master_data[df_master_data['PatientID']==patientID]
#         #if data['Modality'].iloc[0]=='CT':
#         df_pat = df_master[df_master['PatientID']==patientID]
#         # Extract CACS_NUM information
#         CACS_MANUAL_OK = (df_pat['ClassManualCorrection']=='UNDEFINED') | (df_pat['ClassManualCorrection']=='CACS')
#         CACS_AUTO_OK = (df_pat['CACSExtended']) & (df_pat['ITT']<2)
#         CACS_NUM = (CACS_MANUAL_OK & CACS_AUTO_OK).sum()
#         # Extract CACS_FBP_NUM information
#         CACS_FBP_MANUAL_OK = (df_pat['ClassManualCorrection']=='UNDEFINED') | (df_pat['ClassManualCorrection']=='CACS')
#         CACS_FBP_AUTO_OK = (df_pat['CACSExtended'])  & (df_pat['RECO']=='FBP') & (df_pat['ITT']<2)
#         CACS_FBP_NUM = (CACS_MANUAL_OK & CACS_FBP_AUTO_OK).sum()
#         # Extract CACS_IR_NUM information
#         CACS_IR_MANUAL_OK = (df_pat['ClassManualCorrection']=='UNDEFINED') | (df_pat['ClassManualCorrection']=='CACS')
#         CACS_IR_AUTO_OK = (df_pat['CACSExtended']) & (df_pat['RECO']=='IR') & (df_pat['ITT']<2)
#         CACS_IR_NUM = (CACS_MANUAL_OK & CACS_IR_AUTO_OK).sum()    
#         # Extract CTA_NUM information
#         CTA_MANUAL_OK = (df_pat['ClassManualCorrection']=='UNDEFINED') | (df_pat['ClassManualCorrection']=='CTA')
#         CTA_AUTO_OK = (df_pat['CTAExtended']) & (df_pat['ITT']<2)
#         CTA_NUM = (CTA_MANUAL_OK & CTA_AUTO_OK).sum()   
#         # Extract CTA_FBP_NUM information
#         CTA_FBP_MANUAL_OK = (df_pat['ClassManualCorrection']=='UNDEFINED') | (df_pat['ClassManualCorrection']=='CTA')
#         CTA_FBP_AUTO_OK = (df_pat['CTAExtended'])  & (df_pat['RECO']=='FBP') & (df_pat['ITT']<2)
#         CTA_FBP_NUM = (CTA_FBP_MANUAL_OK & CTA_FBP_AUTO_OK).sum()  
#         # Extract CTA_IR_NUM information
#         CTA_IR_MANUAL_OK = (df_pat['ClassManualCorrection']=='UNDEFINED') | (df_pat['ClassManualCorrection']=='CTA')
#         CTA_IR_AUTO_OK = (df_pat['CTAExtended'])  & (df_pat['RECO']=='FBP') & (df_pat['ITT']<2)
#         CTA_IR_NUM = (CTA_IR_MANUAL_OK & CTA_IR_AUTO_OK).sum()          
#         # Extract NCS_CACS_NUM information
#         NCS_CACS_MANUAL_OK = (df_pat['ClassManualCorrection']=='UNDEFINED') | (df_pat['ClassManualCorrection']=='NCS_CACS')
#         NCS_CACS_AUTO_OK = (df_pat['NCS_CACSExtended']) & (df_pat['ITT']<2)
#         NCS_CACS_NUM = (NCS_CACS_MANUAL_OK & NCS_CACS_AUTO_OK).sum()       
#         # Extract NCS_CACS_FBP_NUM information
#         NCS_CACS_FBP_MANUAL_OK = (df_pat['ClassManualCorrection']=='UNDEFINED') | (df_pat['ClassManualCorrection']=='NCS_CACS')
#         NCS_CACS_FBP_AUTO_OK = (df_pat['NCS_CACSExtended']) & (df_pat['RECO']=='FBP') & (df_pat['ITT']<2)
#         NCS_CACS_FBP_NUM = (NCS_CACS_FBP_MANUAL_OK & NCS_CACS_FBP_AUTO_OK).sum()   
#         # Extract NCS_CACS_IR_NUM information
#         NCS_CACS_IR_MANUAL_OK = (df_pat['ClassManualCorrection']=='UNDEFINED') | (df_pat['ClassManualCorrection']=='NCS_CACS')
#         NCS_CACS_IR_AUTO_OK = (df_pat['NCS_CACSExtended']) & (df_pat['RECO']=='IR') & (df_pat['ITT']<2)
#         NCS_CACS_IR_NUM = (NCS_CACS_IR_MANUAL_OK & NCS_CACS_IR_AUTO_OK).sum()   
#         # Extract NCS_CACS_NUM information
#         NCS_CTA_MANUAL_OK = (df_pat['ClassManualCorrection']=='UNDEFINED') | (df_pat['ClassManualCorrection']=='NCS_CTA')
#         NCS_CTA_AUTO_OK = (df_pat['NCS_CTAExtended']) & (df_pat['ITT']<2)
#         NCS_CTA_NUM = (NCS_CTA_MANUAL_OK & NCS_CTA_AUTO_OK).sum()  
#         # Extract NCS_CTA_FBP_NUM information
#         NCS_CTA_FBP_MANUAL_OK = (df_pat['ClassManualCorrection']=='UNDEFINED') | (df_pat['ClassManualCorrection']=='NCS_CTA')
#         NCS_CTA_FBP_AUTO_OK = (df_pat['NCS_CTAExtended']) & (df_pat['RECO']=='FBP') & (df_pat['ITT']<2)
#         NCS_CTA_FBP_NUM = (NCS_CTA_FBP_MANUAL_OK & NCS_CTA_FBP_AUTO_OK).sum() 
#         # Extract NCS_CTA_IR_NUM information
#         NCS_CTA_IR_MANUAL_OK = (df_pat['ClassManualCorrection']=='UNDEFINED') | (df_pat['ClassManualCorrection']=='NCS_CTA')
#         NCS_CTA_IR_AUTO_OK = (df_pat['NCS_CTAExtended']) & (df_pat['RECO']=='IR') & (df_pat['ITT']<2)
#         NCS_CTA_IR_NUM = (NCS_CTA_IR_MANUAL_OK & NCS_CTA_IR_AUTO_OK).sum() 
        

#         SITE = df_pat['Site'].iloc[0]
#         ITT = df_pat['ITT'].iloc[0]
#         modality = df_pat['Modality'].iloc[0]
#         #AcquisitionDate = df_pat['AcquisitionDate'].iloc[0]
        
#         # Check patient scenario
#         CACS_OK = CACS_NUM>=2
#         CTA_OK = CTA_NUM>=2
#         NCS_CACS_OK = NCS_CACS_NUM>=2
#         NCS_CTA_OK = NCS_CTA_NUM>=2
        
#         if CACS_OK and CTA_OK and NCS_CACS_OK and NCS_CTA_OK:
#             status = 'OK'
#         elif ITT==2:
#             status = 'EXCLUDED'
#         elif not modality=='CT':
#             status = 'NOT CT MODALITY'
#         else:
#             status=''
#             if not CACS_OK:
#                 status = status + 'MISSING_CACS, '
#             if not CTA_OK:
#                 status = status + 'MISSING_CTA, '
#             if not NCS_CACS_OK:
#                 status = status + 'MISSING_NCS_CACS, '
#             if not NCS_CTA_OK:
#                 status = status + 'MISSING_NCS_CTA, '
            
#         STATUS_MANUAL_CORRECTION = 'UNDEFINED'
#         COMMENT = ''
#         df_patient = df_patient.append({'Site': SITE, 'PatientID': patientID, 'Modality': modality, 'CACS_NUM': CACS_NUM, 'CACS_FBP_NUM': CACS_FBP_NUM, 'CACS_IR_NUM': CACS_IR_NUM,
#                            'CTA_NUM': CTA_NUM, 'CTA_FBP_NUM': CTA_FBP_NUM, 'CTA_IR_NUM': CTA_IR_NUM,
#                            'NCS_CACS_NUM': NCS_CACS_NUM, 'NCS_CACS_FBP_NUM': NCS_CACS_FBP_NUM, 'NCS_CACS_IR_NUM': NCS_CACS_IR_NUM,
#                            'NCS_CTA_NUM': NCS_CTA_NUM, 'NCS_CTA_FBP_NUM': NCS_CTA_FBP_NUM, 'NCS_CTA_IR_NUM': NCS_CTA_IR_NUM,
#                            'ITT': ITT, 'STATUS': status, 'STATUS_MANUAL_CORRECTION': STATUS_MANUAL_CORRECTION, 'COMMENT': COMMENT}, ignore_index=True)
    
#     df_patient.to_excel(filepath_patient)
    
#     # Remove sheet if already exist
#     sheet_name = 'PATIENT_STATUS' + '_' + date
#     workbook  = writer.book
#     sheetnames = workbook.sheetnames
#     if sheet_name in sheetnames:
#         sheet = workbook[sheet_name]
#         workbook.remove(sheet)
    
#     # Add patient ro master
#     writer = pd.ExcelWriter(filepath_master, engine="openpyxl", mode="a")
#     df_patient.to_excel(writer, sheet_name=sheet_name)
#     writer.save()


def createStudy(settings):
    print('Create StudyInstanceID table.')
    
    conf=True
    
    CACS_NUM_MIN = 1
    CACS_IR_NUM_MIN = 0
    CACS_FBP_NUM_MIN = 0
    
    CTA_NUM_MIN = 1
    CTA_IR_NUM_MIN = 0
    CTA_FBP_NUM_MIN = 0
    
    NCS_CACS_NUM_MIN = 1
    NCS_CACS_IR_NUM_MIN = 0
    NCS_CACS_FBP_NUM_MIN = 0
    
    NCS_CTA_NUM_MIN = 1
    NCS_CTA_IR_NUM_MIN = 0
    NCS_CTA_FBP_NUM_MIN = 0

    filepath_patient = settings['filepath_patient']
    df_master = pd.read_excel(settings['filepath_master'], sheet_name='MASTER_'+ settings['date'], index_col=0)
    df_master.sort_index(inplace=True)
    #df_master_data = pd.read_excel(filepath_master_data, index_col=0)
    df_PatientID = pd.DataFrame(columns=['Site', 'PatientID', 'StudyInstanceUID', 'Modality', 'AcquisitionDate',
                                       'CACS_NUM', 'CACS_FBP_NUM', 'CACS_IR_NUM', 
                                       'CTA_NUM', 'CTA_FBP_NUM', 'CTA_IR_NUM', 
                                       'NCS_CACS_NUM', 'NCS_CACS_FBP_NUM', 'NCS_CACS_IR_NUM', 
                                       'NCS_CTA_NUM', 'NCS_CTA_FBP_NUM', 'NCS_CTA_IR_NUM',
                                       'ITT', 'STATUS', 'STATUS_MANUAL_CORRECTION', 'COMMENT'])
    df_PatientID_conf = pd.DataFrame(columns=['Site', 'PatientID', 'StudyInstanceUID', 'Modality', 'AcquisitionDate',
                                       'CACS_CONF', 'CACS_FBP_CONF', 'CACS_IR_CONF', 
                                       'CTA_CONF', 'CTA_FBP_CONF', 'CTA_IR_CONF', 
                                       'NCS_CACS_CONF', 'NCS_CACS_FBP_CONF', 'NCS_CACS_IR_CONF', 
                                       'NCS_CTA_CONF', 'NCS_CTA_FBP_CONF', 'NCS_CTA_IR_CONF',
                                       'ITT', 'STATUS', 'STATUS_MANUAL_CORRECTION', 'COMMENT'])

    # Filter study list
    func = lambda x: datetime.strptime(x, '%Y%m%d')
    patients = df_master['PatientID'].unique()
    firstdateList = []
    study_list = []
    
    #patients=['06-GOE-0020']
    
    
    for patient in patients:
        df_patient = df_master[(df_master['PatientID']==patient) &  (df_master['Modality']=='CT')]
        if len(df_patient)>0:
            #sys.exit()
            firstdate = df_patient['1. Date of CT scan'].iloc[0]
            
            studydate = df_patient['StudyDate']
            # Convert string to date and replace in df_patient
            studydate = studydate.apply(lambda x: datetime.datetime.strptime(str(x), '%Y%m%d'))
            df_patient['StudyDate'] = studydate
            
            if not pd.isnull(firstdate):
                df_study = df_patient[studydate == firstdate]
                if len(df_study)>0:
                    #df_study_id = df_study['StudyInstanceUID'].iloc[0]
                    df_study_id_list = df_study['StudyInstanceUID']
                else:
                    print('PROBLEM: Patient ' + patient + ' "1. Date of CT scan" not consistent with "StudyDate"')
                    df_study = df_patient
                    df_study = df_study.sort_values(by='StudyDate')
                    df_study_id = df_study['StudyInstanceUID'].iloc[0]
            else:
                print('Patient: ' + patient + ' does not have a 1. Date of CT scan')
                df_study = df_patient
                df_study = df_study.sort_values(by='StudyDate')
                #df_study_id = df_study['StudyInstanceUID'].iloc[0]
                df_study_id_list = [df_study['StudyInstanceUID'].iloc[0]]
            study_list.append(df_study_id_list)
            
    def getConf(conf, NumSeries=1):
        conf = conf.sort_values(ascending=False)
        if len(conf)==0:
            return 0
        if len(conf) <= NumSeries:
            return conf.min()
        else:
            return conf[0:NumSeries].min()
       

    for studyIDList in study_list:
        # if studyID == '1.2.840.113619.6.95.31.0.3.4.1.1018.13.10856850':
        #     sys.exit()
        #data = df_master_data[df_master_data['PatientID']==patientID]
        #if data['Modality'].iloc[0]=='CT':
        df_study = df_master[df_master['StudyInstanceUID'].isin(studyIDList)]
        # Extract CACS_NUM information
        CACS_MANUAL_OK = (df_study['ClassManualCorrection']=='UNDEFINED') | (df_study['ClassManualCorrection']=='CACS')
        #CACS_AUTO_OK = (df_study['CACSExtended']) & (df_study['ITT']<2)
        CACS_AUTO_OK = (df_study['RFCClass']=='CACSExtended') & (df_study['ITT']<2)
        CACS_NUM = (CACS_MANUAL_OK & CACS_AUTO_OK).sum()
        CACS_CONF = getConf(df_study[CACS_MANUAL_OK & CACS_AUTO_OK]['RFCConfidence'], NumSeries=CACS_NUM_MIN)
        
        # Extract CACS_FBP_NUM information
        CACS_FBP_MANUAL_OK = (df_study['ClassManualCorrection']=='UNDEFINED') | (df_study['ClassManualCorrection']=='CACS')
        #CACS_FBP_AUTO_OK = (df_study['CACSExtended'])  & (df_study['RECO']=='FBP') & (df_study['ITT']<2)
        CACS_FBP_AUTO_OK = (df_study['RFCClass']=='CACSExtended')  & (df_study['RECO']=='FBP') & (df_study['ITT']<2)
        CACS_FBP_NUM = (CACS_FBP_MANUAL_OK & CACS_FBP_AUTO_OK).sum()
        CACS_FBP_CONF = getConf(df_study[CACS_MANUAL_OK & CACS_FBP_AUTO_OK]['RFCConfidence'], NumSeries=CACS_FBP_NUM_MIN)
        # Extract CACS_IR_NUM information
        CACS_IR_MANUAL_OK = (df_study['ClassManualCorrection']=='UNDEFINED') | (df_study['ClassManualCorrection']=='CACS')
        #CACS_IR_AUTO_OK = (df_study['CACSExtended']) & (df_study['RECO']=='IR') & (df_study['ITT']<2)
        CACS_IR_AUTO_OK = (df_study['RFCClass']=='CACSExtended') & (df_study['RECO']=='IR') & (df_study['ITT']<2)
        CACS_IR_NUM = (CACS_IR_MANUAL_OK & CACS_IR_AUTO_OK).sum()
        CACS_IR_CONF = getConf(df_study[CACS_MANUAL_OK & CACS_IR_AUTO_OK]['RFCConfidence'], NumSeries=CACS_IR_NUM_MIN)
        # Extract CTA_NUM information
        CTA_MANUAL_OK = (df_study['ClassManualCorrection']=='UNDEFINED') | (df_study['ClassManualCorrection']=='CTA')
        #CTA_AUTO_OK = (df_study['CTAExtended']) & (df_study['ITT']<2)
        CTA_AUTO_OK = (df_study['RFCClass']=='CTAExtended') & (df_study['ITT']<2)
        CTA_NUM = (CTA_MANUAL_OK & CTA_AUTO_OK).sum()  
        CTA_CONF = getConf(df_study[CTA_MANUAL_OK & CTA_AUTO_OK]['RFCConfidence'], NumSeries=CTA_NUM_MIN)
        # Extract CTA_FBP_NUM information
        CTA_FBP_MANUAL_OK = (df_study['ClassManualCorrection']=='UNDEFINED') | (df_study['ClassManualCorrection']=='CTA')
        #CTA_FBP_AUTO_OK = (df_study['CTAExtended'])  & (df_study['RECO']=='FBP') & (df_study['ITT']<2)
        CTA_FBP_AUTO_OK = (df_study['RFCClass']=='CTAExtended')  & (df_study['RECO']=='FBP') & (df_study['ITT']<2)
        CTA_FBP_NUM = (CTA_FBP_MANUAL_OK & CTA_FBP_AUTO_OK).sum()  
        CTA_FBP_CONF = getConf(df_study[CTA_FBP_MANUAL_OK & CTA_FBP_AUTO_OK]['RFCConfidence'], NumSeries=CTA_FBP_NUM_MIN)
        # Extract CTA_IR_NUM information
        CTA_IR_MANUAL_OK = (df_study['ClassManualCorrection']=='UNDEFINED') | (df_study['ClassManualCorrection']=='CTA')
        #CTA_IR_AUTO_OK = (df_study['CTAExtended'])  & (df_study['RECO']=='FBP') & (df_study['ITT']<2)
        CTA_IR_AUTO_OK = (df_study['RFCClass']=='CTAExtended')  & (df_study['RECO']=='FBP') & (df_study['ITT']<2)
        CTA_IR_NUM = (CTA_IR_MANUAL_OK & CTA_IR_AUTO_OK).sum()          
        CTA_IR_CONF = getConf(df_study[CTA_IR_MANUAL_OK & CTA_IR_AUTO_OK]['RFCConfidence'], NumSeries=CTA_IR_NUM_MIN)
        # Extract NCS_CACS_NUM information
        NCS_CACS_MANUAL_OK = (df_study['ClassManualCorrection']=='UNDEFINED') | (df_study['ClassManualCorrection']=='NCS_CACS')
        #NCS_CACS_AUTO_OK = (df_study['NCS_CACSExtended']) & (df_study['ITT']<2)
        NCS_CACS_AUTO_OK = (df_study['RFCClass']=='NCS_CACSExtended') & (df_study['ITT']<2)
        NCS_CACS_NUM = (NCS_CACS_MANUAL_OK & NCS_CACS_AUTO_OK).sum()    
        NCS_CACS_CONF = getConf(df_study[NCS_CACS_MANUAL_OK & NCS_CACS_AUTO_OK]['RFCConfidence'], NumSeries=NCS_CACS_NUM_MIN)
        # Extract NCS_CACS_FBP_NUM information
        NCS_CACS_FBP_MANUAL_OK = (df_study['ClassManualCorrection']=='UNDEFINED') | (df_study['ClassManualCorrection']=='NCS_CACS')
        #NCS_CACS_FBP_AUTO_OK = (df_study['NCS_CACSExtended']) & (df_study['RECO']=='FBP') & (df_study['ITT']<2)
        NCS_CACS_FBP_AUTO_OK = (df_study['RFCClass']=='NCS_CACSExtended') & (df_study['RECO']=='FBP') & (df_study['ITT']<2)
        NCS_CACS_FBP_NUM = (NCS_CACS_FBP_MANUAL_OK & NCS_CACS_FBP_AUTO_OK).sum()  
        NCS_CACS_FBP_CONF = getConf(df_study[NCS_CACS_FBP_MANUAL_OK & NCS_CACS_FBP_AUTO_OK]['RFCConfidence'], NumSeries=NCS_CACS_FBP_NUM_MIN)
        # Extract NCS_CACS_IR_NUM information
        NCS_CACS_IR_MANUAL_OK = (df_study['ClassManualCorrection']=='UNDEFINED') | (df_study['ClassManualCorrection']=='NCS_CACS')
        #NCS_CACS_IR_AUTO_OK = (df_study['NCS_CACSExtended']) & (df_study['RECO']=='IR') & (df_study['ITT']<2)
        NCS_CACS_IR_AUTO_OK = (df_study['RFCClass']=='NCS_CACSExtended') & (df_study['RECO']=='IR') & (df_study['ITT']<2)
        NCS_CACS_IR_NUM = (NCS_CACS_IR_MANUAL_OK & NCS_CACS_IR_AUTO_OK).sum()   
        NCS_CACS_IR_CONF = getConf(df_study[NCS_CACS_IR_MANUAL_OK & NCS_CACS_IR_AUTO_OK]['RFCConfidence'], NumSeries=NCS_CACS_IR_NUM_MIN)
        # Extract NCS_CACS_NUM information
        NCS_CTA_MANUAL_OK = (df_study['ClassManualCorrection']=='UNDEFINED') | (df_study['ClassManualCorrection']=='NCS_CTA')
        #NCS_CTA_AUTO_OK = (df_study['NCS_CTAExtended']) & (df_study['ITT']<2)
        NCS_CTA_AUTO_OK = (df_study['RFCClass']=='NCS_CTAExtended') & (df_study['ITT']<2)
        NCS_CTA_NUM = (NCS_CTA_MANUAL_OK & NCS_CTA_AUTO_OK).sum() 
        NCS_CTA_CONF = getConf(df_study[NCS_CTA_MANUAL_OK & NCS_CTA_AUTO_OK]['RFCConfidence'], NumSeries=NCS_CTA_NUM_MIN)
        # Extract NCS_CTA_FBP_NUM information
        NCS_CTA_FBP_MANUAL_OK = (df_study['ClassManualCorrection']=='UNDEFINED') | (df_study['ClassManualCorrection']=='NCS_CTA')
        #NCS_CTA_FBP_AUTO_OK = (df_study['NCS_CTAExtended']) & (df_study['RECO']=='FBP') & (df_study['ITT']<2)
        NCS_CTA_FBP_AUTO_OK = (df_study['RFCClass']=='NCS_CTAExtended') & (df_study['RECO']=='FBP') & (df_study['ITT']<2)
        NCS_CTA_FBP_NUM = (NCS_CTA_FBP_MANUAL_OK & NCS_CTA_FBP_AUTO_OK).sum() 
        NCS_CTA_FBP_CONF = getConf(df_study[NCS_CTA_FBP_MANUAL_OK & NCS_CTA_FBP_AUTO_OK]['RFCConfidence'], NumSeries=NCS_CTA_FBP_NUM_MIN)
        # Extract NCS_CTA_IR_NUM information
        NCS_CTA_IR_MANUAL_OK = (df_study['ClassManualCorrection']=='UNDEFINED') | (df_study['ClassManualCorrection']=='NCS_CTA')
        #NCS_CTA_IR_AUTO_OK = (df_study['NCS_CTAExtended']) & (df_study['RECO']=='IR') & (df_study['ITT']<2)
        NCS_CTA_IR_AUTO_OK = (df_study['RFCClass']=='NCS_CTAExtended') & (df_study['RECO']=='IR') & (df_study['ITT']<2)
        NCS_CTA_IR_NUM = (NCS_CTA_IR_MANUAL_OK & NCS_CTA_IR_AUTO_OK).sum() 
        NCS_CTA_IR_CONF = getConf(df_study[NCS_CTA_IR_MANUAL_OK & NCS_CTA_IR_AUTO_OK]['RFCConfidence'], NumSeries=NCS_CTA_IR_NUM_MIN)
        
        PATIENTID = df_study['PatientID'].iloc[0]
        SITE = df_study['Site'].iloc[0]
        ITT = df_study['ITT'].iloc[0]
        modality = df_study['Modality'].iloc[0]
        DATE = df_study['AcquisitionDate'].iloc[0]
        STUDYID = ','.join(studyIDList)
        #AcquisitionDate = df_study['AcquisitionDate'].iloc[0]
        
        # Check patient scenario
        CACS_OK = CACS_NUM>=CACS_NUM_MIN
        CTA_OK = CTA_NUM>=CTA_NUM_MIN
        NCS_CACS_OK = NCS_CACS_NUM>=NCS_CACS_NUM_MIN
        NCS_CTA_OK = NCS_CTA_NUM>=NCS_CTA_NUM_MIN
        
        if CACS_OK and CTA_OK and NCS_CACS_OK and NCS_CTA_OK:
            status = 'OK'
        elif ITT==2:
            status = 'EXCLUDED'
        elif not modality=='CT':
            status = 'NOT CT MODALITY'
        else:
            status=''
            if not CACS_OK:
                status = status + 'MISSING_CACS, '
            if not CTA_OK:
                status = status + 'MISSING_CTA, '
            if not NCS_CACS_OK:
                status = status + 'MISSING_NCS_CACS, '
            if not NCS_CTA_OK:
                status = status + 'MISSING_NCS_CTA, '
            
        STATUS_MANUAL_CORRECTION = 'UNDEFINED'
        COMMENT = ''
        df_PatientID = df_PatientID.append({'Site': SITE, 'PatientID': PATIENTID, 'StudyInstanceUID': STUDYID, 'Modality': modality, 'AcquisitionDate': DATE, 'CACS_NUM': CACS_NUM, 'CACS_FBP_NUM': CACS_FBP_NUM, 'CACS_IR_NUM': CACS_IR_NUM,
                           'CTA_NUM': CTA_NUM, 'CTA_FBP_NUM': CTA_FBP_NUM, 'CTA_IR_NUM': CTA_IR_NUM,
                           'NCS_CACS_NUM': NCS_CACS_NUM, 'NCS_CACS_FBP_NUM': NCS_CACS_FBP_NUM, 'NCS_CACS_IR_NUM': NCS_CACS_IR_NUM,
                           'NCS_CTA_NUM': NCS_CTA_NUM, 'NCS_CTA_FBP_NUM': NCS_CTA_FBP_NUM, 'NCS_CTA_IR_NUM': NCS_CTA_IR_NUM,
                           'ITT': ITT, 'STATUS': status, 'STATUS_MANUAL_CORRECTION': STATUS_MANUAL_CORRECTION, 'COMMENT': COMMENT}, ignore_index=True)
        df_PatientID_conf = df_PatientID_conf.append({'Site': SITE, 'PatientID': PATIENTID, 'StudyInstanceUID': STUDYID, 'Modality': modality, 'AcquisitionDate': DATE, 'CACS_CONF': CACS_CONF, 'CACS_FBP_CONF': CACS_FBP_CONF, 'CACS_IR_CONF': CACS_IR_CONF,
                           'CTA_CONF': CTA_CONF, 'CTA_FBP_CONF': CTA_FBP_CONF, 'CTA_IR_CONF': CTA_IR_CONF,
                           'NCS_CACS_CONF': NCS_CACS_CONF, 'NCS_CACS_FBP_CONF': NCS_CACS_FBP_CONF, 'NCS_CACS_IR_CONF': NCS_CACS_IR_CONF,
                           'NCS_CTA_CONF': NCS_CTA_CONF, 'NCS_CTA_FBP_CONF': NCS_CTA_FBP_CONF, 'NCS_CTA_IR_CONF': NCS_CTA_IR_CONF,
                           'ITT': ITT, 'STATUS': status, 'STATUS_MANUAL_CORRECTION': STATUS_MANUAL_CORRECTION, 'COMMENT': COMMENT}, ignore_index=True)
        
    df_PatientID.to_excel(filepath_patient)
    df_PatientID_conf.to_excel(settings['filepath_patient_conf'])

    writer = pd.ExcelWriter(settings['filepath_master'], engine="openpyxl", mode="a")
    # Remove sheet if already exist
    sheet_name = 'PATIENT_STATUS_' + settings['date']
    workbook  = writer.book
    sheetnames = workbook.sheetnames
    if sheet_name in sheetnames:
        sheet = workbook[sheet_name]
        workbook.remove(sheet)
    df_PatientID.to_excel(writer, sheet_name=sheet_name)
        
    if conf:
        sheet_name = 'PATIENT_STATUS_CONF_' + settings['date']
        if sheet_name in sheetnames:
            sheet = workbook[sheet_name]
            workbook.remove(sheet)
        df_PatientID_conf.to_excel(writer, sheet_name=sheet_name)
        
    # Add patient ro master
    
    writer.save()
    
def extractHist(settings):

    df = pd.read_excel(settings['filepath_master'], sheet_name ='MASTER_' + settings['date'], index_col=0)
    df.sort_index(inplace=True)
    bins = 100   
    columns =['SeriesInstanceUID', 'Count', 'CLASS'] + [str(x) for x in range(0,bins)]
    if os.path.exists(settings['filepath_hist']):
        dfHist = pd.read_pickle(settings['filepath_hist'])
    else:
        dfHist = pd.DataFrame('', index=np.arange(len(df)), columns=columns)
    
    start = 0
    end = len(df)
    #end = 1200
    for index, row in df[start:end].iterrows():
        print('index', index)   
        if dfHist.iloc[index,0]=='':
            if keyboard.is_pressed('ctrl+e'):
                print('Button "ctrl + e" pressed to exit execution.')
                sys.exit()
            StudyInstanceUID=row['StudyInstanceUID']
            PatientID=row['PatientID']
            SeriesInstanceUID=row['SeriesInstanceUID']
            #SOPInstanceUID=row['SOPInstanceUID']

            if not (row['Modality']=='CT' or row['Modality']=='OT'):
                print('Series modality is not CT')
                dfHist.loc[index,'SeriesInstanceUID'] = SeriesInstanceUID
                dfHist.loc[index,'Count'] = -1
                dfHist.loc[index,'CLASS'] = row['CLASS']
                dfHist.iloc[index,3:] = np.ones((1, bins))*-1
            else:
                try:
                    patient=CTPatient(StudyInstanceUID, PatientID)
                    series = patient.loadSeries(settings['folderpath_discharge'], SeriesInstanceUID, None)
                    image = series.image.image()
                    hist = np.histogram(image, bins=bins, range=(-2500, 3000))[0]
                    dfHist.loc[index,'SeriesInstanceUID'] = SeriesInstanceUID
                    dfHist.loc[index,'Count'] = image.shape[0]
                    dfHist.loc[index,'CLASS'] = row['CLASS']
                    dfHist.iloc[index,3:] = hist
                except:
                    print('Error index', index)
                    dfHist.loc[index,'SeriesInstanceUID'] = SeriesInstanceUID
                    dfHist.loc[index,'Count'] = -1
                    dfHist.loc[index,'CLASS'] = row['CLASS']
                    dfHist.iloc[index,3:] = np.ones((1, bins))*-1
        if index % 10 == 0:
            dfHist.to_pickle(settings['filepath_hist'])
    dfHist.to_pickle(settings['filepath_hist'])

def checkFileSize():

    filepath = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_master/discharge_master_01042020/discharge_master_01042020.xlsx'
    filepath_failed = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_master/discharge_master_01042020/discharge_master_01042020_failed.pkl'
    filepath_failed_excel = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_master/discharge_master_01042020/discharge_master_01042020_failed.xlsx'
    folderpath_images = 'G:/discharge'
    df = pd.read_excel(filepath, sheet_name ='MASTER_01042020', index_col=0)
    df=df[df['Modality']=='CT']   
    #df=df[df['PatientID']=='11-CAG-0015']   
    
    
    df.reset_index(inplace=True)
    df.sort_index(inplace=True)
    
    print('NumSamples: ', len(df))
    df_failed = pd.DataFrame(columns=df.columns)

    start = 0
    end = len(df)
    for index, row in df[start:end].iterrows():
        print('index', index)   
        StudyInstanceUID=row['StudyInstanceUID']
        PatientID=row['PatientID']
        SeriesInstanceUID=row['SeriesInstanceUID']
        if keyboard.is_pressed('ctrl+e'):
            print('Button "ctrl + e" pressed to exit execution.')
            sys.exit()
            
        filepathSeries = os.path.join(folderpath_images,StudyInstanceUID, SeriesInstanceUID)
        files = glob(filepathSeries + '/*.dcm')
        fileSize = False
        for file in files:
            kB = os.path.getsize(file)/1024
            if kB>10:
                fileSize=True
                break
        if fileSize==False:
            df_failed = df_failed.append(row)
            df_failed.to_pickle(filepath_failed)
            print('Index: ', index, ' failed.')
            print('SeriesInstanceUID', SeriesInstanceUID)   

    df_failed = pd.read_pickle(filepath_failed)

    # Save to excel file
    writer = pd.ExcelWriter(filepath_failed_excel, engine="openpyxl", mode="w")
    df_failed.to_excel(writer, sheet_name='FileSizeFailed')
    writer.save()
    
def checkMultiSlice():

    filepath = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_master/discharge_master_01042020/discharge_master_01042020.xlsx'
    filepath_multi = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_master/discharge_master_01042020/discharge_master_01042020_multi.pkl'
    filepath_multi_excel = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_master/discharge_master_01042020/discharge_master_01042020_multi.xlsx'
    folderpath_images = 'G:/discharge'
    df = pd.read_excel(filepath, sheet_name ='MASTER_01042020', index_col=0)
    df=df[(df['Modality']=='CT') | (df['Modality']=='OT')]   
    #df=df[df['PatientID']=='11-CAG-0015']   
    
    
    df.reset_index(inplace=True)
    df.sort_index(inplace=True)
    
    print('NumSamples: ', len(df))
    df_multi = pd.DataFrame(columns=df.columns)

    start = 0
    end = len(df)
    for index, row in df[start:end].iterrows():
        print('index', index)   
        StudyInstanceUID=row['StudyInstanceUID']
        PatientID=row['PatientID']
        SeriesInstanceUID=row['SeriesInstanceUID']
        if keyboard.is_pressed('ctrl+e'):
            print('Button "ctrl + e" pressed to exit execution.')
            sys.exit()
            
        filepathSeries = os.path.join(folderpath_images,StudyInstanceUID, SeriesInstanceUID)
        files = glob(filepathSeries + '/*.dcm')
        fileSize = 0
        for file in files:
            MB = os.path.getsize(file)/1024/1024
            if MB>15:
                fileSize=fileSize+1
        if fileSize>1:
            print('fileSize', fileSize) 
            df_multi = df_multi.append(row)
            df_multi.to_pickle(filepath_multi)
            print('Index: ', index, ' failed.')
            print('SeriesInstanceUID', SeriesInstanceUID)   

    df_multi = pd.read_pickle(filepath_multi)

    # Save to excel file
    writer = pd.ExcelWriter(filepath_multi_excel, engine="openpyxl", mode="w")
    df_multi.to_excel(writer, sheet_name='FileSizeFailed')
    writer.save()

def mergeManualSelection(settings):
    """ Merge manual selected series
        
    :param settings: Dictionary of settings
    :type settings: dict
    """
    
    filepath_master = settings['filepath_master']
    discharge_manual_selection = settings['folderpath_manual_selection']
    df_master = pd.read_excel(filepath_master, sheet_name='MASTER_'+ settings['date'], index_col=0)
    
    # Extract manual files
    centers = defaultdict(lambda:None, {})
    filepath_manual = glob(discharge_manual_selection + '/*.xlsx')
    for file in filepath_manual:
        filesplit = file[0:-5].split("_")
        for st in filesplit:
            if st[0]=='P' and len(st)==3:
                centers[st] = file
    
    # Update master for each center
    for center in centers.keys():
        print('Processing center', center)
        filepath_manual = centers[center]
        columns_copy=['ClassManualCorrection', 'Comment', 'Responsible Person']
        df_manual = pd.read_excel(filepath_manual, index_col=0)
        df_manual.replace(to_replace=[np.nan], value='', inplace=True)
        df_manual_P = df_manual[df_manual['Site']==center]
        df_manual_P = df_manual_P[columns_copy]
        df_manual_P = df_manual_P[~(df_manual_P['ClassManualCorrection']=='UNDEFINED')]
        df_manual_P['RFCLabel'] = df_manual_P['ClassManualCorrection']
        df_master.update(df_manual_P)
    
    # Save master
    writer = pd.ExcelWriter(filepath_master, engine="openpyxl", mode="a")
    sheet_name = 'MASTER_'+ settings['date']
    workbook  = writer.book
    sheetnames = workbook.sheetnames
    if sheet_name in sheetnames:
        sheet = workbook[sheet_name]
        workbook.remove(sheet)
    df_master.to_excel(writer, sheet_name=sheet_name)
    writer.save()

def checkAutomaticManual(settings):

    filepath_master = settings['filepath_master']
    df_master = pd.read_excel(filepath_master, sheet_name='MASTER_'+ settings['date'], index_col=0)
    df_master_copy = df_master.copy()
    df_master_copy.replace(to_replace=[np.nan], value=0.0, inplace=True)
    
    
    for index, row in df_master.iterrows():
        print('index', index)
        # Check multiple multi-slice-scans
        if (df_master_copy.loc[index, 'NumberOfFrames'] > 1) and (df_master_copy.loc[index, 'Count']>1):
            df_master.loc[index, 'ClassManualCorrection'] = 'PROBLEM'
            df_master.loc[index, 'Comment'] = 'Multiple Multi-Slice-CTs under one SeriesInstanceUID.'
            df_master.loc[index, 'Responsible Person'] = 'BF_AUT'
        # Define other selection based on count smaller 15
        if (df_master_copy.loc[index, 'NumberOfFrames'] == 0.0) and (df_master_copy.loc[index, 'Count']<15):
            df_master.loc[index, 'ClassManualCorrection'] = 'OTHER'
            df_master.loc[index, 'Comment'] = 'Selected as other becaus number of slices is smaller 15'
            df_master.loc[index, 'Responsible Person'] = 'BF_AUT'
            
    # Save master
    writer = pd.ExcelWriter(filepath_master, engine="openpyxl", mode="a")
    sheet_name = 'MASTER_'+ settings['date']
    workbook  = writer.book
    sheetnames = workbook.sheetnames
    if sheet_name in sheetnames:
        sheet = workbook[sheet_name]
        workbook.remove(sheet)
    df_master.to_excel(writer, sheet_name=sheet_name)
    writer.save()
    
    
def createMaster():
    """ Create master file
    """
    
    # Load settings
    filepath_settings = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/data/settings.json'
    settings=initSettings()
    saveSettings(settings, filepath_settings)
    settings = fillSettingsTags(loadSettings(filepath_settings))
    
    # Extract histograms
    #extractHist(settings)
    # Extract dicom tags
    #extractDICOMTags(settings, NumSamples=10)
    # Create tables
    checkTables(settings)
    # Create data
    createData(settings)
    # Create random forest classification columns
    createRFClassification(settings)
    # Create manual selection
    createManualSelection(settings)
    # Create prediction 
    createPredictions(settings)
    # Merge master 
    mergeMaster(settings)
    # Create tracking table 
    createTrackingTable(settings)
    updateMasterFromTrackingTable(settings)
    # Init RF classifier
    initRFClassification(settings)
    classifieRFClassification(settings)
    # Update manual selection
    mergeManualSelection(settings)
    #Classifie based on manual selection
    classifieRFClassification(settings)
    # Check manual selection
    checkAutomaticManual(settings)
    # Filter according to 10StepsGuide
    filer10StepsGuide(settings)
    # Merge study sheet 
    createStudy(settings)
    # Format master
    formatMaster(settings)

