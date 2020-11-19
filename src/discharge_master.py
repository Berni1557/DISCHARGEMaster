# -*- coding: utf-8 -*-
"""
Created on Wed May 13 13:59:31 2020

@author: bernifoellmer
"""

import sys, os
import pandas as pd
import openpyxl
import ntpath
import datetime
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.styles import Font, Color, Border, Side
from openpyxl.styles import colors
from openpyxl.styles import Protection
from openpyxl.styles import PatternFill
from glob import glob
from shutil import copyfile
from cta import update_table
from discharge_extract import extract_specific_tags_df
from discharge_ncs import discharge_ncs
import numpy as np
from collections import defaultdict
from ActiveLearner import ActiveLearner, DISCHARGEFilter
from featureSelection import featureSelection
from openpyxl.utils import get_column_letter

sys.path.append('H:/cloud/cloud_data/Projects/DL/Code/src')
sys.path.append('H:/cloud/cloud_data/Projects/DL/Code/src/ct')
from CTDataStruct import CTPatient
import keyboard
from sklearn.metrics import confusion_matrix
from sklearn.model_selection import train_test_split
from sklearn.metrics import accuracy_score
from sklearn.ensemble import RandomForestClassifier
from numpy.random import shuffle
from openpyxl.styles.differential import DifferentialStyle
from openpyxl import Workbook
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
from openpyxl.formatting import Rule
from computeCTA import computeCTA
from datetime import datetime

from discharge_extract import extract_specific_tags

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
recoClasses = ['FBP', 'IR', 'UNDEFINED']
changeClasses = ['NO_CHANGE', 'SOURCE_CHANGE', 'MASTER_CHANGE', 'MASTER_SOURCE_CHANGE']

def openMaster(folderpath_master, master_process=False):
    print('Open master')
    date = folderpath_master.split('_')[-1]
    if master_process==False:
        filepath_master = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')
    else:
        filepath_master_tmp = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')
        folderpath, filename, file_extension = splitFilePath(filepath_master_tmp)
        filepath_master = os.path.join(folderpath, filename + '_process' + file_extension)
    os.system(filepath_master)

def mergeSameCells(sheet, df_master, merge_column='PatientID'):
    cols = list(df_master.columns)
    column_df = cols.index(merge_column)
    column_excel = column_df + 2
    start_row = 2
    for i in range(0, len(df_master)-1):
        if not df_master.iloc[i, column_df] == df_master.iloc[i+1, column_df]:
            end_row = i + 2
            sheet.merge_cells( start_row=start_row, start_column=column_excel,end_row=end_row, end_column=column_excel)
            start_row = end_row + 1
    end_row = i + 2
    sheet.merge_cells( start_row=start_row, start_column=column_excel,end_row=end_row, end_column=column_excel)

def set_upper_border(sheet, df_master, merge_column='PatientID'):
    cols = list(df_master.columns)
    column_df = cols.index(merge_column)
    column_excel = column_df + 2
    thin = Side(border_style="thin", color="000000")
    for i in range(0, len(df_master)-1):
        if i % 100 == 0:
            print('index:', i, '/', len(df_master))
        if not df_master.iloc[i, column_df] == df_master.iloc[i+1, column_df]:
            end_row = i + 2
            for column in sheet.columns:
                cell = column[end_row]
                cell.border = Border(top=thin)
            start_row = end_row + 1
    end_row = i + 2
    for column in sheet.columns:
        cell = column[end_row]
        cell.border = Border(top=thin)
    # Set border for index
    for i in range(0, len(df_master)-1):
        end_row = i + 2  
        column = sheet['A']
        cell = column[end_row]
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        
def setColor(workbook, sheet, rows, NumColumns, color):
    for r in rows:
        if r % 100 == 0:
            print('index:', r, '/', max(rows))
        for c in range(1,NumColumns):
            cell = sheet.cell(r, c)
            cell.fill = PatternFill(start_color=color, end_color=color, fill_type = 'solid')
            #cell.fill = PatternFill(fgColor=font_color, bgColor=bg_color)
            #formatColor = workbook.add_format({'font_color': font_color, 'bg_color': bg_color})     
            #sheet.conditional_format(first_row, first_col, last_row, last_col,{'type': 'no_blanks','format': formatColor})                  

def setColorFormula(sheet, formula, color, colorrange):
    #color_fill = PatternFill(bgColor=color)
    #dxf = DifferentialStyle(fill=color_fill)
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

def filter_CACS_10StepsGuide(df_data):
    df = df_data.copy()
    df.replace(to_replace=[np.nan], value=0.0, inplace=True)
    print('Apply filter_CACS_10StepsGuide')
    df_CACS = pd.DataFrame(columns=['CACS10StepsGuide'])
    for index, row in df.iterrows():
        if index % 1000 == 0:
            print('index:', index, '/', len(df))
        criteria1 = row['ReconstructionDiameter'] <= 200
        criteria2 = (row['SliceThickness']==3.0) or (row['SliceThickness']==2.5 and row['Site'] in ['P10', 'P13', 'P29'])
        criteria3 = row['Modality'] == 'CT'
        criteria4 = isNaN(row['ContrastBolusAgent'])
        criteria5 = row['Count']>=30 and row['Count']<=90
        result = criteria1 and criteria2 and criteria3 and criteria4 and criteria5
        df_CACS = df_CACS.append({'CACS10StepsGuide': result}, ignore_index=True)
    return df_CACS

def filter_CACS_Extended(df_data):
    print('Apply filter_CACS_Extended')
    df_CACS = pd.DataFrame(columns=['CACSExtended'])
    for index, row in df_data.iterrows():
        if index % 1000 == 0:
            print('index:', index, '/', len(df_data))
        criteria1 = row['ReconstructionDiameter'] <= 300
        criteria2 = (row['SliceThickness']==3.0) or (row['SliceThickness']==2.5 and row['Site'] in ['P10', 'P13', 'P29'])
        criteria3 = row['Modality'] == 'CT'
        criteria4 = isNaN(row['ContrastBolusAgent'])
        criteria5 = row['Count']>=30 and row['Count']<=90
        result = criteria1 and criteria2 and criteria3 and criteria4 and criteria5
        df_CACS = df_CACS.append({'CACSExtended': result}, ignore_index=True)
    return df_CACS

def filter_NCS_Extended(df_data):
    df = pd.DataFrame()
    df_lung, df_body = discharge_ncs(df_data)
    df['NCS_CACSExtended'] = df_lung
    df['NCS_CTAExtended'] = df_body
    return df

def filterReconstruction(df_data):
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
            reco = recoClasses[0]
        elif isir:
            reco = recoClasses[1]
        else:
            reco = recoClasses[2]
        df_reco = df_reco.append({'RECO': reco}, ignore_index=True)
    return df_reco

def filter_CTA_Extended(folderpath_master):
    df_cta = computeCTA(folderpath_master)
    df = pd.DataFrame()
    df['phase'] = df_cta['CTA_phase']
    df['arteries'] = df_cta['CTA_arteries']
    df['source'] = df_cta['CTA_source']
    df['CTAExtended'] = df_cta['CTAExtended']
    df.fillna(value=np.nan, inplace=True)   
    return df    

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

def update_discharge_master(folderpath_discharge_master_sources, folderpath_discharge_master_data):
    
    #folderpath_discharge_master_sources = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_master/discharge_master_sources_20200401'
    folderpath_discharge_master_sources = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_master/discharge_master_sources_20200501'
        
    #filepath_dicom = glob(folderpath_discharge_master_sources + '/*dicom*.xlsx')[0]
    filepath_eCRF = glob(folderpath_discharge_master_sources + '/*_ecrf_*.xlsx')[0]
    #filepath_eCRF01 = glob(folderpath_discharge_master_sources + '/*eCRF01*.xlsx')[0]
    #filepath_eCRF02 = glob(folderpath_discharge_master_sources + '/*eCRF01*.xlsx')[0]
    filepath_tracking = glob(folderpath_discharge_master_sources + '/*tracking*.xlsx')[0]
    filepath_ITT = glob(folderpath_discharge_master_sources + '/*ITT*.xlsx')[0]
    filepath_phase_exclude_stenosis = glob(folderpath_discharge_master_sources + '/*phase_exclude_stenosis*.xlsx')[0]
    filepath_prct = glob(folderpath_discharge_master_sources + '/*prct*.xlsx')[0]
    filepath_stenosis_bigger_20_phases = glob(folderpath_discharge_master_sources + '/*stenosis_bigger_20_phases*.xlsx')[0]
    filepath_tags = glob(folderpath_discharge_master_sources + '/*tags*.xlsx')[0]
    filepath_master_plus_ecrf = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_master/discharge_master_20200401/tmp/discharge_master_template.xlsx'
    filepath_template = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_master/discharge_master_template/discharge_master_template.xlsx'
    filepath_master_old = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_master/discharge_master_20200401/discharge_master_20200401.xlsx'
    filepath_master = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_master/discharge_master_20200501/discharge_master_20200501.xlsx'
    filepath_tags_tmp = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_master/discharge_master_20200501/tmp/discharge_tags_tmp.xlsx'
    #df_data_new = create_discharge_master_merge(folderpath_discharge_master_sources, folderpath_discharge_master_data)
    #df_data = df_data_new.copy()
    
    df_data_old = pd.read_excel(filepath_master_old, 'DATA', index_col=0)
    
    # Merge dicom
    df_dicom = pd.read_excel(filepath_tags, 'linear', index_col=0)
    df_dicom = df_dicom[tags_dicom]
    #df_dicom = df_dicom[0:305]
    df_data = mergeDicom(df_dicom, df_data_old)
    
    # Filter by CACS based on 10-Steps-Guide
    df_CACS = filter_CACS_10StepsGuide(df_data)
    df_CACS.rename(columns={'CACS10StepsGuide':'CACS10StepsGuide'}, inplace=True)
    df_data['CACS10StepsGuide'] = df_CACS
    
    # Filter by CACS extended
    df_CACS = filter_CACS_Extended(df_data)
    df_CACS.rename(columns={'CACS10StepsGuide':'CACSExtended'}, inplace=True)
    df_data['CACSExtended'] = df_CACS
    
    # Filter by CTA based on 10-Steps-Guide
    df_CACS  = filter_CACS_10StepsGuide(df_data)
    df_CACS.rename(columns={'CACS10StepsGuide':'CTA10StepsGuide'}, inplace=True)
    df_data['CTA10StepsGuide'] = df_CACS
    
    # Filter by CTA extended
    df_CACS  = filter_CACS_10StepsGuide(df_data)
    df_CACS.rename(columns={'CACS10StepsGuide':'CTAExtended'}, inplace=True)
    df_data['CTAExtended'] = df_CACS
    
    # Filter by NCS_CACS based on 10-Steps-Guide
    df_CACS = filter_CACS_10StepsGuide(df_data)
    df_CACS.rename(columns={'CACS10StepsGuide':'NCS_CACS10StepsGuide'}, inplace=True)
    df_data['NCS_CACS10StepsGuide'] = df_CACS
    
    # Filter by CACS extended
    df_CACS = filter_CACS_10StepsGuide(df_data)
    df_CACS.rename(columns={'CACS10StepsGuide':'NCS_CACSExtended'}, inplace=True)
    df_data['NCS_CACSExtended'] = df_CACS
    
    # Filter by CTA based on 10-Steps-Guide
    df_CACS  = filter_CACS_10StepsGuide(df_data)
    df_CACS.rename(columns={'CACS10StepsGuide':'NCS_CTA10StepsGuide'}, inplace=True)
    df_data['NCS_CTA10StepsGuide'] = df_CACS
    
    # Filter by CTA extended
    df_CACS  = filter_CACS_10StepsGuide(df_data)
    df_CACS.rename(columns={'CACS10StepsGuide':'NCS_CTAExtended'}, inplace=True)
    df_data['NCS_CTAExtended'] = df_CACS
        
    # Filter by ICA extended
    df_CACS  = filter_CACS_10StepsGuide(df_data)
    df_CACS.rename(columns={'CACS10StepsGuide':'ICA'}, inplace=True)
    df_data['ICA'] = df_CACS

    # Create Class 10StepGuide
    df_CACS  = filter_CACS_10StepsGuide(df_data)
    df_CACS.rename(columns={'CACS10StepsGuide':'CLASS10StepsGuide'}, inplace=True)
    df_data['CLASS10StepsGuide'] = df_CACS

    # Create Class Extended
    df_CACS  = filter_CACS_10StepsGuide(df_data)
    df_CACS.rename(columns={'CACS10StepsGuide':'CLASSExtended'}, inplace=True)
    df_data['CLASSExtended'] = df_CACS
    
    # Create ClassManual
    df_data['ClassManualCorrection'] = df_data_old['ClassManualCorrection']
    
    # Create ClassExtended
    df_data['ClassExtended'] = df_data_old['ClassExtended']
        
    # Merge tracking table
    df_tracking = pd.read_excel(filepath_tracking, 'tracking table')
    df_data = mergeTracking(df_tracking, df_data, df_data_old)
    
    # Create filepath_tags_tmp
    tags_tmp = ['ImageComments', 'NominalPercentageOfCardiacPhase', 'CardiacRRIntervalSpecified', 'SeriesInstanceUID', 'ConvolutionKernel', 'PatientID']
    df_tags = pd.read_excel(filepath_tags)
    df_tags_tmp = df_tags[tags_tmp]
    df_tags_tmp.to_excel(filepath_tags_tmp)
    
    
    # Merge cta
    path_name_list = [
                    (filepath_tags_tmp, 'discharge'),
                    (filepath_master_old, 'master_huhu'),
                    (filepath_ITT,'itt'),
                    (filepath_phase_exclude_stenosis,'phase_exclude_stenosis'),
                    (filepath_stenosis_bigger_20_phases,'stenosis_bigger_20_phases'),
                    (filepath_prct,'prct'),
                    (filepath_eCRF,'ecrf')
                    ]
    DATABASE = 'tests'
    SQL_SCRIPTS = ['ecrf_queries.sql', 'connect_dicom.sql'] 
    MYSQL_ENGINE = 'mysql://root:password@localhost/?charset=utf8'
    EXCEL_OUT = 'master_new.xlsx'
    update_table(path_name_list, DATABASE, SQL_SCRIPTS, MYSQL_ENGINE, EXCEL_OUT)
    
    # Sort table
    df_data = df_data.sort_values(["PatientID", "SeriesInstanceUID"], ascending = (True, True))
    df_data.reset_index(inplace = True, drop=True)

    # Copy template and replabe DATA sheet
    copyfile(filepath_template, filepath_master)
    
    # Write data
    writer = pd.ExcelWriter(filepath_master, engine="openpyxl", mode="a")
    
    # Save workbook
    workbook  = writer.book
    sheet = workbook['DATA']
    workbook.remove(sheet)
    
    #df_data.loc[0,'Count'] = 333
    df_data.to_excel(writer, sheet_name="DATA")

    sheet = workbook['DATA']
    # Write CACS_10StepsGuide
    #sheet = update_CACS_10StepsGuide(df_CACS, sheet)
    # Highlight_columns
    sheet = highlight_columns(df_master, sheet)
    
    writer.save()

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

def createTables(folderpath_discharge, folderpath_master, folderpath_tables):
    
    date = folderpath_master.split('_')[-1]
    # Create folderpath_master
    if not os.path.isdir(folderpath_master):
        os.mkdir(folderpath_master)
    folderpath_sources = os.path.join(folderpath_master, 'discharge_sources_' + date)
   
    # Create folderpath_sources
    if not os.path.isdir(folderpath_sources):
        os.mkdir(folderpath_sources)
    # Copy dicom table
    filepath_dicom_table = os.path.join(folderpath_tables, 'discharge_dicom.xlsx')
    filepath_dicom = os.path.join(folderpath_sources, 'discharge_dicom_' + date +'.xlsx')
    copyfile(filepath_dicom_table, filepath_dicom)
    
    # Copy ITT table
    filepath_ITT_table = os.path.join(folderpath_tables, 'discharge_ITT.xlsx')
    filepath_ITT = os.path.join(folderpath_sources, 'discharge_ITT_' + date + '.xlsx')
    copyfile(filepath_ITT_table, filepath_ITT)
    
    # Copy ecrf table
    filepath_ecrf_table = os.path.join(folderpath_tables, 'discharge_ecrf.xlsx')
    filepath_ecrf = os.path.join(folderpath_sources, 'discharge_ecrf_' + date + '.xlsx')
    copyfile(filepath_ecrf_table, filepath_ecrf)

    # Copy prct table
    filepath_prct_table = os.path.join(folderpath_tables, 'discharge_prct.xlsx')
    filepath_prct = os.path.join(folderpath_sources, 'discharge_prct_' + date + '.xlsx')
    copyfile(filepath_prct_table, filepath_prct)

    # Copy phase_exclude_stenosis table
    filepath_phase_exclude_stenosis_table = os.path.join(folderpath_tables, 'discharge_phase_exclude_stenosis.xlsx')
    filepath_phase_exclude_stenosis = os.path.join(folderpath_sources, 'discharge_phase_exclude_stenosis_' + date + '.xlsx')
    copyfile(filepath_phase_exclude_stenosis_table, filepath_phase_exclude_stenosis)

    # Copy phase_exclude_stenosis table
    filepath_stenosis_bigger_20_phases_table = os.path.join(folderpath_tables, 'discharge_stenosis_bigger_20_phases.xlsx')
    filepath_stenosis_bigger_20_phases = os.path.join(folderpath_sources, 'discharge_stenosis_bigger_20_phases_' + date + '.xlsx')
    copyfile(filepath_stenosis_bigger_20_phases_table, filepath_stenosis_bigger_20_phases)

    # Copy tracking table
    filepath_tracking_table = os.path.join(folderpath_tables, 'discharge_tracking.xlsx')
    filepath_tracking = os.path.join(folderpath_sources, 'discharge_tracking_' + date + '.xlsx')
    copyfile(filepath_tracking_table, filepath_tracking)

def createData(folderpath_master, NumSamples=None):
    
    # Create filepath
    date = folderpath_master.split('_')[-1]
    print('Create data for discharge_master_' + date)
    folderpath_sources = os.path.join(folderpath_master, 'discharge_sources_' + date)
    folderpath_components = os.path.join(folderpath_master, 'discharge_components_' + date)
    if not os.path.isdir(folderpath_components):
        os.mkdir(folderpath_components)
        
    # Create filepaths
    filepath_ecrf = glob(folderpath_sources + '/*_ecrf_*.xlsx')[0]
    filepath_ITT = glob(folderpath_sources + '/*ITT*.xlsx')[0]
    filepath_phase_exclude_stenosis = glob(folderpath_sources + '/*phase_exclude_stenosis*.xlsx')[0]
    filepath_prct = glob(folderpath_sources + '/*prct*.xlsx')[0]
    filepath_stenosis_bigger_20_phases = glob(folderpath_sources + '/*stenosis_bigger_20_phases*.xlsx')[0]
    filepath_dicom = glob(folderpath_sources + '/*dicom*.xlsx')[0]
    filepath_data = os.path.join(folderpath_components, 'discharge_data_' + date + '.xlsx')

    # Extract dicom data
    df_dicom = pd.read_excel(filepath_dicom, index_col=0)
    
    # Reorder datafame
    dicom_cols = ['Site','PatientID','StudyInstanceUID','SeriesInstanceUID','AcquisitionDate','SeriesNumber', 'Count', 'NumberOfFrames', 'SeriesDescription',
     'Modality','Rows', 'InstanceNumber','ProtocolName','ContrastBolusAgent','ImageComments','PixelSpacing','SliceThickness','ConvolutionKernel',
     'ReconstructionDiameter','RequestedProcedureDescription','ContrastBolusStartTime','NominalPercentageOfCardiacPhase','CardiacRRIntervalSpecified',
     'StudyDate']
    
    # NumStudy=[]
    # PatientID=df_dicom['PatientID'].unique()
    # for patient in PatientID:
    #     if not patient == '08-BUD-0423':
    #         df_pat = df_dicom[df_dicom['PatientID']==patient]
    #         df_study = df_pat['StudyInstanceUID'].unique()
    #         NumStudy.append(df_study.shape[0])

    
    df_dicom = df_dicom[dicom_cols]
    df_dicom = df_dicom[(df_dicom['Modality']=='CT') | (df_dicom['Modality']=='OT')]
    df_dicom = df_dicom.reset_index(drop=True)

    cols = df_dicom.columns.tolist()
    cols_new = cols_first + [x for x in cols if x not in cols_first]
    df_dicom = df_dicom[cols_new]
    df_data = df_dicom.copy()
    df_data = df_data.reset_index(drop=True)
    
    if NumSamples is not None:
        df_data = df_data[NumSamples[0]:NumSamples[1]]

    # Extract ecrf data
    df_ecrf = pd.read_excel(filepath_ecrf)
    df_data = mergeEcrf(df_ecrf, df_data)
    
    # Extract ITT 
    df_ITT = pd.read_excel(filepath_ITT, 'Tabelle1')
    df_data = mergeITT(df_ITT, df_data)
    
    # Extract phase_exclude_stenosis 
    df_phase_exclude_stenosis = pd.read_excel(filepath_phase_exclude_stenosis)
    df_data = mergePhase_exclude_stenosis(df_phase_exclude_stenosis, df_data)
    
    # Extract prct
    df_prct = pd.read_excel(filepath_prct)
    df_data = mergePrct(df_prct, df_data)
    
    # Extract stenosis_bigger_20_phases
    df_stenosis_bigger_20_phases = pd.read_excel(filepath_stenosis_bigger_20_phases)
    df_data = mergeStenosis_bigger_20_phase(df_stenosis_bigger_20_phases, df_data)  
    
    # Reoder columns
    cols = df_data.columns.tolist()
    cols_new = cols_first + [x for x in cols if x not in cols_first]
    
    filepath_master_data = os.path.join(folderpath_components, 'discharge_master_data_' + date + '.xlsx')
    df_data.to_excel(filepath_data)
    copyfile(filepath_data, filepath_master_data)

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
    
def createPredictions(folderpath_master, dataname='discharge_data_'):
    date = folderpath_master.split('_')[-1]
    folderpath_components = os.path.join(folderpath_master, 'discharge_components_' + date)
    if not os.path.isdir(folderpath_components):
        os.mkdir(folderpath_components)
    filepath_data = os.path.join(folderpath_components, dataname + date + '.xlsx')
    filepath_pred = os.path.join(folderpath_components, 'discharge_pred_' + date + '.xlsx')
    
    df_data = pd.read_excel(filepath_data)
    
    # Repace Count for multi-slice format
    #idx=df_data['NumberOfFrames']>0
    #df_data['Count'][idx] = df_data['NumberOfFrames'][idx]
    
    df_pred = pd.DataFrame()
    
    # Filter by CACS based on 10-Steps-Guide
    df = filter_CACS_10StepsGuide(df_data)
    df_pred['CACS10StepsGuide'] = df['CACS10StepsGuide']
    
    # Filter by CACS based extended selection
    df = filter_CACS_Extended(df_data)
    df_pred['CACSExtended'] = df['CACSExtended']
    
    # Filter by NCS_CACS and NCS_CTA based on extended criteria
    df = filter_NCS_Extended(df_data)
    df_pred['NCS_CTAExtended'] = df['NCS_CTAExtended']
    df_pred['NCS_CACSExtended'] = df['NCS_CACSExtended']
    
    # Filter by CTA
    df = filter_CTA_Extended(folderpath_master)
    df_pred['CTAExtended'] = df['CTAExtended'].astype('bool')
    df_pred['CTA_phase'] = df['phase']
    df_pred['CTA_arteries'] = df['arteries']
    df_pred['CTA_source'] = df['source']
   
    # Filter by ICA
    df = pd.DataFrame('', index=np.arange(len(df_pred)), columns=['ICA'])
    df_pred['ICA'] = df['ICA']
 
    # Filter by reconstruction
    df = filterReconstruction(df_data)
    df_pred['RECO'] = df['RECO']
    
    # Predict CLASSExtended
    classes = ['CACSExtended', 'CTAExtended', 'NCS_CTAExtended', 'NCS_CACSExtended']
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
        df_pred.loc[i, 'CLASSExtended'] = value

    # Save predictions    
    df_pred.to_excel(filepath_pred)

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

def createRFClassification(folderpath_master):
    print('Create columns for RF Classification')
    date = folderpath_master.split('_')[-1]
    folderpath_components = os.path.join(folderpath_master, 'discharge_components_' + date)
    if not os.path.isdir(folderpath_components):
        os.mkdir(folderpath_components)

    filepath_data = os.path.join(folderpath_components, 'discharge_data_' + date + '.xlsx')
    filepath_rfc = os.path.join(folderpath_components, 'discharge_rfc_' + date + '.xlsx')
    df_data = pd.read_excel(filepath_data)
    # Repace Count for multi-slice format
    #idx=df_data['NumberOfFrames']>0
    #df_data['Count'][idx] = df_data['NumberOfFrames'][idx]
    
    df_rfc0 = pd.DataFrame('UNDEFINED', index=np.arange(len(df_data)), columns=['RFCLabel'])
    df_rfc1 = pd.DataFrame('UNDEFINED', index=np.arange(len(df_data)), columns=['RFCClass'])
    df_rfc2 = pd.DataFrame(0, index=np.arange(len(df_data)), columns=['RFCConfidence'])
    df_rfc = pd.concat([df_rfc0, df_rfc1, df_rfc2], axis=1)
    df_rfc.to_excel(filepath_rfc)

def initRFClassification(folderpath_master, master_process=False):
    
    date = folderpath_master.split('_')[-1]
    folderpath_components = os.path.join(folderpath_master, 'discharge_components_' + date)
    if not os.path.isdir(folderpath_components):
        os.mkdir(folderpath_components)
        
    if master_process:
        filepath_master = os.path.join(folderpath_master, 'discharge_master_' + date + '_process' + '.xlsx')
    else:
        filepath_master = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')
    print('Classifie columns for RF Classification from', filepath_master)
    #filepath_master = os.path.join(folderpath_components, 'discharge_master_' + date + '.xlsx')
    # Read dataframe
    #df_rfc = pd.read_excel(filepath_rfc, index_col=0)
    sheet_name = 'MASTER_' + date
    if os.path.exists(filepath_master):
        df_master = pd.read_excel(filepath_master, sheet_name=sheet_name, index_col=0)
        
        # Create active learner
        learner = ActiveLearner()
        target = featureSelection(filtername='CACSFilter_V02')
        discharge_filter = target['FILTER']
        
        # Extract features
        #learner.extractFeatures(df_master, discharge_filter)
        
        filepath_hist = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_master/discharge_master_01042020/discharge_master_01042020_hist.pkl'
        dfHist = pd.read_pickle(filepath_hist)
        dfData = pd.concat([dfHist.iloc[:,3:], dfHist.iloc[:,1]],axis=1)
        X = np.array(dfData)
        scanClassesRF = defaultdict(lambda:-1,{'CACSExtended': 0, 'CTAExtended': 1, 'NCS_CACSExtended': 2, 'NCS_CTAExtended': 3, 'OTHER': 4})
        #scanClassesRFInv = defaultdict(lambda:'',{'CACSExtended': 1, 'CTAExtended': 2, 'NCS_CACSExtended': 3, 'NCS_CTAExtended': 4})
        scanClassesRFInv = defaultdict(lambda:'UNDEFINED' ,{0: 'CACSExtended', 1: 'CTAExtended', 2: 'NCS_CACSExtended', 3: 'NCS_CTAExtended', 4: 'OTHER'})
        
        #Y = [scanClassesRF[x] for x in list(dfHist['CLASSExtended'])]
        Y = [scanClassesRF[x] for x in list(df_master['CLASSExtended'])]
        Y = np.array(Y)
        X = np.where(X=='', -1, X)
        
        Target = 'RFCLabel'
        df_class = df_master.copy()
        Y = Y[0:len(df_class[Target])]
        X = X[0:len(df_class[Target])]

        learner.df_features = X

        # Update data
        if sum(Y>0)>0:
            #Yarray = np.array(Y)
            #Yarray[Yarray==0] = -1
            
            df_class[Target] = Y
            
            # Predict random forest
            confidence, C, ACC, pred_class, df_features = learner.confidencePredictor(df_class, discharge_filter, Target = Target)
            print('Confusion matrix:', C)
            pred_class = [scanClassesRFInv[x] for x in list(pred_class)]
            df_master['RFCConfidence'] = confidence
            df_master['RFCClass'] = pred_class
            df_master['RFCLabel'] = [scanClassesRFInv[x] for x in list(Y)]
            
            # Write results to master
            writer = pd.ExcelWriter(filepath_master, engine="openpyxl", mode="a")
            workbook  = writer.book
            sheet = workbook[sheet_name]
            workbook.remove(sheet)
            df_master.to_excel(writer, sheet_name=sheet_name)
            writer.save()
        else:
            print('data are not labled')
    else:
        print('Master', filepath_master, 'not found')
        
def classifieRFClassification(folderpath_master, master_process = False):
    
    date = folderpath_master.split('_')[-1]
    folderpath_components = os.path.join(folderpath_master, 'discharge_components_' + date)
    if not os.path.isdir(folderpath_components):
        os.mkdir(folderpath_components)
    #filepath_master = os.path.join(folderpath_master, 'discharge_master_' + date + '_process' + '.xlsx')
    if master_process==False:
        filepath_master = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')
    else:
        filepath_master_tmp = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')
        folderpath, filename, file_extension = splitFilePath(filepath_master_tmp)
        filepath_master = os.path.join(folderpath, filename + '_process' + file_extension)
        
    print('Classifie columns for RF Classification from', filepath_master)
    #filepath_master = os.path.join(folderpath_components, 'discharge_master_' + date + '.xlsx')
    # Read dataframe
    #df_rfc = pd.read_excel(filepath_rfc, index_col=0)
    sheet_name = 'MASTER_' + date
    if os.path.exists(filepath_master):
        df_master = pd.read_excel(filepath_master, sheet_name=sheet_name, index_col=0)
        
        # Create active learner
        learner = ActiveLearner()
        target = featureSelection(filtername='CACSFilter_V02')
        discharge_filter = target['FILTER']

        filepath_hist = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_master/discharge_master_01042020/discharge_master_01042020_hist.pkl'
        dfHist = pd.read_pickle(filepath_hist)
        
        df0 = df_master[['SeriesInstanceUID','RFCLabel']]
        df_merge = df0.merge(dfHist, on=['SeriesInstanceUID', 'SeriesInstanceUID'])
        
        #dfHist = dfHist[idx_ct]
        #dfData = pd.concat([dfHist.iloc[:,3:], dfHist.iloc[:,1]],axis=1)
        #dfData = pd.concat([df_merge.iloc[:,4:104], df_merge.iloc[:,2]],axis=1)
        dfData = pd.concat([df_merge.iloc[:,4:104]],axis=1)
        X = np.array(dfData)
        scanClassesRF = defaultdict(lambda:-1,{'CACSExtended': 0, 'CTAExtended': 1, 'NCS_CACSExtended': 2, 'NCS_CTAExtended': 3, 'OTHER': 4})
        #scanClassesRFInv = defaultdict(lambda:'',{'CACSExtended': 1, 'CTAExtended': 2, 'NCS_CACSExtended': 3, 'NCS_CTAExtended': 4})
        scanClassesRFInv = defaultdict(lambda:'UNDEFINED' ,{0: 'CACSExtended', 1: 'CTAExtended', 2: 'NCS_CACSExtended', 3: 'NCS_CTAExtended', 4: 'OTHER'})


        Y = [scanClassesRF[x] for x in list(df_merge['RFCLabel'])]
        Y = np.array(Y)
        X = np.where(X=='', -1, X)
        
        Target = 'RFCLabel'
        #df_class = df_master.copy()
        #Y = Y[0:len(result[Target])]
        #X = X[0:len(result[Target])]
        
        
        #Xsel=X[Y>-1]
        #Ysel=Y[Y>-1]
    
        learner.df_features = X
        
        #Target = 'RFCLabel'
        
        # Update data
        #Y = df_master[Target].copy()
        #Y = [scanClassesRF[x] for x in list(Y)]
        
        if sum(Y>0)>0:
            #Yarray = np.array(Y)
            #Yarray[Yarray==0] = -1
            #df_class = df_master.copy()
            #df_class[Target] = Y
            df_merge[Target] = Y
            
            # Predict random forest
            confidence, C, ACC, pred_class, df_features = learner.confidencePredictor(df_merge, discharge_filter, Target = Target)
            print('Confusion matrix:', C)
            pred_class = [scanClassesRFInv[x] for x in list(pred_class)]
            df_master['RFCConfidence'] = confidence
            df_master['RFCClass'] = pred_class
            
            # Write results to master
            writer = pd.ExcelWriter(filepath_master, engine="openpyxl", mode="a")
            workbook  = writer.book
            sheet = workbook[sheet_name]
            workbook.remove(sheet)
            df_master.to_excel(writer, sheet_name=sheet_name)
            writer.save()
        else:
            print('data are not labled')
    else:
        print('Master', filepath_master, 'not found')
        
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

    
def createManualSelection(folderpath_master):
    print('Create manual selection')
    date = folderpath_master.split('_')[-1]
    folderpath_components = os.path.join(folderpath_master, 'discharge_components_' + date)
    if not os.path.isdir(folderpath_components):
        os.mkdir(folderpath_components)
    folderpath_sources = os.path.join(folderpath_master, 'discharge_sources_' + date)

    filepath_data = os.path.join(folderpath_components, 'discharge_data_' + date + '.xlsx')
    filepath_manual = os.path.join(folderpath_components, 'discharge_manual_' + date + '.xlsx')
    
    df_data = pd.read_excel(filepath_data, index_col=0)
    df_manual0 = pd.DataFrame('UNDEFINED', index=np.arange(len(df_data)), columns=['ClassManualCorrection'])
    df_manual1 = pd.DataFrame('', index=np.arange(len(df_data)), columns=['Comment'])
    df_manual2 = pd.DataFrame('', index=np.arange(len(df_data)), columns=['Responsible Person'])
    df_manual = pd.concat([df_manual0, df_manual1, df_manual2], axis=1)
    df_manual.to_excel(filepath_manual)

    
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

def createTrackingTable(folderpath_master, create_master_track=True, master_process=False):
    print('Create tracking table')
    
    date = folderpath_master.split('_')[-1]
    folderpath_components = os.path.join(folderpath_master, 'discharge_components_' + date)
    columns = ['ProblemID', 'Site', 'PatientID', 'StudyInstanceUID', 'SeriesInstanceUID', 'Problem Summary',
               'Problem', 'Date of Query', 'Date of the change/sending', 'Results', 'Answer from the site', 'Status', 'Responsible Person']


    df_track = pd.DataFrame(columns=columns)
    #folderpath_sources = os.path.join(folderpath_master, 'discharge_sources_' + date)
    #filepath_track = os.path.join(folderpath_components, 'discharge_track_' + date + '.xlsx')
    #filepath_tracking = glob(folderpath_sources + '/*tracking*.xlsx')[0]
    #filepath_master_track = os.path.join(folderpath_components, 'discharge_master_track_' + date + '.xlsx')
    #filepath_master_data = os.path.join(folderpath_components, 'discharge_master_data_' + date + '.xlsx')
    #filepath_master = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')
    if master_process==False:
        filepath_master = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')
    else:
        filepath_master_tmp = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')
        folderpath, filename, file_extension = splitFilePath(filepath_master_tmp)
        filepath_master = os.path.join(folderpath, filename + '_process' + file_extension)
    
    # Read tracking table
    # = pd.read_excel(filepath_tracking, 'tracking table')
    #df_tracking.replace(to_replace=[np.nan], value='', inplace=True)
    #df_data = pd.read_excel(filepath_master_data, index_col=0)
    #df_master = pd.read_excel(filepath_master, index_col=0)
    
    #df_data = df_data.copy()
    #df_tracking = df_tracking.copy()
    #df_data.replace(to_replace=[np.nan], value='', inplace=True)
    #df_tracking.replace(to_replace=[np.nan], value='', inplace=True)

    # df_track.to_excel(filepath_track)
    # if create_master_track:
    #     copyfile(filepath_track, filepath_master_track)

    # Update master
    writer = pd.ExcelWriter(filepath_master, engine="openpyxl", mode="a")
    # Remove sheet if already exist
    sheet_name = 'TRACKING' + '_' + date
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
    
    
def updateMasterFromTrackingTable(folderpath_master, master_process=False):
    print('Create tracking table')

    date = folderpath_master.split('_')[-1]
    folderpath_components = os.path.join(folderpath_master, 'discharge_components_' + date)

    folderpath_sources = os.path.join(folderpath_master, 'discharge_sources_' + date)
    folderpath_tracking = os.path.join(folderpath_master, 'discharge_tracking')
    filepath_tracking = glob(folderpath_tracking + '/*tracking*.xlsx')[0]
    filepath_master_track = os.path.join(folderpath_components, 'discharge_master_track_' + date + '.xlsx')
    filepath_master_data = os.path.join(folderpath_components, 'discharge_master_data_' + date + '.xlsx')
    #filepath_master = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')
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
    writer = pd.ExcelWriter(filepath_master, engine="openpyxl", mode="a")
    # Remove sheet if already exist
    sheet_name = 'TRACKING' + '_' + date
    workbook  = writer.book
    sheetnames = workbook.sheetnames
    if sheet_name in sheetnames:
        sheet = workbook[sheet_name]
        workbook.remove(sheet)
        
    # Add patient ro master
    df_track.to_excel(writer, sheet_name=sheet_name)
    writer.save()
    
    # Update tracking
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
    writer.save()

def orderMasterData(df_master):

    # Reoder columns
    cols_first=['Site', 'PatientID', 'StudyInstanceUID', 'SeriesInstanceUID', 'CLASSExtended', 'RFCClass', 'RFCLabel', 'RFCConfidence',
                'ClassManualCorrection', 'Comment', 'Responsible Person', 'Count', 'SeriesDescription', 'SeriesNumber', 'AcquisitionDate']
    
    cols = df_master.columns.tolist()
    cols_new = cols_first + [x for x in cols if x not in cols_first]
    df_master = df_master[cols_new]
    df_master = df_master.sort_values(['PatientID', 'StudyInstanceUID', 'SeriesInstanceUID'], ascending = (True, True, True))
    df_master.reset_index(inplace=True, drop=True)
    return df_master
    
def mergeMaster(folderpath_master, folderpath_master_before):
    print('Create master')
    
    folderpath_master_before_list = glob(folderpath_master_before + '/*master*')
    folderpath_master_before_list = sortFolderpath(folderpath_master, folderpath_master_before_list)
    
    # Filer folderpath_master_before_list by process file existance
    folderpath_master_before_list_tmp = folderpath_master_before_list
    folderpath_master_before_list=[]
    for folder in folderpath_master_before_list_tmp:
        if len(glob(folder + '/*process*.xlsx'))>0:
            folderpath_master_before_list.append(folder)

    date = folderpath_master.split('_')[-1]
    folderpath_components = os.path.join(folderpath_master, 'discharge_components_' + date)
    if not os.path.isdir(folderpath_components):
        os.mkdir(folderpath_components)
    folderpath_sources = os.path.join(folderpath_master, 'discharge_sources_' + date)
    filepath_master = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')
    filepath_data = os.path.join(folderpath_components, 'discharge_data_' + date + '.xlsx')
    filepath_master_data = os.path.join(folderpath_components, 'discharge_master_data_' + date + '.xlsx')
    filepath_pred = os.path.join(folderpath_components, 'discharge_pred_' + date + '.xlsx')
    filepath_rfc = os.path.join(folderpath_components, 'discharge_rfc_' + date + '.xlsx')
    filepath_manual = os.path.join(folderpath_components, 'discharge_manual_' + date + '.xlsx')
    filepath_track = os.path.join(folderpath_components, 'discharge_track_' + date + '.xlsx')
    filepath_master_track = os.path.join(folderpath_components, 'discharge_master_track_' + date + '.xlsx')
    filepath_change = os.path.join(folderpath_components, 'discharge_change_' + date + '.xlsx')
    filepath_patient = os.path.join(folderpath_components, 'discharge_patient_' + date + '.xlsx')
    
    # Read tables
    print('Read discharge_master_data')
    df_master_data = pd.read_excel(filepath_master_data, index_col=0)
    print('Read discharge_data')
    df_data = pd.read_excel(filepath_data, index_col=0)
    print('Read discharge_pred')
    df_pred = pd.read_excel(filepath_pred, index_col=0)
    df_pred['CTAExtended'] = df_pred['CTAExtended'].astype('bool')
    #print('Read discharge_patient')
    #df_patient = pd.read_excel(filepath_patient, index_col=0)
    print('Read discharge_rfc')
    df_rfc = pd.read_excel(filepath_rfc, index_col=0)
    print('Read discharge_manual')
    df_manual = pd.read_excel(filepath_manual, index_col=0)
    print('Read discharge_track')
    #df_master_track = pd.read_excel(filepath_master_track, index_col=0)
    if len(folderpath_master_before_list)>1:
        print('Read discharge_change')
        df_change = pd.read_excel(filepath_change, index_col=0)
    print('Create discharge_master')
    #df_master = pd.concat([df_master_data, df_pred, df_rfc, df_manual, df_master_track], axis=1)
    df_master = pd.concat([df_master_data, df_pred, df_rfc, df_manual], axis=1)
    df_master = orderMasterData(df_master)
    
    #df_master.to_excel(filepath_master)

    if len(folderpath_master_before_list)>0:
        filepathMasters = [glob(folder + '/*master*.xlsx')[0] for folder in folderpath_master_before_list]
        filepathMasters = filepathMasters + glob(folderpath_master_before_list[-1] + '/*process*.xlsx')
    else:
        filepathMasters = []

    writer = pd.ExcelWriter(filepath_master, engine="openpyxl", mode="w")
    
    for file in filepathMasters:
        if 'process' in file:
            date_before = file[0:-5].split('_')[-2]
            df = pd.read_excel(file, index_col=0)
            df.to_excel(writer, sheet_name = 'DATA' + '_' + date_before + '_process')
        else:
            date_before = file[0:-5].split('_')[-1]
            df = pd.read_excel(file, index_col=0)
            df.to_excel(writer, sheet_name = 'DATA' + '_' + date_before)

    # #df_master.to_excel(writer, sheet_name = 'DATA' + '_' + date)
    df_data.to_excel(writer, sheet_name = 'DATA' + '_' + date)
    if len(folderpath_master_before_list)>1:
        df_change.to_excel(writer, sheet_name = 'DATA_CHANGES' + '_' + date)
        
    #df_master.set_index(['PatientID','StudyInstanceUID','SeriesInstanceUID'])
    df_master = orderMasterData(df_master)
    df_master.to_excel(writer, sheet_name = 'MASTER' + '_' + date)
    # Add patient data
    #df_patient.to_excel(writer, sheet_name = 'PATIENT_STATUS' + '_' + date)
    writer.save()
    
    #return df_master 

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


def formatMaster(folderpath_master, master_process = False, format='ALL'):
    print('Format master')
    date = folderpath_master.split('_')[-1]
    folderpath_components = os.path.join(folderpath_master, 'discharge_components_' + date)
    if not os.path.isdir(folderpath_components):
        os.mkdir(folderpath_components)
    folderpath_sources = os.path.join(folderpath_master, 'discharge_sources_' + date)
    if master_process==False:
        filepath_master = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')
    else:
        filepath_master_tmp = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')
        folderpath, filename, file_extension = splitFilePath(filepath_master_tmp)
        filepath_master = os.path.join(folderpath, filename + '_process' + file_extension)
    filepath_master_data = os.path.join(folderpath_components, 'discharge_master_data_' + date + '.xlsx')
    filepath_pred = os.path.join(folderpath_components, 'discharge_pred_' + date + '.xlsx')
    filepath_rfc = os.path.join(folderpath_components, 'discharge_rfc_' + date + '.xlsx')
    filepath_manual = os.path.join(folderpath_components, 'discharge_manual_' + date + '.xlsx')
    filepath_master_track = os.path.join(folderpath_components, 'discharge_master_track_' + date + '.xlsx')
    filepath_patient = os.path.join(folderpath_components, 'discharge_patient_' + date + '.xlsx')
    
    # Read tables
    print('Read discharge_data')
    df_data = pd.read_excel(filepath_master_data, index_col=0)
    print('Read discharge_pred')
    df_pred = pd.read_excel(filepath_pred, index_col=0)
    print('Read discharge_rfc')
    df_rfc = pd.read_excel(filepath_rfc, index_col=0)
    print('Read discharge_manual')
    df_manual = pd.read_excel(filepath_manual, index_col=0)
    print('Read discharge_track')
    df_track = pd.read_excel(filepath_master_track, index_col=0)
    print('Read patient_data')
    df_patient = pd.read_excel(filepath_patient, index_col=0)
    print('Create discharge_master')
    #df_master = pd.concat([df_data, df_pred, df_rfc, df_manual, df_track], axis=1)
    #df_master = pd.concat([df_data, df_pred, df_rfc, df_manual], axis=1)
    df_master = pd.read_excel(filepath_master, sheet_name='MASTER_' + date, index_col=0)
    df_master = orderMasterData(df_master)
    #filepath_master = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')

    writer = pd.ExcelWriter(filepath_master, engine="openpyxl", mode="a")
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
            
            dft = pd.read_excel(filepath_master, sheet_name=sheetname, index_col=0)
            
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
    
def createPatient(folderpath_master):
    print('Create patient table.')

    date = folderpath_master.split('_')[-1]
    folderpath_components = os.path.join(folderpath_master, 'discharge_components_' + date)
    filepath_pred = os.path.join(folderpath_components, 'discharge_pred_' + date + '.xlsx')
    filepath_master_data = os.path.join(folderpath_components, 'discharge_master_data_' + date + '.xlsx')
    filepath_patient = os.path.join(folderpath_components, 'discharge_patient_' + date + '.xlsx')
    filepath_master = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')
    
    df_master = pd.read_excel(filepath_master, sheet_name='MASTER_'+ date, index_col=0)
    df_master_data = pd.read_excel(filepath_master_data, index_col=0)
    #df_pred = pd.read_excel(filepath_pred, index_col=0)
    #df_pred['CTAExtended'] = df_pred['CTAExtended'].astype('bool')
    df_patient = pd.DataFrame(columns=['Site', 'PatientID', 'Modality', 'AcquisitionDate',
                                       'CACS_NUM', 'CACS_FBP_NUM', 'CACS_IR_NUM', 
                                       'CTA_NUM', 'CTA_FBP_NUM', 'CTA_IR_NUM', 
                                       'NCS_CACS_NUM', 'NCS_CACS_FBP_NUM', 'NCS_CACS_IR_NUM', 
                                       'NCS_CTA_NUM', 'NCS_CTA_FBP_NUM', 'NCS_CTA_IR_NUM',
                                       'ITT', 'STATUS', 'STATUS_MANUAL_CORRECTION', 'COMMENT'])

    patient_list = df_master_data['PatientID'].unique()
    for patientID in patient_list:
        #data = df_master_data[df_master_data['PatientID']==patientID]
        #if data['Modality'].iloc[0]=='CT':
        df_pat = df_master[df_master['PatientID']==patientID]
        # Extract CACS_NUM information
        CACS_MANUAL_OK = (df_pat['ClassManualCorrection']=='UNDEFINED') | (df_pat['ClassManualCorrection']=='CACS')
        CACS_AUTO_OK = (df_pat['CACSExtended']) & (df_pat['ITT']<2)
        CACS_NUM = (CACS_MANUAL_OK & CACS_AUTO_OK).sum()
        # Extract CACS_FBP_NUM information
        CACS_FBP_MANUAL_OK = (df_pat['ClassManualCorrection']=='UNDEFINED') | (df_pat['ClassManualCorrection']=='CACS')
        CACS_FBP_AUTO_OK = (df_pat['CACSExtended'])  & (df_pat['RECO']=='FBP') & (df_pat['ITT']<2)
        CACS_FBP_NUM = (CACS_MANUAL_OK & CACS_FBP_AUTO_OK).sum()
        # Extract CACS_IR_NUM information
        CACS_IR_MANUAL_OK = (df_pat['ClassManualCorrection']=='UNDEFINED') | (df_pat['ClassManualCorrection']=='CACS')
        CACS_IR_AUTO_OK = (df_pat['CACSExtended']) & (df_pat['RECO']=='IR') & (df_pat['ITT']<2)
        CACS_IR_NUM = (CACS_MANUAL_OK & CACS_IR_AUTO_OK).sum()    
        # Extract CTA_NUM information
        CTA_MANUAL_OK = (df_pat['ClassManualCorrection']=='UNDEFINED') | (df_pat['ClassManualCorrection']=='CTA')
        CTA_AUTO_OK = (df_pat['CTAExtended']) & (df_pat['ITT']<2)
        CTA_NUM = (CTA_MANUAL_OK & CTA_AUTO_OK).sum()   
        # Extract CTA_FBP_NUM information
        CTA_FBP_MANUAL_OK = (df_pat['ClassManualCorrection']=='UNDEFINED') | (df_pat['ClassManualCorrection']=='CTA')
        CTA_FBP_AUTO_OK = (df_pat['CTAExtended'])  & (df_pat['RECO']=='FBP') & (df_pat['ITT']<2)
        CTA_FBP_NUM = (CTA_FBP_MANUAL_OK & CTA_FBP_AUTO_OK).sum()  
        # Extract CTA_IR_NUM information
        CTA_IR_MANUAL_OK = (df_pat['ClassManualCorrection']=='UNDEFINED') | (df_pat['ClassManualCorrection']=='CTA')
        CTA_IR_AUTO_OK = (df_pat['CTAExtended'])  & (df_pat['RECO']=='FBP') & (df_pat['ITT']<2)
        CTA_IR_NUM = (CTA_IR_MANUAL_OK & CTA_IR_AUTO_OK).sum()          
        # Extract NCS_CACS_NUM information
        NCS_CACS_MANUAL_OK = (df_pat['ClassManualCorrection']=='UNDEFINED') | (df_pat['ClassManualCorrection']=='NCS_CACS')
        NCS_CACS_AUTO_OK = (df_pat['NCS_CACSExtended']) & (df_pat['ITT']<2)
        NCS_CACS_NUM = (NCS_CACS_MANUAL_OK & NCS_CACS_AUTO_OK).sum()       
        # Extract NCS_CACS_FBP_NUM information
        NCS_CACS_FBP_MANUAL_OK = (df_pat['ClassManualCorrection']=='UNDEFINED') | (df_pat['ClassManualCorrection']=='NCS_CACS')
        NCS_CACS_FBP_AUTO_OK = (df_pat['NCS_CACSExtended']) & (df_pat['RECO']=='FBP') & (df_pat['ITT']<2)
        NCS_CACS_FBP_NUM = (NCS_CACS_FBP_MANUAL_OK & NCS_CACS_FBP_AUTO_OK).sum()   
        # Extract NCS_CACS_IR_NUM information
        NCS_CACS_IR_MANUAL_OK = (df_pat['ClassManualCorrection']=='UNDEFINED') | (df_pat['ClassManualCorrection']=='NCS_CACS')
        NCS_CACS_IR_AUTO_OK = (df_pat['NCS_CACSExtended']) & (df_pat['RECO']=='IR') & (df_pat['ITT']<2)
        NCS_CACS_IR_NUM = (NCS_CACS_IR_MANUAL_OK & NCS_CACS_IR_AUTO_OK).sum()   
        # Extract NCS_CACS_NUM information
        NCS_CTA_MANUAL_OK = (df_pat['ClassManualCorrection']=='UNDEFINED') | (df_pat['ClassManualCorrection']=='NCS_CTA')
        NCS_CTA_AUTO_OK = (df_pat['NCS_CTAExtended']) & (df_pat['ITT']<2)
        NCS_CTA_NUM = (NCS_CTA_MANUAL_OK & NCS_CTA_AUTO_OK).sum()  
        # Extract NCS_CTA_FBP_NUM information
        NCS_CTA_FBP_MANUAL_OK = (df_pat['ClassManualCorrection']=='UNDEFINED') | (df_pat['ClassManualCorrection']=='NCS_CTA')
        NCS_CTA_FBP_AUTO_OK = (df_pat['NCS_CTAExtended']) & (df_pat['RECO']=='FBP') & (df_pat['ITT']<2)
        NCS_CTA_FBP_NUM = (NCS_CTA_FBP_MANUAL_OK & NCS_CTA_FBP_AUTO_OK).sum() 
        # Extract NCS_CTA_IR_NUM information
        NCS_CTA_IR_MANUAL_OK = (df_pat['ClassManualCorrection']=='UNDEFINED') | (df_pat['ClassManualCorrection']=='NCS_CTA')
        NCS_CTA_IR_AUTO_OK = (df_pat['NCS_CTAExtended']) & (df_pat['RECO']=='IR') & (df_pat['ITT']<2)
        NCS_CTA_IR_NUM = (NCS_CTA_IR_MANUAL_OK & NCS_CTA_IR_AUTO_OK).sum() 
        

        SITE = df_pat['Site'].iloc[0]
        ITT = df_pat['ITT'].iloc[0]
        modality = df_pat['Modality'].iloc[0]
        #AcquisitionDate = df_pat['AcquisitionDate'].iloc[0]
        
        # Check patient scenario
        CACS_OK = CACS_NUM>=2
        CTA_OK = CTA_NUM>=2
        NCS_CACS_OK = NCS_CACS_NUM>=2
        NCS_CTA_OK = NCS_CTA_NUM>=2
        
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
        df_patient = df_patient.append({'Site': SITE, 'PatientID': patientID, 'Modality': modality, 'CACS_NUM': CACS_NUM, 'CACS_FBP_NUM': CACS_FBP_NUM, 'CACS_IR_NUM': CACS_IR_NUM,
                           'CTA_NUM': CTA_NUM, 'CTA_FBP_NUM': CTA_FBP_NUM, 'CTA_IR_NUM': CTA_IR_NUM,
                           'NCS_CACS_NUM': NCS_CACS_NUM, 'NCS_CACS_FBP_NUM': NCS_CACS_FBP_NUM, 'NCS_CACS_IR_NUM': NCS_CACS_IR_NUM,
                           'NCS_CTA_NUM': NCS_CTA_NUM, 'NCS_CTA_FBP_NUM': NCS_CTA_FBP_NUM, 'NCS_CTA_IR_NUM': NCS_CTA_IR_NUM,
                           'ITT': ITT, 'STATUS': status, 'STATUS_MANUAL_CORRECTION': STATUS_MANUAL_CORRECTION, 'COMMENT': COMMENT}, ignore_index=True)
    
    df_patient.to_excel(filepath_patient)
    
    # Remove sheet if already exist
    sheet_name = 'PATIENT_STATUS' + '_' + date
    workbook  = writer.book
    sheetnames = workbook.sheetnames
    if sheet_name in sheetnames:
        sheet = workbook[sheet_name]
        workbook.remove(sheet)
    
    # Add patient ro master
    writer = pd.ExcelWriter(filepath_master, engine="openpyxl", mode="a")
    df_patient.to_excel(writer, sheet_name=sheet_name)
    writer.save()

# def createStudy(folderpath_master, master_process=False, conf=True):
#     print('Create StudyInstanceID table.')
    
#     CACS_NUM_MIN = 2
#     CACS_IR_NUM_MIN = 1
#     CACS_FBP_NUM_MIN = 1
    
#     CTA_NUM_MIN = 2
#     CTA_IR_NUM_MIN = 1
#     CTA_FBP_NUM_MIN = 1
    
#     NCS_CACS_NUM_MIN = 2
#     NCS_CACS_IR_NUM_MIN = 1
#     NCS_CACS_FBP_NUM_MIN = 1
    
#     NCS_CTA_NUM_MIN = 2
#     NCS_CTA_IR_NUM_MIN = 1
#     NCS_CTA_FBP_NUM_MIN = 1

#     date = folderpath_master.split('_')[-1]
#     folderpath_components = os.path.join(folderpath_master, 'discharge_components_' + date)
#     filepath_pred = os.path.join(folderpath_components, 'discharge_pred_' + date + '.xlsx')
#     filepath_master_data = os.path.join(folderpath_components, 'discharge_master_data_' + date + '.xlsx')
#     filepath_patient = os.path.join(folderpath_components, 'discharge_patient_' + date + '.xlsx')
#     filepath_patient_conf = os.path.join(folderpath_components, 'discharge_patient_conf_' + date + '.xlsx')
#     #filepath_master = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')
#     if master_process==False:
#         filepath_master = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')
#     else:
#         filepath_master_tmp = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')
#         folderpath, filename, file_extension = splitFilePath(filepath_master_tmp)
#         filepath_master = os.path.join(folderpath, filename + '_process' + file_extension)
#     df_master = pd.read_excel(filepath_master, sheet_name='MASTER_'+ date, index_col=0)
#     df_master.sort_index(inplace=True)
#     #df_master_data = pd.read_excel(filepath_master_data, index_col=0)
#     df_studyInstanceUID = pd.DataFrame(columns=['Site', 'PatientID', 'StudyInstanceUID', 'Modality', 'AcquisitionDate',
#                                        'CACS_NUM', 'CACS_FBP_NUM', 'CACS_IR_NUM', 
#                                        'CTA_NUM', 'CTA_FBP_NUM', 'CTA_IR_NUM', 
#                                        'NCS_CACS_NUM', 'NCS_CACS_FBP_NUM', 'NCS_CACS_IR_NUM', 
#                                        'NCS_CTA_NUM', 'NCS_CTA_FBP_NUM', 'NCS_CTA_IR_NUM',
#                                        'ITT', 'STATUS', 'STATUS_MANUAL_CORRECTION', 'COMMENT'])
#     df_studyInstanceUID_conf = pd.DataFrame(columns=['Site', 'PatientID', 'StudyInstanceUID', 'Modality', 'AcquisitionDate',
#                                        'CACS_CONF', 'CACS_FBP_CONF', 'CACS_IR_CONF', 
#                                        'CTA_CONF', 'CTA_FBP_CONF', 'CTA_IR_CONF', 
#                                        'NCS_CACS_CONF', 'NCS_CACS_FBP_CONF', 'NCS_CACS_IR_CONF', 
#                                        'NCS_CTA_CONF', 'NCS_CTA_FBP_CONF', 'NCS_CTA_IR_CONF',
#                                        'ITT', 'STATUS', 'STATUS_MANUAL_CORRECTION', 'COMMENT'])

#     # Filter study list
#     func = lambda x: datetime.strptime(x, '%Y%m%d')
#     patients = df_master['PatientID'].unique()
#     firstdateList = []
#     study_list = []
    
#     patients=['06-GOE-0020']
    
    
#     for patient in patients:
#         df_patient = df_master[(df_master['PatientID']==patient) &  (df_master['Modality']=='CT')]
#         if len(df_patient)>0:
#             firstdate = df_patient['1. Date of CT scan'].iloc[0]
            
#             studydate = df_patient['StudyDate']
#             # Convert string to date and replace in df_patient
#             studydate = studydate.apply(lambda x: datetime.strptime(str(x), '%Y%m%d'))
#             df_patient['StudyDate'] = studydate
            
#             if not pd.isnull(firstdate):
#                 df_study = df_patient[studydate == firstdate]
#                 if len(df_study)>0:
#                     df_study_id = df_study['StudyInstanceUID'].iloc[0]
#                 else:
#                     print('PROBLEM: Patient ' + patient + ' "1. Date of CT scan" not consistent with "StudyDate"')
#                     df_study = df_patient
#                     df_study = df_study.sort_values(by='StudyDate')
#                     df_study_id = df_study['StudyInstanceUID'].iloc[0]
#             else:
#                 print('Patient: ' + patient + ' does not have a 1. Date of CT scan')
#                 df_study = df_patient
#                 df_study = df_study.sort_values(by='StudyDate')
#                 df_study_id = df_study['StudyInstanceUID'].iloc[0]
#             study_list.append(df_study_id)
            
#     def getConf(conf, NumSeries=1):
#         conf = conf.sort_values(ascending=False)
#         if len(conf)==0:
#             return 0
#         if len(conf) <= NumSeries:
#             return conf.min()
#         else:
#             return conf[0:NumSeries].min()
       

#     for studyID in study_list:
#         # if studyID == '1.2.840.113619.6.95.31.0.3.4.1.1018.13.10856850':
#         #     sys.exit()
#         #data = df_master_data[df_master_data['PatientID']==patientID]
#         #if data['Modality'].iloc[0]=='CT':
#         df_study = df_master[df_master['StudyInstanceUID']==studyID]
#         # Extract CACS_NUM information
#         CACS_MANUAL_OK = (df_study['ClassManualCorrection']=='UNDEFINED') | (df_study['ClassManualCorrection']=='CACS')
#         #CACS_AUTO_OK = (df_study['CACSExtended']) & (df_study['ITT']<2)
#         CACS_AUTO_OK = (df_study['RFCClass']=='CACSExtended') & (df_study['ITT']<2)
#         CACS_NUM = (CACS_MANUAL_OK & CACS_AUTO_OK).sum()
#         CACS_CONF = getConf(df_study[CACS_MANUAL_OK & CACS_AUTO_OK]['RFCConfidence'], NumSeries=CACS_NUM_MIN)
        
#         # Extract CACS_FBP_NUM information
#         CACS_FBP_MANUAL_OK = (df_study['ClassManualCorrection']=='UNDEFINED') | (df_study['ClassManualCorrection']=='CACS')
#         #CACS_FBP_AUTO_OK = (df_study['CACSExtended'])  & (df_study['RECO']=='FBP') & (df_study['ITT']<2)
#         CACS_FBP_AUTO_OK = (df_study['RFCClass']=='CACSExtended')  & (df_study['RECO']=='FBP') & (df_study['ITT']<2)
#         CACS_FBP_NUM = (CACS_FBP_MANUAL_OK & CACS_FBP_AUTO_OK).sum()
#         CACS_FBP_CONF = getConf(df_study[CACS_MANUAL_OK & CACS_FBP_AUTO_OK]['RFCConfidence'], NumSeries=CACS_FBP_NUM_MIN)
#         # Extract CACS_IR_NUM information
#         CACS_IR_MANUAL_OK = (df_study['ClassManualCorrection']=='UNDEFINED') | (df_study['ClassManualCorrection']=='CACS')
#         #CACS_IR_AUTO_OK = (df_study['CACSExtended']) & (df_study['RECO']=='IR') & (df_study['ITT']<2)
#         CACS_IR_AUTO_OK = (df_study['RFCClass']=='CACSExtended') & (df_study['RECO']=='IR') & (df_study['ITT']<2)
#         CACS_IR_NUM = (CACS_IR_MANUAL_OK & CACS_IR_AUTO_OK).sum()
#         CACS_IR_CONF = getConf(df_study[CACS_MANUAL_OK & CACS_IR_AUTO_OK]['RFCConfidence'], NumSeries=CACS_IR_NUM_MIN)
#         # Extract CTA_NUM information
#         CTA_MANUAL_OK = (df_study['ClassManualCorrection']=='UNDEFINED') | (df_study['ClassManualCorrection']=='CTA')
#         #CTA_AUTO_OK = (df_study['CTAExtended']) & (df_study['ITT']<2)
#         CTA_AUTO_OK = (df_study['RFCClass']=='CTAExtended') & (df_study['ITT']<2)
#         CTA_NUM = (CTA_MANUAL_OK & CTA_AUTO_OK).sum()  
#         CTA_CONF = getConf(df_study[CTA_MANUAL_OK & CTA_AUTO_OK]['RFCConfidence'], NumSeries=CTA_NUM_MIN)
#         # Extract CTA_FBP_NUM information
#         CTA_FBP_MANUAL_OK = (df_study['ClassManualCorrection']=='UNDEFINED') | (df_study['ClassManualCorrection']=='CTA')
#         #CTA_FBP_AUTO_OK = (df_study['CTAExtended'])  & (df_study['RECO']=='FBP') & (df_study['ITT']<2)
#         CTA_FBP_AUTO_OK = (df_study['RFCClass']=='CTAExtended')  & (df_study['RECO']=='FBP') & (df_study['ITT']<2)
#         CTA_FBP_NUM = (CTA_FBP_MANUAL_OK & CTA_FBP_AUTO_OK).sum()  
#         CTA_FBP_CONF = getConf(df_study[CTA_FBP_MANUAL_OK & CTA_FBP_AUTO_OK]['RFCConfidence'], NumSeries=CTA_FBP_NUM_MIN)
#         # Extract CTA_IR_NUM information
#         CTA_IR_MANUAL_OK = (df_study['ClassManualCorrection']=='UNDEFINED') | (df_study['ClassManualCorrection']=='CTA')
#         #CTA_IR_AUTO_OK = (df_study['CTAExtended'])  & (df_study['RECO']=='FBP') & (df_study['ITT']<2)
#         CTA_IR_AUTO_OK = (df_study['RFCClass']=='CTAExtended')  & (df_study['RECO']=='FBP') & (df_study['ITT']<2)
#         CTA_IR_NUM = (CTA_IR_MANUAL_OK & CTA_IR_AUTO_OK).sum()          
#         CTA_IR_CONF = getConf(df_study[CTA_IR_MANUAL_OK & CTA_IR_AUTO_OK]['RFCConfidence'], NumSeries=CTA_IR_NUM_MIN)
#         # Extract NCS_CACS_NUM information
#         NCS_CACS_MANUAL_OK = (df_study['ClassManualCorrection']=='UNDEFINED') | (df_study['ClassManualCorrection']=='NCS_CACS')
#         #NCS_CACS_AUTO_OK = (df_study['NCS_CACSExtended']) & (df_study['ITT']<2)
#         NCS_CACS_AUTO_OK = (df_study['RFCClass']=='NCS_CACSExtended') & (df_study['ITT']<2)
#         NCS_CACS_NUM = (NCS_CACS_MANUAL_OK & NCS_CACS_AUTO_OK).sum()    
#         NCS_CACS_CONF = getConf(df_study[NCS_CACS_MANUAL_OK & NCS_CACS_AUTO_OK]['RFCConfidence'], NumSeries=NCS_CACS_NUM_MIN)
#         # Extract NCS_CACS_FBP_NUM information
#         NCS_CACS_FBP_MANUAL_OK = (df_study['ClassManualCorrection']=='UNDEFINED') | (df_study['ClassManualCorrection']=='NCS_CACS')
#         #NCS_CACS_FBP_AUTO_OK = (df_study['NCS_CACSExtended']) & (df_study['RECO']=='FBP') & (df_study['ITT']<2)
#         NCS_CACS_FBP_AUTO_OK = (df_study['RFCClass']=='NCS_CACSExtended') & (df_study['RECO']=='FBP') & (df_study['ITT']<2)
#         NCS_CACS_FBP_NUM = (NCS_CACS_FBP_MANUAL_OK & NCS_CACS_FBP_AUTO_OK).sum()  
#         NCS_CACS_FBP_CONF = getConf(df_study[NCS_CACS_FBP_MANUAL_OK & NCS_CACS_FBP_AUTO_OK]['RFCConfidence'], NumSeries=NCS_CACS_FBP_NUM_MIN)
#         # Extract NCS_CACS_IR_NUM information
#         NCS_CACS_IR_MANUAL_OK = (df_study['ClassManualCorrection']=='UNDEFINED') | (df_study['ClassManualCorrection']=='NCS_CACS')
#         #NCS_CACS_IR_AUTO_OK = (df_study['NCS_CACSExtended']) & (df_study['RECO']=='IR') & (df_study['ITT']<2)
#         NCS_CACS_IR_AUTO_OK = (df_study['RFCClass']=='NCS_CACSExtended') & (df_study['RECO']=='IR') & (df_study['ITT']<2)
#         NCS_CACS_IR_NUM = (NCS_CACS_IR_MANUAL_OK & NCS_CACS_IR_AUTO_OK).sum()   
#         NCS_CACS_IR_CONF = getConf(df_study[NCS_CACS_IR_MANUAL_OK & NCS_CACS_IR_AUTO_OK]['RFCConfidence'], NumSeries=NCS_CACS_IR_NUM_MIN)
#         # Extract NCS_CACS_NUM information
#         NCS_CTA_MANUAL_OK = (df_study['ClassManualCorrection']=='UNDEFINED') | (df_study['ClassManualCorrection']=='NCS_CTA')
#         #NCS_CTA_AUTO_OK = (df_study['NCS_CTAExtended']) & (df_study['ITT']<2)
#         NCS_CTA_AUTO_OK = (df_study['RFCClass']=='NCS_CTAExtended') & (df_study['ITT']<2)
#         NCS_CTA_NUM = (NCS_CTA_MANUAL_OK & NCS_CTA_AUTO_OK).sum() 
#         NCS_CTA_CONF = getConf(df_study[NCS_CTA_MANUAL_OK & NCS_CTA_AUTO_OK]['RFCConfidence'], NumSeries=NCS_CTA_NUM_MIN)
#         # Extract NCS_CTA_FBP_NUM information
#         NCS_CTA_FBP_MANUAL_OK = (df_study['ClassManualCorrection']=='UNDEFINED') | (df_study['ClassManualCorrection']=='NCS_CTA')
#         #NCS_CTA_FBP_AUTO_OK = (df_study['NCS_CTAExtended']) & (df_study['RECO']=='FBP') & (df_study['ITT']<2)
#         NCS_CTA_FBP_AUTO_OK = (df_study['RFCClass']=='NCS_CTAExtended') & (df_study['RECO']=='FBP') & (df_study['ITT']<2)
#         NCS_CTA_FBP_NUM = (NCS_CTA_FBP_MANUAL_OK & NCS_CTA_FBP_AUTO_OK).sum() 
#         NCS_CTA_FBP_CONF = getConf(df_study[NCS_CTA_FBP_MANUAL_OK & NCS_CTA_FBP_AUTO_OK]['RFCConfidence'], NumSeries=NCS_CTA_FBP_NUM_MIN)
#         # Extract NCS_CTA_IR_NUM information
#         NCS_CTA_IR_MANUAL_OK = (df_study['ClassManualCorrection']=='UNDEFINED') | (df_study['ClassManualCorrection']=='NCS_CTA')
#         #NCS_CTA_IR_AUTO_OK = (df_study['NCS_CTAExtended']) & (df_study['RECO']=='IR') & (df_study['ITT']<2)
#         NCS_CTA_IR_AUTO_OK = (df_study['RFCClass']=='NCS_CTAExtended') & (df_study['RECO']=='IR') & (df_study['ITT']<2)
#         NCS_CTA_IR_NUM = (NCS_CTA_IR_MANUAL_OK & NCS_CTA_IR_AUTO_OK).sum() 
#         NCS_CTA_IR_CONF = getConf(df_study[NCS_CTA_IR_MANUAL_OK & NCS_CTA_IR_AUTO_OK]['RFCConfidence'], NumSeries=NCS_CTA_IR_NUM_MIN)
        
#         PATIENTID = df_study['PatientID'].iloc[0]
#         SITE = df_study['Site'].iloc[0]
#         ITT = df_study['ITT'].iloc[0]
#         modality = df_study['Modality'].iloc[0]
#         DATE = df_study['AcquisitionDate'].iloc[0]
#         #AcquisitionDate = df_study['AcquisitionDate'].iloc[0]
        
#         # Check patient scenario
#         CACS_OK = CACS_NUM>=CACS_NUM_MIN
#         CTA_OK = CTA_NUM>=CTA_NUM_MIN
#         NCS_CACS_OK = NCS_CACS_NUM>=NCS_CACS_NUM_MIN
#         NCS_CTA_OK = NCS_CTA_NUM>=NCS_CTA_NUM_MIN
        
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
#         df_studyInstanceUID = df_studyInstanceUID.append({'Site': SITE, 'PatientID': PATIENTID, 'StudyInstanceUID': studyID, 'Modality': modality, 'AcquisitionDate': DATE, 'CACS_NUM': CACS_NUM, 'CACS_FBP_NUM': CACS_FBP_NUM, 'CACS_IR_NUM': CACS_IR_NUM,
#                            'CTA_NUM': CTA_NUM, 'CTA_FBP_NUM': CTA_FBP_NUM, 'CTA_IR_NUM': CTA_IR_NUM,
#                            'NCS_CACS_NUM': NCS_CACS_NUM, 'NCS_CACS_FBP_NUM': NCS_CACS_FBP_NUM, 'NCS_CACS_IR_NUM': NCS_CACS_IR_NUM,
#                            'NCS_CTA_NUM': NCS_CTA_NUM, 'NCS_CTA_FBP_NUM': NCS_CTA_FBP_NUM, 'NCS_CTA_IR_NUM': NCS_CTA_IR_NUM,
#                            'ITT': ITT, 'STATUS': status, 'STATUS_MANUAL_CORRECTION': STATUS_MANUAL_CORRECTION, 'COMMENT': COMMENT}, ignore_index=True)
#         df_studyInstanceUID_conf = df_studyInstanceUID_conf.append({'Site': SITE, 'PatientID': PATIENTID, 'StudyInstanceUID': studyID, 'Modality': modality, 'AcquisitionDate': DATE, 'CACS_CONF': CACS_CONF, 'CACS_FBP_CONF': CACS_FBP_CONF, 'CACS_IR_CONF': CACS_IR_CONF,
#                            'CTA_CONF': CTA_CONF, 'CTA_FBP_CONF': CTA_FBP_CONF, 'CTA_IR_CONF': CTA_IR_CONF,
#                            'NCS_CACS_CONF': NCS_CACS_CONF, 'NCS_CACS_FBP_CONF': NCS_CACS_FBP_CONF, 'NCS_CACS_IR_CONF': NCS_CACS_IR_CONF,
#                            'NCS_CTA_CONF': NCS_CTA_CONF, 'NCS_CTA_FBP_CONF': NCS_CTA_FBP_CONF, 'NCS_CTA_IR_CONF': NCS_CTA_IR_CONF,
#                            'ITT': ITT, 'STATUS': status, 'STATUS_MANUAL_CORRECTION': STATUS_MANUAL_CORRECTION, 'COMMENT': COMMENT}, ignore_index=True)
        
#     df_studyInstanceUID.to_excel(filepath_patient)
#     df_studyInstanceUID_conf.to_excel(filepath_patient_conf)

#     writer = pd.ExcelWriter(filepath_master, engine="openpyxl", mode="a")
#     # Remove sheet if already exist
#     sheet_name = 'PATIENT_STATUS_' + date
#     workbook  = writer.book
#     sheetnames = workbook.sheetnames
#     if sheet_name in sheetnames:
#         sheet = workbook[sheet_name]
#         workbook.remove(sheet)
#     df_studyInstanceUID.to_excel(writer, sheet_name=sheet_name)
        
#     if conf:
#         sheet_name = 'PATIENT_STATUS_CONF_' + date
#         if sheet_name in sheetnames:
#             sheet = workbook[sheet_name]
#             workbook.remove(sheet)
#         df_studyInstanceUID_conf.to_excel(writer, sheet_name=sheet_name)
        
#     # Add patient ro master
    
#     writer.save()

def createStudy(folderpath_master, master_process=False, conf=True):
    print('Create StudyInstanceID table.')
    
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

    date = folderpath_master.split('_')[-1]
    folderpath_components = os.path.join(folderpath_master, 'discharge_components_' + date)
    filepath_pred = os.path.join(folderpath_components, 'discharge_pred_' + date + '.xlsx')
    filepath_master_data = os.path.join(folderpath_components, 'discharge_master_data_' + date + '.xlsx')
    filepath_patient = os.path.join(folderpath_components, 'discharge_patient_' + date + '.xlsx')
    filepath_patient_conf = os.path.join(folderpath_components, 'discharge_patient_conf_' + date + '.xlsx')
    #filepath_master = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')
    if master_process==False:
        filepath_master = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')
    else:
        filepath_master_tmp = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')
        folderpath, filename, file_extension = splitFilePath(filepath_master_tmp)
        filepath_master = os.path.join(folderpath, filename + '_process' + file_extension)
    df_master = pd.read_excel(filepath_master, sheet_name='MASTER_'+ date, index_col=0)
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
            studydate = studydate.apply(lambda x: datetime.strptime(str(x), '%Y%m%d'))
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
    df_PatientID_conf.to_excel(filepath_patient_conf)

    writer = pd.ExcelWriter(filepath_master, engine="openpyxl", mode="a")
    # Remove sheet if already exist
    sheet_name = 'PATIENT_STATUS_' + date
    workbook  = writer.book
    sheetnames = workbook.sheetnames
    if sheet_name in sheetnames:
        sheet = workbook[sheet_name]
        workbook.remove(sheet)
    df_PatientID.to_excel(writer, sheet_name=sheet_name)
        
    if conf:
        sheet_name = 'PATIENT_STATUS_CONF_' + date
        if sheet_name in sheetnames:
            sheet = workbook[sheet_name]
            workbook.remove(sheet)
        df_PatientID_conf.to_excel(writer, sheet_name=sheet_name)
        
    # Add patient ro master
    
    writer.save()
    
def extractHist():

    filepath = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_master/discharge_master_01042020/discharge_master_01042020.xlsx'
    filepath_hist = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_master/discharge_master_01042020/discharge_master_01042020_hist.pkl'
    folderpath_images = 'G:/discharge'
    df = pd.read_excel(filepath, sheet_name ='MASTER_01042020', index_col=0)
    df.sort_index(inplace=True)
    bins = 100   
    columns =['SeriesInstanceUID', 'Count', 'CLASSExtended'] + [str(x) for x in range(0,bins)]
    if os.path.exists(filepath_hist):
        dfHist = pd.read_pickle(filepath_hist)
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
                dfHist.loc[index,'CLASSExtended'] = row['CLASSExtended']
                dfHist.iloc[index,3:] = np.ones((1, bins))*-1
            else:
                try:
                    patient=CTPatient(StudyInstanceUID, PatientID)
                    #series = patient.loadSeries(folderpath_images, SeriesInstanceUID, SOPInstanceUID)
                    series = patient.loadSeries(folderpath_images, SeriesInstanceUID, None)
                    image = series.image.image()
                    hist = np.histogram(image, bins=bins, range=(-2500, 3000))[0]
                    dfHist.loc[index,'SeriesInstanceUID'] = SeriesInstanceUID
                    dfHist.loc[index,'Count'] = image.shape[0]
                    dfHist.loc[index,'CLASSExtended'] = row['CLASSExtended']
                    dfHist.iloc[index,3:] = hist
                except:
                    print('Error index', index)
                    dfHist.loc[index,'SeriesInstanceUID'] = SeriesInstanceUID
                    dfHist.loc[index,'Count'] = -1
                    dfHist.loc[index,'CLASSExtended'] = row['CLASSExtended']
                    dfHist.iloc[index,3:] = np.ones((1, bins))*-1
        if index % 10 == 0:
            dfHist.to_pickle(filepath_hist)
    dfHist.to_pickle(filepath_hist)

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

def extractDICOMTags(folderpath_tables, folderpath_discharge):
    root = folderpath_discharge
    fout = os.path.join(folderpath_tables, 'discharge_dicom.xlsx')
    extract_specific_tags(root, fout, NumSamples=None)
    
def createInitialMaster():
    # Create initial master
    folderpath_tables = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_master/discharge_tables'
    folderpath_master_before = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_master'
    folderpath_discharge = 'G:/discharge'
    folderpath_master = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_master/discharge_master_01042020'
    #NumSamples=(0, 2000)
    NumSamples=None
    
    # Extract dicom tags
    ##extractDICOMTags(folderpath_tables, folderpath_discharge)
    # Create tables
    createTables(folderpath_discharge, folderpath_master, folderpath_tables)
    # Create data
    createData(folderpath_master, NumSamples=NumSamples)
    # Create random forest classification columns
    createRFClassification(folderpath_master)
    # Create manual selection
    createManualSelection(folderpath_master)
    # Create prediction 
    createPredictions(folderpath_master)
    # Merge master 
    mergeMaster(folderpath_master, folderpath_master_before)
    # Create tracking table 
    createTrackingTable(folderpath_master)
    updateMasterFromTrackingTable(folderpath_master)
    # Init RF classifier
    initRFClassification(folderpath_master)
    classifieRFClassification(folderpath_master)
    # Merge study sheet 
    createStudy(folderpath_master, conf=True)
    # Format master
    formatMaster(folderpath_master)

    # # Create tables
    # createTables(folderpath_discharge, folderpath_master, folderpath_tables)
    # # Create data
    # createData(folderpath_master, NumSamples=NumSamples)
    # # Create random forest classification columns
    # createRFClassification(folderpath_master)
    # # Create manual selection
    # createManualSelection(folderpath_master)
    # # Create prediction 
    # createPredictions(folderpath_master)
    # # Create tracking data 
    # #createTracking(folderpath_master)
    # # Merge master 
    # mergeMaster(folderpath_master, folderpath_master_before)
    
    # createStudy(folderpath_master, conf=True)
    # formatMaster(folderpath_master)
    
    # createTrackingTable(folderpath_master)
    # updateMasterFromTrackingTable(folderpath_master)
    # #updateTrackingTableFromMaster(folderpath_master)
    
    
    # # Format master
    # formatMaster(folderpath_master)
    # # Create process version
    # createMasterProcess(folderpath_master)
    # formatMaster(folderpath_master, master_process=True)
    # # Init RF classifier
    # initRFClassification(folderpath_master)
    #classifieRFClassification(folderpath_master)
    
    # createMasterProcess(folderpath_master)
    # classifieRFClassification(folderpath_master, master_process=True)
    # formatMaster(folderpath_master, master_process=True)
    
    # createStudy(folderpath_master, master_process=True)
    # createTrackingTable(folderpath_master, master_process=True)
    # formatMaster(folderpath_master, master_process=True)
    
    # # Clssifie with RF
    # classifieRFClassification(folderpath_master)
    # formatMaster(folderpath_master, master_process=True)

def updateMaster():
    # Define filepath
    folderpath_tables = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_master/discharge_tables'
    folderpath_master_before = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_master'
    folderpath_discharge = 'G:/discharge'
    folderpath_master = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_master/discharge_master_01052020'
    #NumSamples=(26228, 26747)
    NumSamples=(0, 2000)
    
    
    # Create tables
    createTables(folderpath_discharge, folderpath_master, folderpath_tables)
    # Create data
    createData(folderpath_master, NumSamples=NumSamples)
    # Update data
    updateData(folderpath_master, folderpath_master_before)
    # Update random forest classification
    updateRFClassification(folderpath_master, folderpath_master_before)
    # Update manual selection
    updateManualSelection(folderpath_master, folderpath_master_before)
    # Update prediction
    updatePredictions(folderpath_master)
    # Update tracking
    updateTracking(folderpath_master, folderpath_master_before)
    # Merge new master
    mergeMaster(folderpath_master, folderpath_master_before)
    # Update patient table
    updatePatient(folderpath_master, folderpath_master_before)
    # Format master 
    formatMaster(folderpath_master)
    # Create master process
    createMasterProcess(folderpath_master)
    # Format master process
    formatMaster(folderpath_master, master_process=True)
    # Update rando forest classification
    createRFClassification(folderpath_master, mode='classifie')
                
def test():
    # Create initial master
    createInitialMaster()
        
    # Update master
    updateMaster()
    
    # Extract histograms
    extractHist()

#########  TODO  ###############

#extractHist()

#createInitialMaster()

#checkFileSize()

#checkMultiSlice()

# Extract dicom tags
# folderpath_tables = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_master/discharge_tables'
# folderpath_discharge = 'G:/discharge'
# extractDICOMTags(folderpath_tables, folderpath_discharge)