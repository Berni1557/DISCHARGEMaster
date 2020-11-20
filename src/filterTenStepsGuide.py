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

def filter_10StepsGuide(df_master_in):
    df_master = df_master_in.copy()
    df_master.replace(to_replace=[np.nan], value=0.0, inplace=True)
    #folderpath_master = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_master/discharge_master_01042020'
    # date = folderpath_master.split('_')[-1]
    # if master_process==False:
    #     filepath_master = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')
    # else:
    #     filepath_master_tmp = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')
    #     folderpath, filename, file_extension = splitFilePath(filepath_master_tmp)
    #     filepath_master = os.path.join(folderpath, filename + '_process' + file_extension)
    # #df_master = pd.read_excel(filepath_master, sheet_name='MASTER_'+ date, index_col=0)
    # df_master.replace(to_replace=[np.nan], value=0.0, inplace=True)
    
    idx_manual = ~(df_master['ClassManualCorrection']=='UNDEFINED')
    SeriesClass = df_master['RFCClass']
    SeriesClass[idx_manual]=df_master['ClassManualCorrection'][idx_manual]
    
    df_guide = pd.DataFrame(columns=['GUIDE', 'COMMENT'])

    for index, row in df_master.iterrows():
        print('index', index)
        guide=True
        comment=''
        if SeriesClass[index]=='CACS':
            # Check 10-steps guide for CACS
            if row['ReconstructionDiameter'] > 200:
                guide=False
                comment = comment + 'ReconstructionDiameter bigger than 200mm,'
            if not (row['SliceThickness']==3.0) or (row['SliceThickness']==2.5 and row['Site'] in ['P10', 'P13', 'P29']):
                guide=False
                comment = comment + 'SliceThickness not correct,'
            if not row['Modality'] == 'CT':
                guide=False
                comment = comment + 'Modality is not CT,'
        elif SeriesClass[index]=='CTA':
            ConvFilter = 'B35s|Qr36d|FC12|FC51|FC17|B60f|B70f|B30f|B31f|B08s|B19f|B20f|B20s|B30s|B31s|B40f|B41s|B50f|B50s|B65f|B70f|B70s|B80s|Bf32dB80s|Bf32d|Bl57d|Br32d|Bv36d|Bv40f|FC08|FC08-H|FC15|FC18|FC35|FC52|FL03|FL04|FL05|H20f|H31s|IMR1|IMR2|IMR2|Qr36d|T20f|T20s|Tr20f|UB|XCA|YA'
            ConvFilterList = ConvFilter.split("|")
            for filt in ConvFilterList:
                if filt in str(row['ConvolutionKernel']):
                    guide=False
                    comment = comment + 'FilterKernel ' + filt + ' in ConvolutionKernel,'
                if filt in str(row['SeriesDescription']):
                    guide=False
                    comment = comment + 'SeriesDescription ' + filt + ' in SeriesDescription,'
            if  row['ReconstructionDiameter'] > 260:
                guide=False
                comment = comment + 'ReconstructionDiameter bigger than 260mm,'
            if row['SliceThickness'] >= 0.8:
                guide=False
                comment = comment + 'SliceThickness bigger than 0.8mm,'
        elif SeriesClass[index]=='NCS_CACS':
            if row['ReconstructionDiameter'] < 320:
                guide=False
                comment = comment + 'ReconstructionDiameter smaller than 320mm,'
            if not (row['SliceThickness']==1.0) or (row['SliceThickness']==0.625 and row['Site'] in ['P10', 'P13', 'P29']):
                guide=False
                comment = comment + 'SliceThickness not correct,'
        elif SeriesClass[index]=='NCS_CTA':
            if row['ReconstructionDiameter'] < 320:
                guide=False
                comment = comment + 'ReconstructionDiameter smaller than 320mm,'
            if not (row['SliceThickness']==1.0) or (row['SliceThickness']==0.625 and row['Site'] in ['P10', 'P13', 'P29']):
                guide=False
                comment = comment + 'SliceThickness not correct,'
        else:
            pass
        df_guide = df_guide.append(dict({'GUIDE': guide, 'COMMENT': comment}), ignore_index=True)
        
    return df_guide

def filer10StepsGuide(settings):
    
    # date = folderpath_master.split('_')[-1]
    # folderpath_components = os.path.join(folderpath_master, 'discharge_components_' + date)
    # if master_process==False:
    #     filepath_master = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')
    # else:
    #     filepath_master_tmp = os.path.join(folderpath_master, 'discharge_master_' + date + '.xlsx')
    #     folderpath, filename, file_extension = splitFilePath(filepath_master_tmp)
    #     filepath_master = os.path.join(folderpath, filename + '_process' + file_extension)
    
    filepath_master = settings['filepath_master']
    df_master = pd.read_excel(filepath_master, sheet_name='MASTER_'+ settings['date'], index_col=0)
    #df_master.replace(to_replace=[np.nan], value=0.0, inplace=True)
    
    # Filter according to 10-Steps guide
    df_guide = filter_10StepsGuide(df_master)
    #df_guide = pd.DataFrame(columns=['10-STEPS-GUIDE', '10-STEPS-GUIDE-COMMENT'])
    df_master['10-STEPS-GUIDE'] = df_guide['GUIDE']
    df_master['10-STEPS-GUIDE-COMMENT'] = df_guide['COMMENT']

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