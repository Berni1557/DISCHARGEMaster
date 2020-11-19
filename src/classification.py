# -*- coding: utf-8 -*-
import os, sys
import pandas as pd
import numpy as np
from collections import defaultdict
from ActiveLearner import DISCHARGEFilter, ActiveLearner

def createRFClassification(settings):
    print('Create RF Classification')
    # date = folderpath_master.split('_')[-1]
    # folderpath_components = os.path.join(folderpath_master, 'discharge_components_' + date)
    # if not os.path.isdir(folderpath_components):
    #     os.mkdir(folderpath_components)

    #filepath_data = os.path.join(settings['folderpath_components'], 'discharge_data_' + date + '.xlsx')
    #filepath_rfc = os.path.join(folderpath_components, 'discharge_rfc_' + date + '.xlsx')
    df_data = pd.read_excel(settings['filepath_data'])
    # Repace Count for multi-slice format
    #idx=df_data['NumberOfFrames']>0
    #df_data['Count'][idx] = df_data['NumberOfFrames'][idx]
    
    df_rfc0 = pd.DataFrame('UNDEFINED', index=np.arange(len(df_data)), columns=['RFCLabel'])
    df_rfc1 = pd.DataFrame('UNDEFINED', index=np.arange(len(df_data)), columns=['RFCClass'])
    df_rfc2 = pd.DataFrame(0, index=np.arange(len(df_data)), columns=['RFCConfidence'])
    df_rfc = pd.concat([df_rfc0, df_rfc1, df_rfc2], axis=1)
    df_rfc.to_excel(settings['filepath_rfc'])

def initRFClassification(folderpath_master, master_process=False):

    filepath_master = settings['filepath_master']
    sheet_name = 'MASTER_' + settings['date']
    if os.path.exists(filepath_master):
        df_master = pd.read_excel(filepath_master, sheet_name=sheet_name, index_col=0)
        
        # Create active learner
        learner = ActiveLearner()
        target = featureSelection(filtername='CACSFilter_V02')
        discharge_filter = target['FILTER']
        
        # Extract features
        #learner.extractFeatures(df_master, discharge_filter)
        
        #filepath_hist = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_master/discharge_master_01042020/discharge_master_01042020_hist.pkl'
        dfHist = pd.read_pickle(filepath_hist)
        dfData = pd.concat([dfHist.iloc[:,3:], dfHist.iloc[:,1]],axis=1)
        X = np.array(dfData)
        scanClassesRF = defaultdict(lambda:-1,{'CACS': 0, 'CTA': 1, 'NCS_CACS': 2, 'NCS_CTA': 3, 'OTHER': 4})
        #scanClassesRFInv = defaultdict(lambda:'',{'CACS': 1, 'CTA': 2, 'NCS_CACS': 3, 'NCS_CTA': 4})
        scanClassesRFInv = defaultdict(lambda:'UNDEFINED' ,{0: 'CACS', 1: 'CTA', 2: 'NCS_CACS', 3: 'NCS_CTA', 4: 'OTHER'})
        
        #Y = [scanClassesRF[x] for x in list(dfHist['CLASS'])]
        Y = [scanClassesRF[x] for x in list(df_master['CLASS'])]
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
        
def classifieRFClassification(settings):
    

    filepath_master = settings['filepath_master']
    sheet_name = 'MASTER_' + settings['date']
    if os.path.exists(filepath_master):
        df_master = pd.read_excel(filepath_master, sheet_name=sheet_name, index_col=0)
        
        # Create active learner
        learner = ActiveLearner()
        target = featureSelection(filtername='CACSFilter_V02')
        discharge_filter = target['FILTER']

        #filepath_hist = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_master/discharge_master_01042020/discharge_master_01042020_hist.pkl'
        dfHist = pd.read_pickle(filepath_hist)
        
        df0 = df_master[['SeriesInstanceUID','RFCLabel']]
        df_merge = df0.merge(dfHist, on=['SeriesInstanceUID', 'SeriesInstanceUID'])
        
        #dfHist = dfHist[idx_ct]
        #dfData = pd.concat([dfHist.iloc[:,3:], dfHist.iloc[:,1]],axis=1)
        #dfData = pd.concat([df_merge.iloc[:,4:104], df_merge.iloc[:,2]],axis=1)
        dfData = pd.concat([df_merge.iloc[:,4:104]],axis=1)
        X = np.array(dfData)
        scanClassesRF = defaultdict(lambda:-1,{'CACS': 0, 'CTA': 1, 'NCS_CACS': 2, 'NCS_CTA': 3, 'OTHER': 4})
        #scanClassesRFInv = defaultdict(lambda:'',{'CACS': 1, 'CTA': 2, 'NCS_CACS': 3, 'NCS_CTA': 4})
        scanClassesRFInv = defaultdict(lambda:'UNDEFINED' ,{0: 'CACS', 1: 'CTA', 2: 'NCS_CACS', 3: 'NCS_CTA', 4: 'OTHER'})

        idx_manual = ~(df_master['ClassManualCorrection']=='UNDEFINED')
        RFCLabel = df_master['RFCLabel']
        RFCLabel[idx_manual] = df_master['ClassManualCorrection'][idx_manual]
    
        Y = [scanClassesRF[x] for x in list(df_merge['RFCLabel'])]
        Y = np.array(Y)
        X = np.where(X=='', -1, X)
        
        Target = 'RFCLabel'

        learner.df_features = X

        if sum(Y>0)>0:

            df_merge[Target] = Y
            
            # Predict random forest
            confidence, C, ACC, pred_class, df_features = learner.confidencePredictor(df_merge, discharge_filter, Target = Target)
            print('Confusion matrix:', C)
            pred_class = [scanClassesRFInv[x] for x in list(pred_class)]
            
            print('df_mastershape', df_master.shape)
            print('dfDatashape', dfData.shape)
            
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
        
        
def featureSelection(filtername='CACSFilter_V02', filt='CLASS_CACS_alt0'):
    if filtername=='CACSFilter_V02':
        ReconstructionDiameterFilter = DISCHARGEFilter()
        ReconstructionDiameterFilter.createFilter(feature='ReconstructionDiameter', name='ReconstructionDiameter', featFunc=lambda v : v)
        # Create SliceThickness filter for 3.0 mm
        SliceThicknessFilter = DISCHARGEFilter()
        SliceThicknessFilter.createFilter(feature='SliceThickness', name='SliceThicknessFilter', featFunc=lambda v : v)

        # Create site filter
        SiteFilter = DISCHARGEFilter()
        SiteFilter.createFilter(feature='Site', name='SiteFilter', featFunc=lambda v : float(v['Site'][1:]))
        #SiteFilter.createFilter(feature='Site', name='SiteFilter', featFunc=lambda v : print('v', v['Site']))
        # Create Modality Filter
        ModalityFilter = DISCHARGEFilter()
        ModalityFilter.createFilter(feature='Modality', name='ModalityFilter', featFunc=lambda v : v == 'CT')
        # Create ProtocolName 
        ProtocolNameFilter = DISCHARGEFilter()
        ProtocolNameFilter.createFilter(feature='ProtocolName', updateTarget=True, name='ProtocolName_Calcium Score', featFunc=DISCHARGEFilter.includeStringList(
            ['Calcium Score','CaScoring', 'CaScoring','CACS','CaScSeq','SMART SCORE','Calsium score','Calcium Score DISCHARGE','ca score', 'CAVE', 'CALCIUM SCORE']))
        # Create CountFilter 
        CountFilter = DISCHARGEFilter()
        CountFilter.createFilter(feature='Count', name='CountFilter', featFunc=lambda v : v)
        
        words_SeriesDescription = ['CTA', 'REMOVED','Calcium Score','CaScoring', 'CaScoring','CACS','CaScSeq','SMART SCORE','Calsium score','Calcium Score DISCHARGE','ca score', 'CAVE', 'CALCIUM SCORE','LUNG']
        SeriesDescriptionFilter = DISCHARGEFilter()
        SeriesDescriptionFilter.createFilter(feature='SeriesDescription', name='SeriesDescription', featFunc=DISCHARGEFilter.StringIndex(words_SeriesDescription, name='SeriesDescription'))
        
        words_ProtocolName = ['CTA','REMOVED', 'Calcium Score','CaScoring', 'CaScoring','CACS','CaScSeq','SMART SCORE','Calsium score','Calcium Score DISCHARGE','ca score', 'CAVE', 'CALCIUM SCORE','LUNG']
        ProtocolNameFilter = DISCHARGEFilter()
        ProtocolNameFilter.createFilter(feature='ProtocolName', name='ProtocolName', featFunc=DISCHARGEFilter.StringIndex(words_ProtocolName, name='ProtocolName'))
        
        words_ContrastBolusAgent = ['APPLIED', 'Iodine', 'CE', 'Omnipaque', 'Ultravist', 'REMOVED', 'Deldentified', 'ml', 'REMOVED', 'NONE']
        ContrastBolusAgentFilter = DISCHARGEFilter()
        ContrastBolusAgentFilter.createFilter(feature='ContrastBolusAgent', name='ContrastBolusAgent', featFunc=DISCHARGEFilter.StringIndex(words_ContrastBolusAgent, name='ContrastBolusAgent'))
        
        words_ImageComments = ['CTA','REMOVED', 'Calcium Score','CaScoring', 'CaScoring','CACS','CaScSeq','SMART SCORE','Calsium score','Calcium Score DISCHARGE','ca score', 'CAVE', 'CALCIUM SCORE', 'LUNG']
        ImageCommentsFilter = DISCHARGEFilter()
        ImageCommentsFilter.createFilter(feature='ImageComments', name='ImageComments', featFunc=DISCHARGEFilter.StringIndex(words_ImageComments, name='ImageComments'))
        
        ITTFilter = DISCHARGEFilter()
        ITTFilter.createFilter(feature='ITT', name='ITT', featFunc=lambda v : v)
        
        #words_ConvolutionKernel = ['FC51', 'FC03', 'FC17', 'FC12' , 'B35f', 'B19f', 'B30f', 'FL05', 'T20f']
        words_ConvolutionKernelFilter = ['B35f', 'Qr36', 'FC12', 'FC03', 'FC51', 'FC17', 'STANDART', 'LUNG', 'B60', 'B46', 'B70', 'A', 'B', 'CB']
        ConvolutionKernelFilter = DISCHARGEFilter()
        ConvolutionKernelFilter.createFilter(feature='ConvolutionKernel', name='ConvolutionKernel', featFunc=DISCHARGEFilter.StringIndex(words_ConvolutionKernelFilter, name='ConvolutionKernel'))

        # Append filter
        discharge_filter=[]
        discharge_filter.append(ReconstructionDiameterFilter)
        discharge_filter.append(SliceThicknessFilter)
        discharge_filter.append(SiteFilter)
        discharge_filter.append(ModalityFilter)
        discharge_filter.append(ContrastBolusAgentFilter)
        discharge_filter.append(CountFilter)
        discharge_filter.append(ProtocolNameFilter)
        discharge_filter.append(SeriesDescriptionFilter)
        discharge_filter.append(ImageCommentsFilter)
        discharge_filter.append(ITTFilter)
        discharge_filter.append(ConvolutionKernelFilter)
        
        target = defaultdict(lambda: None, {'FILTER': discharge_filter, 'TARGET': 'CACS_alt0', 'FONT_COLOR': 'blue', 'BG_COLOR': 'white'})
        
        return target
    
    elif filtername=='CACS-RECOFilter_ORG_V02':
        
        words_SeriesDescription = ['AIDR', 'IR', 'ASiR', 'ORG', 'FBP', 'IDOSE','IMR']
        SeriesDescriptionFilter = DISCHARGEFilter()
        SeriesDescriptionFilter.createFilter(feature='SeriesDescription', name='SeriesDescription', featFunc=DISCHARGEFilter.StringIndex(words_SeriesDescription, name='SeriesDescription'))

        discharge_filter=[]
        discharge_filter.append(SeriesDescriptionFilter)
        target = defaultdict(lambda: None, {'FILTER': discharge_filter, 'TARGET': 'CACS-RECO-ORG', 'FONT_COLOR': None, 'BG_COLOR': 'yellow', 'FILT': filt})
        return target
    elif filtername=='CACS-RECOFilter_IR_V02':
        
        words_SeriesDescription = ['AIDR', 'IR', 'ASiR', 'ORG', 'FBP', 'IDOSE', 'IMR']
        SeriesDescriptionFilter = DISCHARGEFilter()
        SeriesDescriptionFilter.createFilter(feature='SeriesDescription', name='SeriesDescription', featFunc=DISCHARGEFilter.StringIndex(words_SeriesDescription, name='SeriesDescription'))

        discharge_filter=[]
        discharge_filter.append(SeriesDescriptionFilter)
        target = defaultdict(lambda: None, {'FILTER': discharge_filter, 'TARGET': 'CACS-RECO-IR', 'FONT_COLOR': None, 'BG_COLOR': 'green', 'FILT': filt})
        return target
    
    else:
        raise ValueError('Filtername: ' + filtername + ' does not exist.')