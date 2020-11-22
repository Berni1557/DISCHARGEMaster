# -*- coding: utf-8 -*-
import os, sys
import pandas as pd
import numpy as np
from collections import defaultdict
from ActiveLearner import DISCHARGEFilter, ActiveLearner
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import confusion_matrix
from sklearn.metrics import accuracy_score

def createRFClassification(settings):
    print('Create RF Classification')
    #df_data = pd.read_excel(settings['filepath_data'])
    df_data = pd.read_pickle(settings['filepath_data'])
    df_rfc0 = pd.DataFrame('UNDEFINED', index=np.arange(len(df_data)), columns=['RFCLabel'])
    df_rfc1 = pd.DataFrame('UNDEFINED', index=np.arange(len(df_data)), columns=['RFCClass'])
    df_rfc2 = pd.DataFrame(0, index=np.arange(len(df_data)), columns=['RFCConfidence'])
    df_rfc = pd.concat([df_rfc0, df_rfc1, df_rfc2], axis=1)
    #df_rfc.to_excel(settings['filepath_rfc'])
    df_rfc.to_pickle(settings['filepath_rfc'])

def initRFClassification(settings):
    """ Init RF classifier
        
    :param settings: Dictionary of settings
    :type settings: dict
    """ 

    filepath_master = settings['filepath_master']
    
    sheet_name = 'MASTER_' + settings['date']
    df_master = pd.read_excel(filepath_master, sheet_name=sheet_name, index_col=0)
    df_master['RFCConfidence'] = 0.00001
    df_master['RFCClass'] = 'UNDEFINED'
    df_master['RFCLabel'] = 'UNDEFINED'

    # Write results to master
    writer = pd.ExcelWriter(filepath_master, engine="openpyxl", mode="a")
    workbook  = writer.book
    sheet = workbook[sheet_name]
    workbook.remove(sheet)
    df_master.to_excel(writer, sheet_name=sheet_name)

    

    # filepath_master = settings['filepath_master']
    # sheet_name = 'MASTER_' + settings['date']
    # if os.path.exists(filepath_master):
    #     df_master = pd.read_excel(filepath_master, sheet_name=sheet_name, index_col=0)
        
    #     # Create active learner
    #     learner = ActiveLearner()
    #     target = featureSelection(filtername='CACSFilter_V02')
    #     discharge_filter = target['FILTER']
        
    #     # Extract features
    #     #learner.extractFeatures(df_master, discharge_filter)
        
    #     #filepath_hist = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_master/discharge_master_01042020/discharge_master_01042020_hist.pkl'
    #     dfHist = pd.read_pickle(settings['filepath_hist'])
    #     dfData = pd.concat([dfHist.iloc[:,3:], dfHist.iloc[:,1]],axis=1)
    #     X = np.array(dfData)
    #     scanClassesRF = defaultdict(lambda:-1,{'CACS': 0, 'CTA': 1, 'NCS_CACS': 2, 'NCS_CTA': 3, 'OTHER': 4})
    #     #scanClassesRFInv = defaultdict(lambda:'',{'CACS': 1, 'CTA': 2, 'NCS_CACS': 3, 'NCS_CTA': 4})
    #     scanClassesRFInv = defaultdict(lambda:'UNDEFINED' ,{0: 'CACS', 1: 'CTA', 2: 'NCS_CACS', 3: 'NCS_CTA', 4: 'OTHER'})
        
    #     #Y = [scanClassesRF[x] for x in list(dfHist['CLASS'])]
    #     Y = [scanClassesRF[x] for x in list(df_master['CLASS'])]
    #     Y = np.array(Y)
    #     X = np.where(X=='', -1, X)
        
    #     Target = 'RFCLabel'
    #     df_class = df_master.copy()
    #     Y = Y[0:len(df_class[Target])]
    #     X = X[0:len(df_class[Target])]

    #     learner.df_features = X

    #     # Update data
    #     if sum(Y>0)>0:
    #         #Yarray = np.array(Y)
    #         #Yarray[Yarray==0] = -1
            
    #         df_class[Target] = Y
            
    #         # Predict random forest
    #         confidence, C, ACC, pred_class, df_features = learner.confidencePredictor(df_class, discharge_filter, Target = Target)
    #         print('Confusion matrix:', C)
    #         pred_class = [scanClassesRFInv[x] for x in list(pred_class)]
    #         df_master['RFCConfidence'] = confidence
    #         df_master['RFCClass'] = pred_class
    #         df_master['RFCLabel'] = [scanClassesRFInv[x] for x in list(Y)]
            
    #         # Write results to master
    #         writer = pd.ExcelWriter(filepath_master, engine="openpyxl", mode="a")
    #         workbook  = writer.book
    #         sheet = workbook[sheet_name]
    #         workbook.remove(sheet)
    #         df_master.to_excel(writer, sheet_name=sheet_name)
    #         writer.save()
    #     else:
    #         print('data are not labled')
    # else:
    #     print('Master', filepath_master, 'not found')      
     
    
def classifieRFClassification(settings):
    
    filepath_master = settings['filepath_master']
    sheet_name = 'MASTER_' + settings['date']
    scanClassesRF = defaultdict(lambda:-1,{'CACS': 0, 'CTA': 1, 'NCS_CACS': 2, 'NCS_CTA': 3, 'OTHER': 4})
    scanClassesRFInv = defaultdict(lambda:'UNDEFINED' ,{0: 'CACS', 1: 'CTA', 2: 'NCS_CACS', 3: 'NCS_CTA', 4: 'OTHER'})
        
    if os.path.exists(filepath_master):
        df_master = pd.read_excel(filepath_master, sheet_name=sheet_name, index_col=0)
        
        # Read histrogram
        dfHist = pd.read_pickle(settings['filepath_hist'])
        
        # Merge dataframes of hist and master
        df0 = df_master[['SeriesInstanceUID','RFCLabel']]
        df_merge = df0.merge(dfHist, on=['SeriesInstanceUID', 'SeriesInstanceUID'])
        df_data = df_merge.iloc[:,4:104]
        X = np.array(df_data)
        df0 = df_master[['SeriesInstanceUID','RFCLabel']]
        df_merge = df0.merge(dfHist, on=['SeriesInstanceUID', 'SeriesInstanceUID'])
        
        # extract data and label
        Y = np.array([scanClassesRF[x] for x in list(df_merge['RFCLabel'])])
        X = np.nan_to_num(X, nan=-1)
        
        # Filter undefined label 
        idx=Y>-1
        X_all=X[idx]
        Y_all=Y[idx]
        
        idx_all = np.array([i for i in range(0,X_all.shape[0])])
        np.random.shuffle(idx_all)
        split = int(np.round(0.7*len(idx_all)))
        idx_train = idx_all[0:split]
        idx_valid = idx_all[split:]
        
        X_train = X_all[idx_train]
        Y_train = Y_all[idx_train]
        
        X_valid = X_all[idx_valid]
        Y_valid = Y_all[idx_valid]
        # Train random forest
        clfRF = RandomForestClassifier(max_depth=20, n_estimators=300)
        clfRF.fit(X_train, Y_train)
        
        # Extract confusion matrix and accuracy
        pred_valid = clfRF.predict(X_valid)
        C_valid = confusion_matrix(pred_valid, Y_valid)
        ACC_valid = accuracy_score(pred_valid, Y_valid)
        
        print('C_valid', C_valid)
        print('ACC_valid', ACC_valid)
        
        pred = clfRF.predict(X_all)
        C = confusion_matrix(pred, Y_all)
        ACC = accuracy_score(pred, Y_all)
        
        # print classification results
        print('C', C)
        print('ACC', ACC)
            
        # Predict confidence
        prop = clfRF.predict_proba(X_train)
        pred_class = clfRF.predict(X_train)
        random_guess = 1/(Y_train.max()+1) # random_guess is probability by random guess (1/Number od classes)
        confidence = (np.max(prop, axis=1)-random_guess)*(1/(1-random_guess))
        
        pred_class = [scanClassesRFInv[x] for x in list(pred_class)]
        df_master['RFCConfidence'] = df_master['RFCConfidence'].astype(float)
        df_master['RFCConfidence'][idx] = confidence
        df_master['RFCClass'][idx] = pred_class
        
        #Write results to master
        writer = pd.ExcelWriter(filepath_master, engine="openpyxl", mode="a")
        workbook  = writer.book
        sheet = workbook[sheet_name]
        workbook.remove(sheet)
        df_master.to_excel(writer, sheet_name=sheet_name)
        writer.save()

        
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