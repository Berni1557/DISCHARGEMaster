# -*- coding: utf-8 -*-
import sys,os
sys.path.append('H:/cloud/cloud_data/Projects/DL/Code/src')
from settings import initSettings, saveSettings, loadSettings, fillSettingsTags
import numpy as np
import SimpleITK as sitk
import pandas as pd
from ct.CTDataStruct import CTPatient, CTImage, CTRef
from collections import defaultdict
from sklearn.tree import DecisionTreeClassifier
from sklearn.metrics import confusion_matrix
from sklearn.metrics import accuracy_score
from sklearn.ensemble import RandomForestClassifier
import matplotlib.pyplot as plt
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.styles import Font, Color, Border, Side
from openpyxl.styles import Protection
from openpyxl.styles import PatternFill

def gradientmagitude(imagect, RFCLabel):
    
    image = imagect.image_sitk
    # Create GradientMagnitudeImageFilter
    gmif = sitk.GradientMagnitudeImageFilter()
    gmif.SetUseImageSpacing(True)
    image_std = gmif.Execute(image)
    
    image_np = sitk.GetArrayFromImage(image)
    image_std_np = sitk.GetArrayFromImage(image_std)

    if RFCLabel == 'CACS':
        image_std_np_min = 120
        image_np_min = -200
        image_np_max = 130
    elif RFCLabel == 'NCS_CACS':
        image_std_np_min = 120
        image_np_min = -200
        image_np_max = 130
    elif RFCLabel == 'CTA':
        image_std_np_min = 1000
        image_np_min = -300
        image_np_max = 1000
    elif RFCLabel == 'NCS_CTA':
        image_std_np_min = 1000
        image_np_min = -300
        image_np_max = 1000
    else:
        print('RFCLabel ' + RFCLabel + ' not processed.')
        return None, None
    
    mask0 = (image_std_np < image_std_np_min)
    mask1 = ((image_np > image_np_min) & (image_np < image_np_max))
    mask = mask0 * mask1
    
    image_std_np_mask = image_std_np * mask
    values = image_std_np_mask[image_std_np_mask > 0]
    std = np.mean(values)
    
    # plt.imshow(image_std_np_mask[34,:,:])
    # plt.show()
    
    # plt.imshow(image_np[34,:,:])
    # plt.show()
    
    return std, image_std_np_mask


def extractFeatures(settings, extract_continue=True):
    
    # Extract important settings
    folderpath_discharge = settings['folderpath_discharge']
    
    # Read master
    df_master = pd.read_excel(settings['filepath_master'], 'MASTER_' + settings['date'])
    RFCLabel = df_master['RFCLabel']
    
    columns_features = ['Site', 'PatientID', 'StudyInstanceUID', 'SeriesInstanceUID', 'RFCLabel', 'SliceThickness', 'Count', 'RECO_CLASS', 'ImageFeatures']
    df_features = pd.DataFrame(columns=columns_features) 
    
    # Load df_features
    if os.path.exists(settings['filepath_reco']) and extract_continue:
        df_features = pd.read_pickle(settings['filepath_reco'])
    
    for index, row in df_master.iterrows():
        print('index', index)  
        
        if df_master.loc[index, 'SeriesInstanceUID']== '1.2.392.200036.9116.2.2426555318.1447120890.30.10453000170.1':
            sys.exit()
            
        if index > len(df_features):
            PatientID = df_master.loc[index, 'PatientID']
            StudyInstanceUID = df_master.loc[index, 'StudyInstanceUID']
            SeriesInstanceUID = df_master.loc[index, 'SeriesInstanceUID']
            RFCLabel = df_master.loc[index, 'RFCLabel']
            SliceThickness = df_master.loc[index, 'SliceThickness']
            Count = df_master.loc[index, 'Count']
            RECO_CLASS = df_master.loc[index, 'RECO']
            site = df_master.loc[index, 'Site']
            
            try:
                # Load image
                imagect = CTImage()
                filepath_image = os.path.join(folderpath_discharge, StudyInstanceUID, SeriesInstanceUID)
                imagect.load(filepath_image)
                
                if len(imagect.image().shape)==3:
                    # Extract image features
                    ImageFeatures, image_std_np_mask = gradientmagitude(imagect, RFCLabel)
                else:
                    ImageFeatures = -1
            except:
                ImageFeatures = -1
                

                
            # Plot image_std_np_mask
            # slice_num = 100
            # plt.imshow(image_std_np_mask[slice_num,:,:])
            # plt.show()
            
            # Create feature 
            df_features = df_features.append({'Site': site, 
                                            'PatientID': PatientID, 
                                            'StudyInstanceUID': StudyInstanceUID, 
                                            'SeriesInstanceUID': SeriesInstanceUID,
                                            'RFCLabel': RFCLabel,
                                            'SliceThickness': SliceThickness,
                                            'Count': Count,
                                            'RECO_CLASS': RECO_CLASS,
                                            'ImageFeatures': ImageFeatures,
                                            'TrueClass': -1,
                                            'PredClass': -1,
                                            'Prop': -1},ignore_index=True)
        
        if len(df_features) % 50 == 0:
            df_features.to_pickle(settings['filepath_reco'])
            
    df_features.to_pickle(settings['filepath_reco'])
    


def classifie(settings, ModFilter=['CACS'], SliceThicknessFilt=[3.0]):
    
    df_features = pd.read_pickle(settings['filepath_reco'])
    df_features.reset_index(inplace=True, drop=True)
    
    # def func_site(site):
    #     return int(site[1:])
    
    dict_RFCLabel = defaultdict(lambda: -1, {'CACS':0, 'NCS_CACS':1, 'CTA':2, 'NCS_CTA':3})
    dict_RECO_CLASSF = defaultdict(lambda: -1, {'FBP':0, 'IR':1})
    
    X = np.ones((df_features.shape[0], 5))*-1
    Y = np.ones((df_features.shape[0], 1))*-1
    for index, row in df_features.iterrows():
        if row['ImageFeatures'] and row['Site'][1:].isnumeric() and row['RFCLabel'] in ModFilter and row['SliceThickness'] in SliceThicknessFilt:
            #sys.exit()
            SiteF = int(row['Site'][1:]) 
            RFCLabelF = dict_RFCLabel[row['RFCLabel']]
            SliceThicknessF = float(row['SliceThickness'])
            CountF = int(row['Count'])
            RECO_CLASSF = dict_RECO_CLASSF[row['RECO_CLASS']]
            ImageFeaturesF = float(row['ImageFeatures'])
            X[index,:] = [SiteF, RFCLabelF, SliceThicknessF, CountF, ImageFeaturesF]
            Y[index,:] = [RECO_CLASSF]
        else:
            X[index,:] = [-1, -1, -1, -1, -1]
            Y[index,:] = [-1]

    X = np.array(X)
    Y = np.array(Y)[:,0]
    X = np.nan_to_num(X)

    # Filter by class label
    X_train=X[Y>=0]
    Y_train=Y[Y>=0]
    
    # Train random forest
    clf = RandomForestClassifier(max_depth=10, n_estimators=100)
    clf.fit(X_train, Y_train)

    # Predict classifier
    Y_train_pred = clf.predict(X_train)
    
    # Extract confusion matrix and accuracy
    C = confusion_matrix(Y_train_pred, Y_train)
    ACC = accuracy_score(Y_train_pred, Y_train)
    
    # Predict confidence
    prop_train = clf.predict_proba(X_train)
    confidence = np.max(prop_train, axis=1)
    print('Confusion matrix:', C)
    #print('Confidence:', confidence)
    print('ACC:', ACC)
    
    # Add predictions
    Y_pred = clf.predict(X)
    prop = clf.predict_proba(X)
    df_features_pred = df_features.copy()
    for index, row in df_features_pred.iterrows():
        df_features_pred.loc[index, 'TrueClass'] = Y[index]
        df_features_pred.loc[index, 'PredClass'] = Y_pred[index]
        df_features_pred.loc[index, 'Prop'] = prop[index].max()
    
    df_features_pred.to_pickle(settings['df_features_pred'])
    
    # Save excel
    writer = pd.ExcelWriter(settings['df_features_pred_excel'], engine="openpyxl", mode="w")
    sheet_name = 'reco'
    workbook  = writer.book
    df_features_pred.to_excel(writer, sheet_name=sheet_name)
    writer.save()


    return clf, confidence, C, ACC, df_features_pred

def createReco():
    # Load settings
    filepath_settings = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/data/settings.json'
    settings=initSettings()
    saveSettings(settings, filepath_settings)
    settings = fillSettingsTags(loadSettings(filepath_settings))
    settings['df_features_pred'] = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/data/discharge_master/discharge_master_01092020/discharge_components_01092020/discharge_reco_pred_01092020.pkl'
    settings['df_features_pred_excel'] = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/data/discharge_master/discharge_master_01092020/discharge_components_01092020/discharge_reco_pred_01092020.xlsx'
    
    clfRF, confidence, C, ACC, df_features_pred = classifie(settings)
    
def update_cacs():
    settings=initSettings()
    saveSettings(settings, filepath_settings)
    settings = fillSettingsTags(loadSettings(filepath_settings))
    settings['df_features_pred'] = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/data/discharge_master/discharge_master_01092020/discharge_components_01092020/discharge_reco_pred_01092020.pkl'
    settings['df_features_pred_excel'] = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/data/discharge_master/discharge_master_01092020/discharge_components_01092020/discharge_reco_pred_01092020.xlsx'
    settings['df_cacs'] = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/datasets/CACS_20210801/CACS_20210801.xlsx'
    settings['df_cacs_reco'] = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/datasets/CACS_20210801/CACS_reco_20210801.xlsx'
    #filepath_cacs = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/datasets/CACS_20210801/CACS_20210801.xlsx'
    #filepath_cacs_reco = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/datasets/CACS_20210801/CACS_20210801_reco.xlsx'
    #filepath_reco_save =  'H:/cloud/cloud_data/Projects/DISCHARGEMaster/datasets/CACS_20210801/CACS_20210801_reco_pred.xlsx'
    df_cacs = pd.read_excel(settings['df_cacs'], index_col=0)
    df_reco_read = pd.read_excel(settings['df_features_pred_excel'], index_col=0)
    df_reco = pd.DataFrame()
    df_cacs = df_cacs.drop(['RECO'], axis=1)
    df_reco['RECO'] = df_reco_read['PredClass']
    df_reco['RECO_TRUE'] = df_reco_read['TrueClass']
    df_reco['RECO_PROP'] = df_reco_read['Prop']
    df_reco['SeriesInstanceUID'] = df_reco_read['SeriesInstanceUID']
    df_cacs_merge = df_cacs.merge(df_reco, on=['SeriesInstanceUID', 'SeriesInstanceUID'], how='inner')
    
    for index, row in df_cacs_merge.iterrows():
        print(index)
        if not ((row['CACSSelection'] == 1) and (row['SliceThickness'] == 3.0)):
            df_cacs_merge.loc[index, 'RECO'] = ''
            df_cacs_merge.loc[index, 'RECO_PROP'] = ''
        
    # Save excel
    writer = pd.ExcelWriter(settings['df_cacs_reco'], engine="openpyxl", mode="w")
    sheet_name = 'reco'
    workbook  = writer.book
    df_cacs_merge.to_excel(writer, sheet_name=sheet_name)
    writer.save()
    
# #extractFeatures(settings, extract_continue=True)
#clfRF, confidence, C, ACC, df_features_pred = classifie(settings)

# # Check rsults
# df_master = pd.read_excel(settings['filepath_master'], 'MASTER_' + settings['date'])
# df_reco = pd.read_pickle('H:/cloud/cloud_data/Projects/DISCHARGEMaster/data/discharge_master/discharge_master_01092020/discharge_components_01092020/discharge_reco_pred_01092020.pkl')

# # Update SeriesNumber
# df_reco_class = pd.merge(df_reco, df_master[['SeriesInstanceUID', 'SeriesNumber']] , on=["SeriesInstanceUID"])


# filepath_reco_excel = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/data/discharge_master/discharge_master_01092020/discharge_components_01092020/discharge_reco_pred_01092020.xlsx'
# writer = pd.ExcelWriter(filepath_reco_excel, engine="openpyxl", mode="w")
# sheet_name = 'reco'
# workbook  = writer.book
# df_reco_class.to_excel(writer, sheet_name=sheet_name)
# writer.save()


        
    
    
    
    
    

