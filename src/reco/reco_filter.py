# -*- coding: utf-8 -*-
import sys,os
sys.path.append('H:/cloud/cloud_data/Projects/DL/Code/src')
from settings import initSettings, saveSettings, loadSettings, fillSettingsTags
import numpy as np
import SimpleITK as sitk
import pandas as pd
from ct.CTDataStruct import CTPatient, CTImage
from collections import defaultdict
from sklearn.tree import DecisionTreeClassifier
from sklearn.metrics import confusion_matrix
from sklearn.metrics import accuracy_score
from sklearn.ensemble import RandomForestClassifier
import matplotlib.pyplot as plt

def gradientmagitude(image):
    
    gmif = sitk.GradientMagnitudeImageFilter()
    gmif.SetUseImageSpacing(True)
    image_std = gmif.Execute(image)
    
    image_np = sitk.GetArrayFromImage(image)
    image_std_np = sitk.GetArrayFromImage(image_std)

    mask0 = image_std_np<120
    mask1 = (image_np>-200) & (image_np<130)
    mask = mask0 * mask1
    
    image_std_np_mask = image_std_np * mask
    values = image_std_np_mask[image_std_np_mask>0]
    #std = np.std(values)
    std = np.mean(values)
    #print('std', std)
    
    plt.imshow(image_std_np_mask[34,:,:])
    plt.show()
    
    plt.imshow(image_np[34,:,:])
    plt.show()
    
    return std, image_std_np_mask


def extractFeatures(settings):
    # Read CACS dataframe
    name = 'CACS_20210801'
    folderpath = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/datasets'
    folderpath_data = os.path.join(folderpath, name)
    filepath_dataset = os.path.join(folderpath_data, name +'.xlsx')
    folderpath_discharge = settings['folderpath_discharge']
    df_dataset = pd.read_excel(filepath_dataset)
    df_cacs = df_dataset[df_dataset['CACSSelection']==1]
    
    columns =['SeriesInstanceUID', 'Count', 'CLASS', 'RFCConfidence', 'RECO_CLASS', 'Features']
    dfFeatures = pd.DataFrame(columns=columns) 
    
    for index, row in df_cacs.iterrows():
        print('index', index)    
        PatientID = df_cacs.loc[index, 'PatientID']
        StudyInstanceUID = df_cacs.loc[index, 'StudyInstanceUID']
        SeriesInstanceUID = df_cacs.loc[index, 'SeriesInstanceUID']
        RECO_CLASS = df_cacs.loc[index, 'RECO']
        ReconstructionDiameter = df_cacs.loc[index, 'ReconstructionDiameter']
        filepath_image = os.path.join(folderpath_discharge, StudyInstanceUID, SeriesInstanceUID)
        #imagect = CTImage()
        #imagect.load(filepath_image)
        #if len(imagect.image().shape)==3:                 
        if True:    
            #shape = imagect.image().shape
            shape=(3,0,0)
            #features, _ = gradientmagitude(imagect.image_sitk)
            features = 0
            dfFeatures = dfFeatures.append({'PatinetID': PatientID, 
                                            'StudyInstanceUID': StudyInstanceUID, 
                                            'SeriesInstanceUID': SeriesInstanceUID, 
                                          'Count': shape[0],
                                          #'CLASS': RFCClass,
                                          'ReconstructionDiameter':ReconstructionDiameter,
                                          'RECO_CLASS': RECO_CLASS,
                                          'Center': int(PatientID[0:2]),
                                          'Features': features},ignore_index=True,)
                    
        if len(dfFeatures) % 50 == 0:
            dfFeatures.to_pickle(settings['filepath_reco'])
    dfFeatures.to_pickle(settings['filepath_reco'])
    
    
    
    
    dfFeatures = pd.read_pickle('H:/cloud/cloud_data/Projects/DISCHARGEMaster/data/discharge_master/discharge_master_01092020/discharge_components_01092020/discharge_reco_01092020_V01.pkl')
    dfFeatures0 = pd.read_pickle('H:/cloud/cloud_data/Projects/DISCHARGEMaster/data/discharge_master/discharge_master_01092020/discharge_components_01092020/discharge_reco_01092020.pkl')
    
    dfFeatures.reset_index(inplace=True, drop=True)
    dfFeatures0.reset_index(inplace=True, drop=True)
    
    dfFeatures['ReconstructionDiameter'] = dfFeatures0['ReconstructionDiameter']
    dfFeatures.to_pickle(settings['filepath_reco'])
    
    

def classifie(settings):
    
    dfFeatures = pd.read_pickle(settings['filepath_reco'])
    dfFeatures.reset_index(inplace=True, drop=True)
    
    X = np.zeros((dfFeatures.shape[0], 4))
    for index, row in dfFeatures.iterrows():
        X[index,:] = [row['Features'], row['Count'], row['Center'], row['ReconstructionDiameter']]
    recoClassRF = defaultdict(lambda: -1, {'UNDEFINED':0, 'FBP': 1, 'IR': 2})
    
    Y = [recoClassRF[x] for x in list(dfFeatures['RECO_CLASS'])]
    
    X = np.array(X)
    Y = np.array(Y)
    X = np.nan_to_num(X)

    # Filter by class label
    X_train=X[Y>0]
    Y_train=Y[Y>0]
    dfFeatures_train = dfFeatures[Y>0]

    # Train random forest
    #clf = DecisionTreeClassifier(max_depth=5)
    #clf = clf.fit(X_train, Y_train)
    
    clf = RandomForestClassifier(max_depth=10, n_estimators=100)
    clf.fit(X_train, Y_train)
    
    # Predict classifier
    pred_train = clf.predict(X_train)
    
    # Extract confusion matrix and accuracy
    
    C = confusion_matrix(pred_train, Y_train)
    ACC = accuracy_score(pred_train, Y_train)
    
    # Predict confidence
    prop = clf.predict_proba(X_train)
    pred = clf.predict(X_train)
    thr = 1/(Y_train.max()+1)
    confidence = (np.max(prop, axis=1)-thr)*(1/(1-thr))

    dfFeatures_train.reset_index(drop=True, inplace=True)
    TC = pd.DataFrame(columns=['TrueClass'])
    TC['TrueClass'] = Y_train
    PC = pd.DataFrame(columns=['PredClass'])
    PC['PredClass'] = pred_train
    PROP = pd.DataFrame(columns=['Prop'])
    PROP['Prop'] = prop.max(axis=1)
    PROP.reset_index(drop=True, inplace=True)
    dfFeatures_train=pd.concat([dfFeatures_train,TC,PC, PROP],axis=1)
    
    print('Confusion matrix:', C)
    print('Confidence:', confidence)
    print('ACC:', ACC)
    
    # Apply patient rule
    k=0
    patientList = pd.unique(dfFeatures_train['PatinetID'])
    for patient in patientList:
        dfPat = dfFeatures_train[dfFeatures_train['PatinetID']==patient]
        if len(dfPat)==2:
            if dfPat.iloc[0]['Count']==dfPat.iloc[1]['Count']:
                k = k+1
                print('k', k)
                if dfPat.iloc[0]['Features']>dfPat.iloc[1]['Features']:
                    idx0 = dfPat.index[0]
                    idx1 = dfPat.index[1]
                    dfFeatures_train.loc[idx0,'PredClass'] = 1
                    dfFeatures_train.loc[idx1,'PredClass'] = 2
                if dfPat.iloc[0]['Features']<dfPat.iloc[1]['Features']:
                    idx0 = dfPat.index[0]
                    idx1 = dfPat.index[1]
                    dfFeatures_train.loc[idx0,'PredClass'] = 2
                    dfFeatures_train.loc[idx1,'PredClass'] = 1    
    
    pred_train = np.array(dfFeatures_train['PredClass'])
    C = confusion_matrix(pred_train, Y_train)
    ACC = accuracy_score(pred_train, Y_train)
                
    
    return clfRF, confidence, C, ACC, pred


# Load settings
filepath_settings = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/data/settings.json'
settings=initSettings()
saveSettings(settings, filepath_settings)
settings = fillSettingsTags(loadSettings(filepath_settings))

#extractFeatures(settings)
#clfRF, confidence, C, ACC, pred = classifie(settings)