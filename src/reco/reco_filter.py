# -*- coding: utf-8 -*-

import sys, os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import keyboard
sys.path.append('H:/cloud/cloud_data/Projects/CACSFilter/src')
sys.path.append('H:/cloud/cloud_data/Projects/DL/Code/src')
sys.path.append('H:/cloud/cloud_data/Projects/DL/Code/src/ct')
from ActiveLearner import ActiveLearner, DISCHARGEFilter
from CTDataStruct import CTPatient, CTImage
from collections import defaultdict
import math 
from glob import glob
from sklearn.metrics import accuracy_score
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import confusion_matrix

class RecoFilter(object):
    """Create RecoFilter
    """
    
    border = 20

    def __init__(self):
        """ Init CTImage
        """       
        
        self.border = 20
        self.NumFeatures = self.border * 2
            
    def fourier(self, imagect):
        image = imagect.image()
        image_slice = image[round(image.shape[0]/2),:,:]
        FS = np.abs(np.fft.fft2(image_slice, norm='ortho'))
        r = round(FS.shape[0]/2)
        c = round(FS.shape[1]/2)
        r0 = r
        r1 = r + self.border
        c0 = c
        c1 = c + self.border
        featuresX = np.reshape(FS[r0:r1, c], [-1])
        featuresY = np.reshape(FS[r, c0:c1], [-1])
        features = np.concatenate([featuresX, featuresY])
        return features

    def classifier(self, Y, X, NumTrees=100):
        """ Calculate fonfidence score using random forest classifier
        
        :param df_linear: Dataframe of the data
        :type df_linear: pd.Dataframe
        :param discharge_filter_list: List of DISCHARGEFilter
        :type discharge_filter_list: list
        :param Target: Column name of the target
        :type Target: str
        """
        
        #df = df_linear.copy()
        #df_features = df_features
            
        # Read label
        X = np.array(X)
        Y = np.array(Y)
    
        # Filter by class label
        X_train=X[Y>-1]
        Y_train=Y[Y>-1]
        
        # Replace nan by -1
        X_train = np.nan_to_num(X_train, nan=-1)
        X = np.nan_to_num(X, nan=-1)
    
        # Train random forest
        clfRF = RandomForestClassifier(max_depth=10, n_estimators=NumTrees)
        clfRF.fit(X_train, Y_train)
        
        # Extract confusion matrix and accuracy
        pred_train = clfRF.predict(X_train)
        C = confusion_matrix(pred_train, Y_train)
        ACC = accuracy_score(pred_train, Y_train)
    
        # Predict confidence
        prop = clfRF.predict_proba(X)
        #print('prop', prop[0:10])
        pred = clfRF.predict(X)
        thr = 1/(Y_train.max()+1)
        #print('thr', thr)
        confidence = (np.max(prop, axis=1)-thr)*(1/(1-thr))
    
        
        return clfRF, confidence, C, ACC, pred
                   
            
    def extractFourier(self, filepath_master, filepath_data, filepath_pred, filepath_rfc, filepath_fourier, folderpath_images, sheet_name, NumSamples):
    
        # df = pd.read_excel(filepath_master, sheet_name=sheet_name, index_col=0)
        # df.reset_index(drop=True, inplace=True)

        df_pred = pd.read_excel(filepath_pred, index_col=0)
        df_pred.reset_index(drop=True, inplace=True)
        df_rfc = pd.read_excel(filepath_rfc, index_col=0)
        df_rfc.reset_index(drop=True, inplace=True)
        df_data = pd.read_excel(filepath_data, index_col=0)
        df_data.reset_index(drop=True, inplace=True)
        df_master = pd.read_excel(filepath_master, sheet_name=sheet_name, index_col=0)
        df_master.reset_index(drop=True, inplace=True)
        
        # Read fourier coefficients
        columns =['SeriesInstanceUID', 'Count', 'CLASS', 'RFCConfidence', 'RECO_CLASS', 'Features']
        # if os.path.exists(filepath_fourier):
        #     dfFourier = pd.read_pickle(filepath_fourier)
        # else:
        #     dfFourier = pd.DataFrame(columns=columns)
        dfFourier = pd.DataFrame(columns=columns)
        
        if NumSamples is None:
            start = 0
            end = len(df_master)
        else:
            start = NumSamples[0]
            end = NumSamples[1]
        
        #end=100
        for index, row in df_data[start:end].iterrows():
            print('index', index)   
            if keyboard.is_pressed('ctrl+e'):
                print('Button "ctrl + e" pressed to exit execution.')
                sys.exit()
                
            StudyInstanceUID = df_master.loc[index, 'StudyInstanceUID']
            SeriesInstanceUID = df_master.loc[index, 'SeriesInstanceUID']
            RECO_CLASS = df_master.loc[index, 'RECO']
            # RFCClass=row['RFCClass']
            # RFCConfidence=row['RFCConfidence']
            RFCClass = df_master.loc[index, 'RFCClass']
            RFCConfidence = df_master.loc[index, 'RFCConfidence']
            
            if RFCClass=='CACS' and RFCConfidence>0.9:
                #sys.exit()
                try:
                    print('CACS')
                    filepath_image = os.path.join(folderpath_images, StudyInstanceUID, SeriesInstanceUID)
                    imagect = CTImage()
                    imagect.load(filepath_image)
                    if len(imagect.image().shape)==3:                 
                        shape = imagect.image().shape
                        features = self.fourier(imagect) 
                        dfFourier = dfFourier.append({'SeriesInstanceUID': SeriesInstanceUID, 
                                                      'Count': shape[0],
                                                      'CLASS': RFCClass,
                                                      'RFCConfidence':RFCConfidence,
                                                      'RECO_CLASS': RECO_CLASS,
                                                      'Features': features},ignore_index=True)
                        
                    # else:
                    #     features = np.ones((1, self.NumFeatures))*-1
                    #     dfFourier = dfFourier.append({'SeriesInstanceUID': SeriesInstanceUID, 
                    #                                   'Count': 0,
                    #                                   'CLASS': RFCClass,
                    #                                   'RFCConfidence':RFCConfidence,
                    #                                   'RECO_CLASS': RECO_CLASS,
                    #                                   'Features': features},ignore_index=True)
                        
                except:
                    print('Error index', index)
                    features = np.ones((1, self.NumFeatures))*-1
                    dfFourier = dfFourier.append({'SeriesInstanceUID': SeriesInstanceUID, 
                                                  'Count': 0,
                                                  'CLASS': RFCClass,
                                                  'RFCConfidence':RFCConfidence,
                                                  'RECO_CLASS': RECO_CLASS,
                                                  'Features': features},ignore_index=True)
    
                            
            if index % 500 == 0:
                dfFourier.to_pickle(filepath_fourier)
        dfFourier.to_pickle(filepath_fourier)
        
        
    def trainFourier(self, filepath_fourier):   
        
        #Target = 'RECO_CLASS'
        #filepath_fourier = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_master/discharge_master_01042020/discharge_master_01042020_fourier.pkl'
        #filepath_fourier = 'H:/cloud/cloud_data/Projects/CACSFilter/src/reco/discharge_master_01042020_fourier.pkl'
        dfFourier = pd.read_pickle(filepath_fourier)
        #dfFourier = dfFourier[((dfFourier['RECO_CLASS'] == 'IR') | (dfFourier['RECO_CLASS'] == 'FBP'))]
        dfFourier.reset_index(inplace=True, drop=True)
        
        X = np.zeros((dfFourier.shape[0], self.NumFeatures))
        for index, row in dfFourier.iterrows():
            print(row['Features'].shape)
            X[index,:] = row['Features']
        recoClassRF = defaultdict(lambda: -1, {'UNDEFINED':0, 'FBP': 1, 'IR': 2})
        Y = [recoClassRF[x] for x in list(dfFourier['RECO_CLASS'])]
        clfRF, confidence, C, ACC, pred = self.classifier(Y, X,  NumTrees=100)
        print('Confusion matrix:', C)
        print('Confidence:', confidence)
        
        #series = dfFourier['SeriesInstanceUID']
        
        return clfRF
            
    def predictReco(self, settings):
        
        #dfFourier = pd.read_pickle(filepath_fourier)
        clfRF = self.trainFourier(filepath_fourier)
        X = np.ones((len(df_master), self.NumFeatures))*-1
        k=0
        for index, row in df_master.iterrows():
            if keyboard.is_pressed('ctrl+e'):
                print('Button "ctrl + e" pressed to exit execution.')
                sys.exit()
            print('k', k)
            if row['RFCClass'] == 'CACS':
                filepath_image = os.path.join(folderpath_discharge, row['StudyInstanceUID'], row['SeriesInstanceUID'])
                imagect = CTImage()
                try:
                    imagect.load(filepath_image)
                    if len(imagect.image().shape)==3:
                        features = self.fourier(imagect)
                        X[k,:] = features
                except:
                    print('Could not load image:', filepath_image)
            k=k+1
                
        # Predict
        X=np.array(X)
        print('X', X.shape)
        pred = clfRF.predict(X)
        recoClassRFINV = defaultdict(lambda: -1, {0:'UNDEFINED', 1: 'FBP', 2: 'IR'})
        pred = [recoClassRFINV[x] for x in list(pred)]
        pred_df = pd.DataFrame(pred)
        # Replace non CACS reco with 'UNDEFINED'
        ind = df_master['RFCClass']!='CACS'
        pred_df[ind.values] = 'UNDEFINED'
        df_master['RFRECO'] = list(pred_df[0])
        return df_master
            
        
#SeriesInstanceUIDList=['1.2.392.200036.9116.2.6.1.37.2426555318.1447119423.913994']  
#recoList = predictReco(SeriesInstanceUIDList)
    
#############################################################


#extractFourier()
#trainFourier()


#filepath_fourier = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_master/discharge_master_01042020/discharge_master_01042020_fourier.pkl'
#dfFourier = pd.read_pickle(filepath_fourier)