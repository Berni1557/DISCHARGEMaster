# -*- coding: utf-8 -*-
"""
Created on Thu Apr 16 16:00:00 2020

@author: Bernhard Foellmer

"""

import pandas as pd
pd.options.mode.chained_assignment = None
from inspect import isfunction
import numpy as np
import math
from collections import defaultdict
from sklearn.cluster import KMeans
from scipy.stats import zscore
from sklearn.metrics import confusion_matrix
from sklearn.model_selection import train_test_split
from sklearn.metrics import accuracy_score
from sklearn.ensemble import RandomForestClassifier
from sklearn import tree
from sklearn.tree import export_graphviz
from numpy.random import shuffle
import os
import pickle
from graphviz import Source

def autosize_excel_columns(worksheet, df):
    """ autosize_excel_columns
    """
    
    autosize_excel_columns_df(worksheet, df.index.to_frame())
    autosize_excel_columns_df(worksheet, df, offset=df.index.nlevels)

def autosize_excel_columns_df(worksheet, df, offset=0):
    """ autosize_excel_columns_df
    """
    
    for idx, col in enumerate(df):
        series = df[col]
        max_len = max((  series.astype(str).map(len).max(),      len(str(series.name)) )) + 1        
        max_len = min(max_len, 100)
        worksheet.set_column(idx+offset, idx+offset, max_len)
    
def df_to_excel(writer, sheetname, df):
    """ df_to_excel
    """
    
    df.to_excel(writer, sheet_name = sheetname, freeze_panes = (df.columns.nlevels, df.index.nlevels))
    autosize_excel_columns(writer.sheets[sheetname], df)
    
def format_differences(worksheet, levels, df_ref, df_bool, format_highlighted):
    """ format_differences
    """
    
    for i in range(df_ref.shape[0]):
        for j in range(df_ref.shape[1]):
            if df_bool.iloc[i,j]:
                v  = df_ref.iloc[i,j]
                try:
                    if v!=v:
                        worksheet.write_blank(i+1, j+levels, None, format_highlighted)
                    else:                                        
                        worksheet.write(i+1, j+levels, v, format_highlighted)
                except:
                    print("An exception occurred "+type(v))                    
    return

class DISCHARGEFilter:
    """ Create DISCHARGEFilter
    Filer DISCHARGE column and return boolean pandas series
        
    """
    def __init__(self):
        """ Init DISCHARGEFilter
        """
        self.filtetype = 'FILTER'
        self.feature = ''
        self.filter0 = ''
        self.filter1 = ''
        self.mapFunc = None
        self.featFunc = None
        self.color = ''
        self.df_filt=None
        self.updateTarget=True
        self.operation = 'AND' # 'AND' / 'OR'
        
    def createFilter(self, feature='', minValue=None, maxValue=None, exactValue=[], mapFunc=None, name='', color='', updateTarget=True, featFunc=None, operation ='AND'):
        """ Create normal filter
        
        :param feature: Name of the column which is filtered
        :type feature: str
        :param minValue: Minimum value if filterd by minimal boundary else None
        :type minValue: float
        :param maxValue: Maximum value if filterd by maximum boundary else None
        :type maxValue: float
        :param exactValue: List of values, if element of column is equal to element in the list, True is returnd for element in the column 
        :type exactValue: list
        :param mapFunc: Function which is applied to an element of the column before compared
        :type mapFunc: function
        :param name: Name of the boolean column whcih is added to the output file
        :type name: str
        """
        
        self.filtetype = 'FILTER'
        self.feature = feature
        self.mapFunc = mapFunc
        self.featFunc = featFunc
        self.color = color
        self.updateTarget = updateTarget
        if name:
            self.name = name
        else:
            self.name = feature
        self.operation = operation
                
    def createFilterJoin(self, filter0, filter1, mapFunc, name='FilterJoint', updateTarget=True, featFunc=None, operation ='AND'):
        """ Create joint filter consisting of two normal filter
        
        :param name: Name of the boolean column whcih is added to the output file
        :type name: str
        :param filter0: Filter which is combined with filter1 and operation from type filtetype
        :type filter0: DISCHARGEFilter
        :param filter1: Filter which is combined with filter0 and operation from type filtetype
        :type filter1: DISCHARGEFilter
        :param filtetype: Name of the combination operation of the two filters filter0 and filter1 (e.g. AND, OR)
        :type filtetype: str
        """
        
        self.filtetype = 'JOIN'
        self.filter0 = filter0
        self.filter1 = filter1
        self.name = name
        self.updateTarget = updateTarget
        self.mapFunc = mapFunc
        self.featFunc = featFunc
        self.operation = operation
    
    

    # @staticmethod
    # def StringIndex(s,name='SeriesDescription'):
    #     s = [i.lower() for i in s]
    #     def func(v):
    #         #v = v[name]
    #         posArray = np.zeros((len(s)))
    #         if type(v) == str:
    #             wlist = v.split(' ')
    #             for w in wlist:
    #                 if w.lower() in s:
    #                     pos = s.index(w.lower())
    #                     posArray[pos]=1
    #             return tuple(posArray)
    #         else:
    #             return tuple(posArray)
    #     return func

    @staticmethod
    def StringIndex(wordlist ,name='SeriesDescription'):
        wordlist = [i.lower() for i in wordlist]
        def func(v):
            v = v[name]
            posArray = np.zeros((len(wordlist)))
            if type(v) == str:
                for i,w in enumerate(wordlist):
                    if w in v.lower():
                        posArray[i]=1                    
                return tuple(posArray)
            else:
                return tuple(posArray)
        return func
    
    # words_SeriesDescription = ['AIDR', 'IR', 'ASiR', 'ORG', 'FBP', 'IDOSE','IMR']
    # f=StringIndex(words_SeriesDescription,name='SeriesDescription')
    # q=f('tesdf,AIDR ORG IMRt')
    
    @staticmethod
    def includeString(s):
        def func(v):
            if type(v) == str:
                return s in v.lower()
            else:
                return False
        return func
    
    @staticmethod
    def includeNotString(s):
        def func(v):
            if type(v) == str:
                return not (s.lower() in v.lower())
            else:
                return True
        return func
        
    @staticmethod
    def includeStringList(sList):
        def func(v):
            if type(v) == str:
                for s in sList:
                    if s.lower() in v.lower():
                        return True
                return False
            else:
                return False
        return func

    @staticmethod
    def includeNotStringList(sList):
        def func(v):
            if type(v) == str:
                for s in sList:
                    if (s in v.lower()):
                        return False
                return True
            else:
                return True
        return func
    
    def filter(self, df):
        """ Filter dataframe df
        
        :param df: Pandas dataframe
        :type df: pd.Dataframe
        """

        if self.filtetype=='FILTER':
            df_filt = df[self.feature].apply(self.mapFunc)
            return df_filt
        elif self.filtetype=='JOIN':
            if self.mapFunc=='AND':
                df_filt0 = self.filter0.filter(df)
                df_filt1 = self.filter1.filter(df)
                df_filt = df_filt0 & df_filt1
                return df_filt
            elif self.mapFunc=='OR':
                df_filt0 = self.filter0.filter(df)
                df_filt1 = self.filter1.filter(df)
                df_filt = df_filt0 | df_filt1
                return df_filt
            else:
                raise ValueError('Filtetype: ' + self.filtetype + ' does not exist.')
        else:
            raise ValueError('Filtetype: ' + self.filtetype + ' does not exist.')
        return None
    
    def __str__(self):
        """ print filer (e.g. print(filter))
        """
        out = ['Filter:']
        out.append('Feature: ' + self.feature)
        out.append('Name: ' + self.name)
        return '\n'.join(out)

class ActiveLearner:
    """ Create TagFilter
    Class to filter tags
    """
    #columns=['LABEL', 'CLASS', 'CONFIDENCE', 'COMMENT',  'Count', 'Modality', 'ProtocolName', 'ImageComments', 'SeriesDescription', 'SliceThickness', 'ReconstructionDiameter','Site', 'ContrastBolusAgent','PatientID', 'SeriesInstanceUID']
    columns=['Count', 'Modality', 'ProtocolName', 'ImageComments', 'SeriesDescription', 'SliceThickness', 'ReconstructionDiameter','Site', 'ContrastBolusAgent','PatientID', 'SeriesInstanceUID']
    df_features = None
    loaded = False
    featurenames=[]
    
    def read(self, filepath_discharge, filepath_discharge_filt=None, target=None):
        """ Filter DISCHARGE excel sheet
        
        :param filepath_discharge: Filpath of the input DISCHARGE excel sheet
        :type filepath_discharge: str
        :param filepath_discharge_filt: Filpath of the output DISCHARGE excel sheet with boolean columns from filters
        :type filepath_discharge_filt: str
        :param discharge_filter_list: List of DISCHARGEFilter
        :type discharge_filter_list: list
        :param Target: Column name of the target
        :type Target: str
        """
        self.loaded = True
        if filepath_discharge_filt is None or not os.path.exists(filepath_discharge_filt):
            # Read discharge tags from linear sheet
            print('Reading file', filepath_discharge)
            sheet = 'linear'
            columns=self.columns
            columns_class = ['LABEL', 'CLASS', 'CONFIDENCE', 'COMMENT']
            columns_class = [ x + '_' + target['TARGET'] for x in columns_class]
            df = pd.read_excel(filepath_discharge, sheet)
            df_class = pd.DataFrame(data=-1, index=df.index, columns = columns_class)
            df_linear = pd.concat([df, df_class], axis=1)
            df_order = df_linear[columns + columns_class]
            print('keyslin', df_linear.keys())
            print('keysord', df_order.keys())
        else:
            sheet = 'linear'
            
            df_linear = pd.read_excel(filepath_discharge_filt, sheet)
            columns = list(pd.keys())
            df_order = df_linear[columns]
            self.columns = columns
            
        return df_linear, df_order

    def write(self, df_linear, filepath_discharge_filt):
        """ Filter DISCHARGE excel sheet
        
        :param filepath_discharge: Filpath of the input DISCHARGE excel sheet
        :type filepath_discharge: str
        :param filepath_discharge_filt: Filpath of the output DISCHARGE excel sheet with boolean columns from filters
        :type filepath_discharge_filt: str
        :param discharge_filter_list: List of DISCHARGEFilter
        :type discharge_filter_list: list
        :param Target: Column name of the target
        :type Target: str
        """

        # Create ordered list
        df_ordered = df_linear.copy()
        df_ordered = df_linear.set_index(['PatientID','StudyInstanceUID','SeriesInstanceUID'])
        df_ordered.sort_index(inplace=True)
        
        # Create workbook list    
        print('Create workbook')
        writer = pd.ExcelWriter(filepath_discharge_filt)            
        df_to_excel(writer, "ordered", df_ordered)    
        df_to_excel(writer, "linear", df_linear)    
        workbook  = writer.book
        
        # Highlight Target
        
        cl = defaultdict(lambda: None, {'FILTER': discharge_filter_opt, 'TARGET': 'CLASS', 'FONT_COLOR': 'red'})
        discharge_targets=[cl]
        
        for target in discharge_targets:
            print('Highlight Target:' + target['TARGET'])
            formatColor = workbook.add_format({'font_color': target['FONT_COLOR'], 'bg_color': target['BG_COLOR']})
            print('color', target['COLOR'])
            for i in range(df_ordered[target['TARGET']].shape[0]):
                if df_ordered[target['TARGET']][i]:
                    first_row=i+1
                    last_row=i+1
                    first_col=2
                    last_col=1000
                    writer.sheets['ordered'].conditional_format(first_row, first_col, last_row, last_col,{'type': 'no_blanks','format': formatColor})
                    writer.sheets['ordered'].conditional_format(first_row, first_col, last_row, last_col,{'type': 'blanks','format': formatColor})
        
            # # Highlight features
            # for filt in discharge_filter_list:
            #     if filt.color:
            #         format_red = workbook.add_format({'font_color': filt.color})
            #         for i in range(filt.df_filt.shape[0]):
            #             if filt.df_filt[i]:
            #                 first_row=i+1
            #                 last_row=i+1
            #                 first_col=0
            #                 last_col=1000
            #                 writer.sheets['ordered'].conditional_format(first_row, first_col, last_row, last_col,{'type': 'no_blanks','format': format_red})
            #                 writer.sheets['ordered'].conditional_format(first_row, first_col, last_row, last_col,{'type': 'blanks','format': format_red})
        
        # Add sheet for number of highligted CACS per patient
        # columns = ['PatientID']
        # for target in discharge_targets:
        #     columns.append(target['TARGET'] + '_num')
        # columns.append('Modality_CT_FOUND')
        # columns.append('Confidence_alt01_min')
        
        # df_Patient = pd.DataFrame(columns = columns)
        # patientList = list(df_linear['PatientID'].unique())
        # for p, patient in enumerate(patientList):
        #     df_pat = df_linear[df_linear['PatientID']==patient]
        #     NumTarget=[]
        #     for i, target in enumerate(discharge_targets):
        #         NumTarget.append((df_pat[target['TARGET']]==True).sum())
        #     Modality_CT = 'CT' in list(df_pat['Modality'])
        #     Confidence_alt01 = min(list(df_pat['CACS_alt01_confidence']))
        #     df_Patient.loc[p] = [patient] + NumTarget + [Modality_CT, Confidence_alt01]
        # #df_Patient['Modality_CT'] = df_linear['Modality']
        # df_to_excel(writer, "patients", df_Patient)  
            
        # Write excel sheet
        writer.save()
        
        # Read discharge tags from linear sheet
        # writer = pd.ExcelWriter(filepath_discharge_filt)            
        # df_to_excel(writer, "linear", df_linear)
        # writer.save()
        
    def extractFeatures(self, df_linear, discharge_filter_list):
        
        df = df_linear.copy()
        sheet = 'linear'
        
        # Define features
        df_features = pd.DataFrame()
        self.featurenames=[]
        for filt in discharge_filter_list:
            if filt.featFunc:
                df_f = df[filt.feature]
                df_f_frame = df_f.to_frame()
                df_f = df_f_frame.apply(filt.featFunc,axis=1, result_type="expand")
                #df_f = df_f.rename(filt.feature)
                df_features = pd.concat([df_features, df_f], axis=1)
                
                if len(df_f.shape)==1:
                    NumFeatures = 1
                else:
                    NumFeatures = df_f.shape[1]
                self.featurenames = self.featurenames + [filt.feature + '_' + str(x) for x in range(NumFeatures)]
        
        # Replace True and False with zero and one
        self.df_features = df_features*1
        
    def apply(self, df_linear, df_order, discharge_filter=[], target=None, filt=None):
        """ Filter DISCHARGE excel sheet
        
        :param filepath_discharge: Filpath of the input DISCHARGE excel sheet
        :type filepath_discharge: str
        :param filepath_discharge_filt: Filpath of the output DISCHARGE excel sheet with boolean columns from filters
        :type filepath_discharge_filt: str
        :param discharge_filter_list: List of DISCHARGEFilter
        :type discharge_filter_list: list
        :param Target: Column name of the target
        :type Target: str
        """
        
        # Add columns
        target_label = 'LABEL' + '_' + target['TARGET']
        if not target_label in list(df_order.keys()):
            columns=self.columns
            print('columnsapp', columns)
            columns_class = ['LABEL', 'CLASS', 'CONFIDENCE']
            columns_class = [ x + '_' + target['TARGET'] for x in columns_class]
            columns_comment = ['COMMENT']
            columns_comment = [ x + '_' + target['TARGET'] for x in columns_comment]
            df_class = pd.DataFrame(data=-1, index=df_linear.index, columns = columns_class)
            df_comment = pd.DataFrame(data='-', index=df_linear.index, columns = columns_comment)
            df_linear = pd.concat([df_linear, df_class, df_comment], axis=1)
            df_order = df_linear[columns + columns_class + columns_comment]
            self.columns = columns + columns_class + columns_comment
        
        Target = 'LABEL' + '_' + target['TARGET']
        print('sum', np.sum(np.array(df_order[Target])>-1))
        if np.sum(np.array(df_order[Target])>-1)>0:
        
            # Update df_linear
            columns=self.columns
            #df_linear[columns] = df_order[columns]
    
            confidence, C, ACC, pred_class, df_features = self.confidencePredictor(df_order, discharge_filter, Target = Target, filt=filt)
            df_order['CLASS' + '_' + target['TARGET']] = pred_class
            df_order['CONFIDENCE' + '_' + target['TARGET']] = confidence
            print('Accuracy forest:', ACC)
            print('Confusion matrix forest:', C)
            print('CONFIDENCE sum forest:', confidence.sum())
            
            # Update df_linear
            df_linear[columns] = df_order[columns]

                
            return df_linear, df_order, df_features
        else:
            print('No Labels found.')
            return df_linear, df_order, None
        
    def confidencePredictor(self, df_linear, discharge_filter_list, Target = 'LABEL', filt=None):
        """ Calculate fonfidence score using random forest classifier
        
        :param df_linear: Dataframe of the data
        :type df_linear: pd.Dataframe
        :param discharge_filter_list: List of DISCHARGEFilter
        :type discharge_filter_list: list
        :param Target: Column name of the target
        :type Target: str
        """
        
        df = df_linear.copy()
        df_features = self.df_features
            
        # Read label
        X = np.array(df_features)
        Y = np.array(df[Target])

        # Filt by column
        if filt:
            idx = np.array(df_linear[filt]==1)
            X = X[idx]
            Y = Y[idx]

        # Filter by class label
        X_train=X[Y>-1]
        Y_train=Y[Y>-1]
        
        # Replace nan by -1
        X_train = np.nan_to_num(X_train, nan=-1)
        X = np.nan_to_num(X, nan=-1)

        # Train random forest
        clfRF = RandomForestClassifier(max_depth=10, n_estimators=100)
        clfRF.fit(X_train, Y_train)
        
        # Extract confusion matrix and accuracy
        pred_train = clfRF.predict(X_train)
        C = confusion_matrix(pred_train, Y_train)
        ACC = accuracy_score(pred_train, Y_train)
    
        # Predict confidence
        prop = clfRF.predict_proba(X)
        #print('prop', prop[0:10])
        pred_class = clfRF.predict(X)
        thr = 1/(Y_train.max()+1)
        #print('thr', thr)
        confidence = (np.max(prop, axis=1)-thr)*(1/(1-thr))

        if filt:
            confidence_all = np.ones((df_linear.shape[0]))
            confidence_all[idx] = confidence
            confidence = confidence_all
            
            pred_class_all = np.ones((df_linear.shape[0])) * -1
            pred_class_all[idx] = pred_class
            pred_class = pred_class_all
        
        return confidence, C, ACC, pred_class, df_features
                  
        
    def treePredictor(self, df_linear, discharge_filter_list, Target = 'LABEL', filt=None, filepath_tree='', filepath_tree_png=''):
        """ Calculate fonfidence score using random forest classifier
        
        :param df_linear: Dataframe of the data
        :type df_linear: pd.Dataframe
        :param discharge_filter_list: List of DISCHARGEFilter
        :type discharge_filter_list: list
        :param Target: Column name of the target
        :type Target: str
        """
        
        df = df_linear.copy()
        df_features = self.df_features
        
        #featurenames = [x.feature for x in discharge_filter_list]
        featurenames = [str(x) for x in range(47)]
            
        # Read label
        X = np.array(df_features)
        Y = np.array(df[Target])

        # Filt by column
        if filt:
            idx = np.array(df_linear[filt]==1)
            X = X[idx]
            Y = Y[idx]

        # Filter by class label
        X_train=X[Y>-1]
        Y_train=Y[Y>-1]
        
        # Replace nan by -1
        X_train = np.nan_to_num(X_train, nan=-1)
        X = np.nan_to_num(X, nan=-1)
        clf = tree.DecisionTreeClassifier()
        clf.fit(X_train, Y_train)
        
        # Extract confusion matrix and accuracy
        pred_train = clf.predict(X_train)
        C = confusion_matrix(pred_train, Y_train)
        ACC = accuracy_score(pred_train, Y_train)
    
        # Predict confidence
        pred_class = clf.predict(X)

        if filt:
            pred_class_all = np.ones((df_linear.shape[0])) * -1
            pred_class_all[idx] = pred_class
            pred_class = pred_class_all
            
        print('Confusion Matrix tree', Target, C)
        print('Accuracy tree', Target, ACC)

        graph = Source( tree.export_graphviz(clf, out_file=None, feature_names=self.featurenames, class_names=[Target[6:], 'NO '+Target[6:]]))
        png_bytes = graph.pipe(format='svg')
        with open(filepath_tree_png,'wb') as f:
            f.write(png_bytes)

        return C, ACC, pred_class, df_features   


#########################################################################################

# # Set parameter
# filepath_discharge = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_tags_16042020.xlsx'
# #filepath_discharge = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_tags_28022020.xlsx'
# filepath_discharge_filt = 'H:/cloud/cloud_data/Projects/CACSFilter/data/discharge_CACS_V03.xlsx'  

# # Create ReconstructionDiameter filter
# ReconstructionDiameterFilter = DISCHARGEFilter()
# ReconstructionDiameterFilter.createFilter(feature='ReconstructionDiameter', name='ReconstructionDiameter', featFunc=lambda v : v)
# # Create SliceThickness filter for 3.0 mm
# SliceThicknessFilter = DISCHARGEFilter()
# SliceThicknessFilter.createFilter(feature='SliceThickness', name='SliceThicknessFilter', featFunc=lambda v : v)

# # Create site filter
# SiteFilter = DISCHARGEFilter()
# SiteFilter.createFilter(feature='Site', name='SiteFilter', featFunc=lambda v : float(v['Site'][1:]))
# # Create Modality Filter
# ModalityFilter = DISCHARGEFilter()
# ModalityFilter.createFilter(feature='Modality', name='ModalityFilter', featFunc=lambda v : v == 'CT')
# # Create ProtocolName 
# ProtocolNameFilter = DISCHARGEFilter()
# ProtocolNameFilter.createFilter(feature='ProtocolName', updateTarget=True, name='ProtocolName_Calcium Score', featFunc=DISCHARGEFilter.includeStringList(
#     ['Calcium Score','CaScoring', 'CaScoring','CACS','CaScSeq','SMART SCORE','Calsium score','Calcium Score DISCHARGE','ca score', 'CAVE', 'CALCIUM SCORE']))
# # Create CountFilter 
# CountFilter = DISCHARGEFilter()
# CountFilter.createFilter(feature='Count', name='CountFilter', featFunc=lambda v : v)

# words_SeriesDescription = ['REMOVED','Calcium Score','CaScoring', 'CaScoring','CACS','CaScSeq','SMART SCORE','Calsium score','Calcium Score DISCHARGE','ca score', 'CAVE', 'CALCIUM SCORE']
# SeriesDescriptionFilter = DISCHARGEFilter()
# SeriesDescriptionFilter.createFilter(feature='SeriesDescription', name='SeriesDescription', featFunc=DISCHARGEFilter.StringIndex(words_SeriesDescription, name='SeriesDescription'))

# words_ProtocolName = ['REMOVED', 'Calcium Score','CaScoring', 'CaScoring','CACS','CaScSeq','SMART SCORE','Calsium score','Calcium Score DISCHARGE','ca score', 'CAVE', 'CALCIUM SCORE']
# ProtocolNameFilter = DISCHARGEFilter()
# ProtocolNameFilter.createFilter(feature='ProtocolName', name='ProtocolName', featFunc=DISCHARGEFilter.StringIndex(words_ProtocolName, name='ProtocolName'))

# words_ContrastBolusAgent = ['APPLIED', 'Iodine', 'CE']
# ContrastBolusAgentFilter = DISCHARGEFilter()
# ContrastBolusAgentFilter.createFilter(feature='ContrastBolusAgent', name='ContrastBolusAgent', featFunc=DISCHARGEFilter.StringIndex(words_ContrastBolusAgent, name='ContrastBolusAgent'))

# words_ImageComments = ['REMOVED', 'Calcium Score','CaScoring', 'CaScoring','CACS','CaScSeq','SMART SCORE','Calsium score','Calcium Score DISCHARGE','ca score', 'CAVE', 'CALCIUM SCORE']
# ImageCommentsFilter = DISCHARGEFilter()
# ImageCommentsFilter.createFilter(feature='ImageComments', name='ImageComments', featFunc=DISCHARGEFilter.StringIndex(words_ImageComments, name='ImageComments'))


# featFunc=DISCHARGEFilter.StringIndex(words_SeriesDescription)

# # Append filter
# discharge_filter_opt=[]
# discharge_filter_opt.append(ReconstructionDiameterFilter)
# discharge_filter_opt.append(SliceThicknessFilter)
# discharge_filter_opt.append(SiteFilter)
# discharge_filter_opt.append(ModalityFilter)
# discharge_filter_opt.append(ContrastBolusAgentFilter)
# discharge_filter_opt.append(CountFilter)
# discharge_filter_opt.append(ProtocolNameFilter)
# discharge_filter_opt.append(SeriesDescriptionFilter)
# discharge_filter_opt.append(ImageCommentsFilter)

# # Apply flter
# learner = ActiveLearner()
# self=learner
# #filepath_discharge_filt = None

# # Read data
# df_linear, df_order = learner.read(filepath_discharge, filepath_discharge_filt)

# # Extract features
# learner.extractFeatures(df_linear, discharge_filter_opt)

# # Apply classification
# df_linear, df_order, df_features = learner.apply(df_linear, df_order, discharge_filter_opt)

# # Write data
# #learner.write(df_linear, filepath_discharge_filt)

