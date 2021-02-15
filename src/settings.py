# -*- coding: utf-8 -*-
import os, sys
import json
from collections import defaultdict

def initSettings():
    settings = defaultdict(lambda:None, {})
    settings['folderpath_master'] = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/data/discharge_master'
    settings['date'] = '01092020'
    settings['folderpath_discharge'] = 'G:/discharge'
    settings['dicom_tags'] = ['Site', 'PatientID', 'StudyInstanceUID', 'SeriesInstanceUID', 'AcquisitionDate',
                              'SeriesNumber', 'Count', 'SeriesDescription', 'Modality', 'AcquisitionTime', 'NumberOfFrames',
                              'Rows', 'Columns', 'InstanceNumber', 'PatientSex', 'PatientAge', 'ProtocolName',
                              'ContrastBolusAgent', 'ImageComments', 'PixelSpacing', 'SliceThickness', 'FilterType',
                              'ConvolutionKernel', 'ReconstructionDiameter', 'RequestedProcedureDescription',
                              'ContrastBolusStartTime', 'NominalPercentageOfCardiacPhase', 'CardiacRRIntervalSpecified', 'StudyDate',
                              'SliceSpacing']
    
    settings['dicom_tags_order'] = ['Site','PatientID','StudyInstanceUID','SeriesInstanceUID','AcquisitionDate','SeriesNumber', 'Count', 'NumberOfFrames', 'SeriesDescription',
                             'Modality','Rows', 'InstanceNumber','ProtocolName','ContrastBolusAgent','ImageComments','PixelSpacing','SliceThickness','ConvolutionKernel',
                             'ReconstructionDiameter','RequestedProcedureDescription','ContrastBolusStartTime','NominalPercentageOfCardiacPhase','CardiacRRIntervalSpecified',
                             'StudyDate']
    settings['dicom_tags_first'] = ['Site', 'PatientID', 'StudyInstanceUID', 'SeriesInstanceUID', 
                                   'AcquisitionDate', 'SeriesNumber', 'Count', 'SeriesDescription']
    
    settings['columns_first'] = ['Site', 'PatientID', 'StudyInstanceUID', 'SeriesInstanceUID', 'CLASS', 'RFCClass', 'RFCLabel', 
                                 'RFCConfidence','ClassManualCorrection', 'Comment', 'Responsible Person', 
                                 'Image Quality', 'Count', 'SeriesDescription', 'SeriesNumber', 'AcquisitionDate']
    
    settings['recoClasses'] = ['FBP', 'IR', 'UNDEFINED']
    settings['columns_tracking'] = ['ProblemID', 'Site', 'PatientID', 'StudyInstanceUID', 'SeriesInstanceUID', 'Problem Summary',
                                    'Problem', 'Date of Query', 'Date of the change/sending', 
                                    'Results', 'Answer from the site', 'Status', 'Responsible Person']
    settings['centers'] = ['P01', 'P02', 'P04', 'P05', 'P06', 'P07', 'P08',
                           'P09', 'P10', 'P11', 'P12', 'P13', 'P14', 'P15', 'P16',
                           'P17', 'P18', 'P19', 'P22', 'P23', 'P29', 'P30', 'P31', 'P32', 'P33', 'P36']  
    return settings


    
def saveSettings(settings, seetingsfile='seetings.json'):
    with open(seetingsfile, 'w') as f:
        json.dump(settings, f)
    
def loadSettings(seetingsfile='seetings.json'):
    with open(seetingsfile, 'r') as fp:
        settings = json.load(fp)
    return settings

def fillSettingsTags(settings):
    
    settings['folderpath_master_date'] = os.path.join(settings['folderpath_master'], 'discharge_master_' + settings['date'])
    settings['folderpath_sources'] = os.path.join(settings['folderpath_master_date'], 'discharge_sources_' + settings['date'])
    settings['folderpath_components'] = os.path.join(settings['folderpath_master_date'], 'discharge_components_' + settings['date'])
    settings['filepath_dicom'] = os.path.join(settings['folderpath_sources'], 'discharge_dicom_' + settings['date'] + '.xlsx')
    settings['filepath_ITT'] = os.path.join(settings['folderpath_sources'], 'discharge_ITT_' + settings['date'] + '.xlsx')
    settings['filepath_ecrf'] = os.path.join(settings['folderpath_sources'], 'discharge_ecrf_' + settings['date'] + '.xlsx')
    settings['filepath_prct'] = os.path.join(settings['folderpath_sources'], 'discharge_prct_' + settings['date'] + '.xlsx')
    settings['filepath_phase_exclude_stenosis'] = os.path.join(settings['folderpath_sources'], 'discharge_phase_exclude_stenosis_' + settings['date'] + '.xlsx')
    settings['filepath_stenosis_bigger_20_phases'] = os.path.join(settings['folderpath_sources'], 'discharge_stenosis_bigger_20_phases_' + settings['date'] + '.xlsx')
    settings['filepath_tracking'] = os.path.join(settings['folderpath_sources'], 'discharge_tracking_' + settings['date'] + '.xlsx')
    settings['filepath_master_track'] = os.path.join(settings['folderpath_components'], 'discharge_master_track_' + settings['date'] + '.xlsx')
    settings['filepath_data'] = os.path.join(settings['folderpath_components'], 'discharge_data_' + settings['date'] + '.pkl')
    settings['filepath_rfc'] = os.path.join(settings['folderpath_components'], 'discharge_rcf_' + settings['date'] + '.pkl')
    settings['filepath_manual'] = os.path.join(settings['folderpath_components'], 'discharge_manual_' + settings['date'] + '.pkl')
    settings['filepath_prediction'] = os.path.join(settings['folderpath_components'], 'discharge_prediction_' + settings['date'] + '.pkl')
    settings['filepath_master'] = os.path.join(settings['folderpath_master_date'], 'discharge_master_' + settings['date'] + '.xlsx')
    settings['filepath_hist'] = os.path.join(settings['folderpath_sources'], 'discharge_hist_' + settings['date'] + '.pkl')
    settings['filepath_patient'] = os.path.join(settings['folderpath_components'], 'discharge_patient_' + settings['date'] + '.pkl')
    settings['folderpath_manual_selection'] = os.path.join(settings['folderpath_master_date'], 'discharge_manual_selection_' + settings['date'])
    settings['filepath_patient_conf'] = os.path.join(settings['folderpath_components'], 'discharge_patient_conf_' + settings['date'] + '.xlsx')
    settings['filepath_reco'] = os.path.join(settings['folderpath_components'], 'discharge_reco_' + settings['date'] + '.pkl')
    
    os.makedirs(settings['folderpath_master_date'], exist_ok=True)
    os.makedirs(settings['folderpath_sources'], exist_ok=True)
    os.makedirs(settings['folderpath_components'], exist_ok=True)
    
    return settings

filepath_settings = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/data/settings.json'
settings=initSettings()
saveSettings(settings, filepath_settings)
settings = fillSettingsTags(loadSettings(filepath_settings))
