# -*- coding: utf-8 -*-

import sys
sys.path.append('H:/cloud/cloud_data/Projects/DL/Code/src')
from reco.reco_filter import RecoFilter
from ct.CTDataStruct import CTPatient
from settings import initSettings, saveSettings, loadSettings, fillSettingsTags
import numpy as np
import SimpleITK as sitk

def sgnalToNoise(image):
    dx = 30
    dy = 30
    s0 = int(np.round(image.shape[0]*0.2))
    s1 = int(np.round(image.shape[0]*0.7))
    for s in range(s0,s1):
        print(s)
        for x in range(0, 512, 10):
            print('x', x)
            for y in range(0, 512, 10):
                print('y', y)
                roi = image[s,x:x+dx,y:y+dy]
                print('min', roi.min())
                print('max', roi.max())
                if roi.min()>0 and roi.max()<130:
                    mean=np.mean(r1)  
                    std=np.std(r1)
                    sn1=mean/std
                    return (sn1, s, x, y)


def stat(v):
    a = sitk.GetArrayFromImage(v).transpose(2,1,0)
    [sx,sy,sz] = a.shape
    #print(sx,sy,sz)
    b = a[:,:,int(sz/2)]   
    #print(b.shape)
    r = 100
    cx = int(sx/2)
    cy = int(sy/2)
    c = b[cx-r:cx+r,cy-r:cy+r]
    #print(c.shape)
    mean = np.mean(c)
    std = np.std(c)   
    #print(mean,std)
    return mean, std

def gradientmagitude(image):
    gmif = sitk.GradientMagnitudeImageFilter()
    gmif.SetUseImageSpacing(True)
    im = gmif.Execute(image)
    return im  

def gradientmagitude(image):
    
    
    gmif = sitk.GradientMagnitudeImageFilter()
    gmif.SetUseImageSpacing(True)
    image_std = gmif.Execute(image)
    
    image_np = sitk.GetArrayFromImage(image)
    image_std_np = sitk.GetArrayFromImage(image_std)
    
    image_np=image_np[33:36]
    image_std_np=image_std_np[33:36]
    
    #plt.imshow(image_np[34,:,:])
    #plt.show()
    
    mask0 = image_std_np<90
    mask1 = (image_np>30) & (image_np<110)
    mask = mask0 * mask1
    
    image_std_np_mask = image_std_np * mask
    values = image_std_np_mask[image_std_np_mask>0]
    std = np.std(values)
    
    return std, image_std_np_mask
                
# Load settings
filepath_settings = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/data/settings.json'
settings=initSettings()
saveSettings(settings, filepath_settings)
settings = fillSettingsTags(loadSettings(filepath_settings))

# FBP reconstruction
PatientID = '01-BER-0012'
StudyInstanceUID = '1.2.840.113619.6.95.31.0.3.4.1.1018.13.10329788'
SeriesInstanceUID = '1.2.392.200036.9116.2.6.1.37.2426555318.1461798161.464800'
patient = CTPatient(StudyInstanceUID, PatientID)
series2 = patient.loadSeries(settings['folderpath_discharge'], SeriesInstanceUID, None)
image0 = series2.image
#im = gradientmagitude(image0.image_sitk)
#mean, std = stat(im)
std,_ = gradientmagitude(image0.image_sitk)
print(std)
# s0 = image0.image()[20,:,:]
# r0=s0[190:280,280:380]
# m0=np.mean(r0)  
# s0=np.std(r0)
# # sn0=m0/s0
# sn0 = sgnalToNoise(series2.image.image())
# print('sn0', sn0)


# IR reconstruction
PatientID = '01-BER-0012'
StudyInstanceUID = '1.2.840.113619.6.95.31.0.3.4.1.1018.13.10329788'
SeriesInstanceUID = '1.2.392.200036.9116.2.6.1.37.2426555318.1461798187.314816'
patient = CTPatient(StudyInstanceUID, PatientID)
series2 = patient.loadSeries(settings['folderpath_discharge'], SeriesInstanceUID, None)
image0 = series2.image
std,_ = gradientmagitude(image0.image_sitk)
print(std)
# image0 = series2.image
# s0 = image0.image()[20,:,:]
# r0=s0[190:280,280:380]
# m0=np.mean(r0)  
# s0=np.std(r0)
# sn0=m0/s0
# sn0 = sgnalToNoise(series2.image.image())
# print('sn0', sn0)

# FBP reconstruction
PatientID = '01-BER-0014'
StudyInstanceUID = '1.2.840.113619.6.95.31.0.3.4.1.1018.13.10347678'
SeriesInstanceUID = '1.2.392.200036.9116.2.6.1.37.2426555318.1462922209.252410'
patient = CTPatient(StudyInstanceUID, PatientID)
series1 = patient.loadSeries(settings['folderpath_discharge'], SeriesInstanceUID, None)
image0 = series1.image
std,_ = gradientmagitude(image0.image_sitk)
print(std)
# image1 = series1.image
# s1 = image1.image()[23,:,:]
# r1=s1[190:280,250:370]
# m1=np.mean(r1)  
# s1=np.std(r1)
# sn1=m1/s1
# sn1 = sgnalToNoise(series1.image.image())
# print('sn1', sn1)

# IR reconstruction
PatientID = '01-BER-0014'
StudyInstanceUID = '1.2.840.113619.6.95.31.0.3.4.1.1018.13.10347678'
SeriesInstanceUID = '1.2.392.200036.9116.2.6.1.37.2426555318.1462922234.743744'
patient = CTPatient(StudyInstanceUID, PatientID)
series1 = patient.loadSeries(settings['folderpath_discharge'], SeriesInstanceUID, None)
image0 = series1.image
std,_ = gradientmagitude(image0.image_sitk)
print(std)

# s1 = image1.image()[23,:,:]
# r1=s1[190:280,250:370]
# m1=np.mean(r1)  
# s1=np.std(r1)
# sn1=m1/s1
# print('sn1', sn1)










