import time
import os, sys
import argparse
import subprocess
from glob import glob
from collections import defaultdict
from YAML import YAML

is_windows = sys.platform.startswith('win')

def initYMLFile(filepath_yaml):
    if is_windows:
        folderpath_templates = 'H:/cloud/cloud_data/Projects/CACSFilter/src/email/templates'
        filepath_email = 'H:/cloud/cloud_data/Projects/CACSFilter/src/email/email.txt'
    else:
        folderpath_templates = 'H/media/SSD2/cloud_data/Projects/CACSFilter/src/email/templates'
        filepath_email = '/media/SSD2/cloud_data/Projects/CACSFilter/src/email/email.txt'
        
    EmailAddressList = {'P01': 'berlin@charite.de', 
                        'P02': 'xxx@charite.de', 
                        'P03': 'xxx@charite.de', 
                        'P04': 'xxx@charite.de', 
                        'P05': 'xxx@charite.de', 
                        'P06': 'xxx@charite.de', 
                        'P07': 'xxx@charite.de', 
                        'P08': 'xxx@charite.de', 
                        'P09': 'xxx@charite.de', 
                        'P10': 'xxx@charite.de', 
                        'P11': 'xxx@charite.de',
                        'P12': 'xxx@charite.de',
                        'P13': 'xxx@charite.de',
                        'P14': 'xxx@charite.de',
                        'P15': 'xxx@charite.de',
                        'P16': 'xxx@charite.de',
                        'P17': 'xxx@charite.de',
                        'P18': 'xxx@charite.de',
                        'P19': 'xxx@charite.de',
                        'P20': 'xxx@charite.de',
                        'P21': 'xxx@charite.de',
                        'P22': 'xxx@charite.de',
                        'P23': 'xxx@charite.de',
                        'P24': 'xxx@charite.de',
                        'P25': 'xxx@charite.de',
                        'P26': 'xxx@charite.de',
                        'P27': 'xxx@charite.de',
                        'P28': 'xxx@charite.de',
                        'P29': 'xxx@charite.de',
                        'P30': 'xxx@charite.de',
                        'P31': 'xxx@charite.de',
                        'P32': 'xxx@charite.de',
                        'P33': 'xxx@charite.de',
                        'P34': 'xxx@charite.de',
                        'P35': 'xxx@charite.de',
                        'P36': 'xxx@charite.de',
                        '': 'xxx@charite.de'}
    
    ProblemSummaryEmail = {'Date of the ICA images are wrong': 'DateICA_template.txt',
                           'Missing CT Images': 'MissingCT_template.txt', 
                           'Missing ICA Images': 'MissingICA_template.txt'}
    
    props=defaultdict(lambda:None, {'EmailAddressList': EmailAddressList, 'ProblemSummaryEmail': ProblemSummaryEmail,
                                    'folderpath_templates': folderpath_templates, 'filepath_email': filepath_email})
    yaml = YAML()
    yaml.save(props, filepath_yaml)
    
def readYMLFile(filepath_yaml):
     yaml = YAML()
     d = yaml.load(filepath_yaml)
     return d
      
def extractProblems(columns, filepath):
    
    # Extract problem list
    file = open(filepath, 'r') 
    Lines = file.readlines() 
    file.close()

    problem={}
    k=-1
    problemList = []
    for line in Lines:
        k=k+1
        if is_windows:
            line = line.strip()
        else:
            line = line[1:-1]
        if line=='---':
            if not problem=={}:
                problemList.append(problem)
            problem={}
            k=-1
        if k>=0:
            problem[columns[k]] = line
    problemList.append(problem)
    
    return problemList

def replaceEmail(filepath_template, filepath_email, problemList, settings):
    
    print('Creating email')
    try:
        time.sleep(3)
        EmailAddressList = settings['EmailAddressList']
        PROBLEM = problemList[0]['Problem']
        PROBLEMLIST=''
        PATIENTID = problemList[0]['PatientID']
        PATIENTIDLIST=''
        
        for problem in problemList:
            PATIENTIDLIST = PATIENTIDLIST + problem['PatientID'] + ', '
            PROBLEMLIST = PROBLEMLIST + problem['Problem'] + ', '
    
        fin = open(filepath_template, "rt")
        
        print('Open filepath_template', filepath_template)
        data = fin.read()
        data = data.replace('EMAIL', EmailAddressList[problem['Site']])
        print('Replacing PROBLEMLIST', PATIENTIDLIST)
        print('Replacing PATIENTIDLIST', PATIENTIDLIST)
        print('Replacing PATIENTID', PATIENTID)
        print('Replacing PROBLEM', PROBLEM)
        data = data.replace('PATIENTIDLIST', PATIENTIDLIST)
        data = data.replace('PROBLEMLIST', PROBLEMLIST)
        data = data.replace('PATIENTID', PATIENTID)
        data = data.replace('PROBLEM', PROBLEM)
        
        #close the input file
        fin.close()
        fin = open(filepath_email, "wt")
        fin.write(data)
        fin.close()
        print('Email creation finished')
        time.sleep(3)
    except:
        print("------------ Error ------------")
        print('Replacing PROBLEMLIST', PATIENTIDLIST)
        print('Replacing PATIENTIDLIST', PATIENTIDLIST)
        print('Replacing PATIENTID', PATIENTID)
        print('Replacing PROBLEM', PROBLEM)
        time.sleep(20)
    
##################################################################

print('Create Email')

if is_windows:
    filepath_yaml = 'H:/cloud/cloud_data/Projects/CACSFilter/src/settings.yml'
else:
    filepath_yaml = '/media/SSD2/cloud_data/Projects/CACSFilter/src/settings.yml'
    

# Init settings file
initYMLFile(filepath_yaml)

# Read settings file
settings = readYMLFile(filepath_yaml)

# Get arguments
#filepath_problems = sys.argv[1]
filepath_problems = '/media/SSD2/cloud_data/Projects/CACSFilter/src/text_tmp.txt'
#filepath_template = 'H:/cloud/cloud_data/Projects/CACSFilter/src/EmailTemplate.txt'
#filepath_email = 'H:/cloud/cloud_data/Projects/CACSFilter/src/email/email.txt'


columns=['Row', 'ProblemID', 'Site', 'PatientID', 'StudyInstanceUID', 'SeriesInstanceUID',
         'ProblemSummary', 'Problem', 'Date of Query', 'Date of the change/sending', 'Results',
         'Status', 'Responsible Person']

# Extract problems
problemList = extractProblems(columns, filepath_problems)
# problemList=[{'Site': 'P01', 'PatientID': 'pat01', 'Problem': 'problem01', 'ProblemSummary': 'Date of the ICA images are wrong'}, 
#              {'Site': 'P01', 'PatientID': 'pat02', 'Problem': 'problem02', 'ProblemSummary': 'Date of the ICA images are wrong'}]

print('problemList', problemList)


# Replace problem in email
filepath_template = os.path.join(settings['folderpath_templates'], settings['ProblemSummaryEmail'][problemList[0]['ProblemSummary']])
filepath_email = settings['filepath_email']

print('replaceEmail')

replaceEmail(filepath_template, filepath_email, problemList, settings)

