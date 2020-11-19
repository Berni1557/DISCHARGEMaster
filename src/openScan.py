import time
import os, sys
import argparse
import subprocess
from glob import glob
from subprocess import Popen

is_windows = sys.platform.startswith('win')
if is_windows:
    print('Open scan in slicer')
    time.sleep(1)
    # Get arguments
    StudyInstanceUID = sys.argv[1]
    SeriesInstanceUID = sys.argv[2]
    print('StudyInstanceUID', StudyInstanceUID)
    print('SeriesInstanceUID', SeriesInstanceUID)
    # Create  filepaths
    filepath3DSlicer = 'H:/ProgramFiles/Slicer 4.10.2/Slicer.exe'
    folderpath = 'G:/discharge'
    folderpath_dcm = os.path.join(folderpath, StudyInstanceUID, SeriesInstanceUID)
    filepath = glob(folderpath_dcm + '/*.dcm')[0]
    if os.path.exists(filepath):
        filepath = filepath.replace("\\", '/')
        # Open slicer
        command_default = """SLICER --python-code "slicer.util.loadVolume('FILEPATH', returnNode=True)"""
        command_default = command_default.replace('SLICER', filepath3DSlicer)
        command = command_default.replace('FILEPATH', filepath)
        process = subprocess.Popen(command)
    else:
        time.sleep(10)
        print('Filepath', filepath, 'not found')
else:

    StudyInstanceUID = sys.argv[1]
    SeriesInstanceUID = sys.argv[2]

    # Create  filepaths
    filepath3DSlicer = '/media/SSD2/cloud_data/Projects/CACSFilter/src/slicer/Slicer-4.10.2-linux-amd64/Slicer'
    folderpath = '/media/bernifoellmer/DISCHARGE_BF/discharge'
    folderpath_dcm = os.path.join(folderpath, StudyInstanceUID, SeriesInstanceUID)
    filepath = glob(folderpath_dcm + '/*.dcm')[0]

    if os.path.exists(filepath):
        file_template = '/media/SSD2/cloud_data/Projects/CACSFilter/src/execute_template.sh'
        file = '/media/SSD2/cloud_data/Projects/CACSFilter/src/execute.sh'
        # Copy file template to file
        p = Popen(['cp','-p','--preserve',file_template, file])
        p.wait()
        # Replace SLICER adn FILEPATH
        fin = open(file, "r")
        data = fin.read()
        data = data.replace('SLICER', filepath3DSlicer)
        data = data.replace('FILEPATH', filepath)
        fin.close()
        fin = open(file, "w")
        fin.write(data)
        fin.close()
        # Execute shell script
        subprocess.call(file)
    else:
        print('Filepath', filepath, 'not found')
        time.sleep(10)
        
