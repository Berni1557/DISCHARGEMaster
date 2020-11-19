import time
import os, sys
import argparse
import subprocess
from glob import glob
import shutil 
from subprocess import Popen
# Get arguments
StudyInstanceUID = sys.argv[1]
SeriesInstanceUID = sys.argv[2]
# Create  filepaths
filepath3DSlicer = '/home/bernifoellmer/Documents/test/Slicer-4.10.2-linux-amd64/Slicer'
folderpath = '/media/bernifoellmer/DISCHARGE_BF/discharge'
folderpath_dcm = os.path.join(folderpath, StudyInstanceUID, SeriesInstanceUID)
filepath = glob(folderpath_dcm + '/*.dcm')[0]
if os.path.exists(filepath):
    file_template = '/home/bernifoellmer/Documents/test/execute_template.sh'
    file = '/home/bernifoellmer/Documents/test/execute.sh'
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
    time.sleep(10)
    print('Filepath', filepath, 'not found')