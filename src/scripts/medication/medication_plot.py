# -*- coding: utf-8 -*-

import sys, os
sys.path.append('H:/cloud/cloud_data/Projects/DL/Code/src')
sys.path.append('H:/cloud/cloud_data/Projects/DL/Code/src/ct')
import pandas as pd
import ntpath
import datetime
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.styles import Font, Color, Border, Side
from openpyxl.styles import Protection
from openpyxl.styles import PatternFill
import plotly.graph_objects as go
import urllib.request
import numpy as np
import plotly.express as px

# # Read data
filepat_data = 'H:/cloud/cloud_data/Projects/DISCHARGEMaster/src/scripts/medication/sanStat.xlsx'
filepath_html = 'H:/cloud/cloud_data/Projects/MixedEffectModels/src/scripts/medication/html/medication_plot.html'
data = pd.read_excel(filepat_data, sheet_name='for graph by R')

# Create labels and links
label_org = list(np.unique((data['sanStat ( left column in the graph)']))) + list(np.unique((data['statins2 ( right column in the graph) '])))
label = [label_org[1],label_org[0]] + label_org[2:]
source = [0,0,1,1,2,2,3,3,4,4,5,5]
target = [6,7,6,7,6,7,6,7,6,7,6,7]

# Update values
value=[]
for s,t in zip(source, target):
    dfs = data[(data['sanStat ( left column in the graph)']==label[s]) & (data['statins2 ( right column in the graph) ']==label[t])]
    value.append(len(dfs))

# Set color for lines and nodes
color_links = ['rgba(0,255,0,0.5)' for i in range(4)] + ['rgba(255,255,0,0.5)' for i in range(4)] + ['rgba(255,0,0,0.5)' for i in range(4)]
color_node = ['rgba(0,255,0,0.5)' for i in range(2)] + ['rgba(255,255,0,0.5)' for i in range(2)] + ['rgba(255,0,0,0.5)' for i in range(2)] + ['rgba(255,0,255,0.5)', 'rgba(128,0,128,0.5)']

# Create figure
fig = go.Figure(data=[go.Sankey(
    node = dict(
      pad = 15,
      thickness = 20,
      line = dict(color = "black", width = 0.5),
      label = label,
      color = color_node,
    ),
    link = dict(
      source = source,
      target = target,
      value = value,
      color = color_links,
  ))])

fig.update_layout(title_text="Statine statistics", font_size=10)
fig.show()


#fig = px.scatter(x=range(10), y=range(10))
fig.write_html(filepath_html)
