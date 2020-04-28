# -*- coding: utf-8 -*-
"""
Created on Mon Aug 27 14:14:07 2018

@author: caotrido
"""

#%% Import libraries
import zipfile
import os
import shutil

import pandas as pd

#%% General configuration (user)

MAPPING_FILE        = 'mapping_table.xlsx'
OLD_TABLEAU_FILE    = '/Tableau Input/Data Discovery.twbx'
DIRECTORY_NEW_DATA  = '/Data new'
NEW_TABLEAU_FILE    = '/Tableau Output/Data Discovery - v2'

current_dir = os.getcwd() # Get current path
os.makedirs(current_dir + '/_tmp') # Create tmp directory
df_mapping = pd.read_excel(MAPPING_FILE) # Get Mapping files

#%% Define functions

def get_all_files(root):
    list_files = []
    for path, subdirs, files in os.walk(root):
        for name in files:
            list_files.append(os.path.join(path, name))
    return list_files

def find_replace_file(df_mapping, list_files, directory_new_files):
    for i in range(0, df_mapping.shape[0]):
        src_file = df_mapping.loc[i, 'old_name']
        destination_file = df_mapping.loc[i, 'new_file']
        response =  [s for s in list_files if src_file in s][0]
        print response
        shutil.copy2(directory_new_files + '/' + destination_file, response) # complete target filename given


#%% decompress tbwx
        
fantasy_zip = zipfile.ZipFile(current_dir + OLD_TABLEAU_FILE)
fantasy_zip.extractall(current_dir + '/_tmp')
fantasy_zip.close()
 
#%% Find and replace data

root = current_dir + '/_tmp'
directory_new_files = current_dir + DIRECTORY_NEW_DATA

# List all files in twbx
list_files = get_all_files(root)
# Replace recursively all filess
find_replace_file(df_mapping, list_files, directory_new_files)
    
#%% Compress the data
shutil.make_archive(current_dir + NEW_TABLEAU_FILE, 'zip', current_dir + '/_tmp')
os.rename(current_dir + NEW_TABLEAU_FILE + '.zip', current_dir + NEW_TABLEAU_FILE + '.twbx')

#%% Delete _tmp directory
shutil.rmtree(current_dir + '/_tmp')
