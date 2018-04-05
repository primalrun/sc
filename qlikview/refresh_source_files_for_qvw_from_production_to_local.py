import os
import sys
import csv
import shutil
import tkinter as tk
from tkinter import filedialog
 
#Requires c:\qv_dev\qvd_list.txt to be updated with qvd files for *dashboard.qvw
local_dir = 'c:/qv_dev/'
qvd_dir = local_dir + 'qvd'
s_file = local_dir + r'qvd_list.txt'
prod_dir = '//qlikview/c$/qlikview/'
qvd_prod_dir = prod_dir + 'qvd'

#user select qvw file
root = tk.Tk()
root.withdraw()
qvw_file = filedialog.askopenfilename()

if len(qvw_file) == 0:
    print('No qvw file selected, try again')
    sys.exit()
qvw_split = qvw_file.split('/')
cat_folder = qvw_split[-2]
qvw_file = qvw_split[-1]

#import file exist?
s_file_exists = os.path.isfile(s_file)
if s_file_exists == False:
    print(r'Unable to find qvd_list.txt on ' + local_dir)
    sys.exit()

#qvd directory
if os.path.exists(qvd_dir):
    shutil.rmtree(qvd_dir)

os.mkdir(qvd_dir)

#sourcedocuments directory
if os.path.exists(local_dir + '/' + 'sourcedocuments'):
    shutil.rmtree(local_dir + '/' + 'sourcedocuments')
os.mkdir(local_dir + '/' + 'sourcedocuments')

#sourcedata directory
if os.path.exists(local_dir + '/' + 'sourcedata'):
    shutil.rmtree(local_dir + '/' + 'sourcedata')
os.mkdir(local_dir + '/' + 'sourcedata')

#get qvd file names
with open(s_file, 'r') as f:
    reader = csv.reader(f)
    qvd_dict = {r[0]:None for r in reader if len(r) > 0}
    
#populate list with directory, folder name, qvd file
file_list = []
dir_list = os.listdir(qvd_prod_dir)
for v in qvd_dict:
    for f in dir_list:        
        if os.path.isfile(qvd_prod_dir + '/' + f + '/' + v):
            file_list.append((qvd_prod_dir, f, v))
            break

#create folders on local pc
for f in file_list:    
    if not os.path.exists(qvd_dir + '/' + f[1]):
        os.mkdir(qvd_dir + '/' + f[1])

#copy qvd files from production to local pc
for f in file_list:
    src_file = f[0] + '/' + f[1] + '/' +f[2]
    dest_file = qvd_dir + '/' + f[1] + '/' +f[2]
    shutil.copyfile(src=src_file,
                     dst=dest_file)
    print('Copying {} to {}'.format(src_file, dest_file))

#copy other general purpose source files
#qvd_history
if not os.path.exists(qvd_dir + '/' + 'Administrative'):
    os.mkdir(qvd_dir + '/' + 'Administrative')
src_file = qvd_prod_dir + '/' + 'Administrative' + '/' + 'qvd_history.qvd'
dest_file = qvd_dir + '/' + 'Administrative' + '/' + 'qvd_history.qvd'    
shutil.copyfile(src=src_file,
                 dst=dest_file)
print('Copying {} to {}'.format(src_file, dest_file))

#sourcedocuments files
os.mkdir(local_dir + '/' + 'sourcedocuments' + '/' +'_includes')
src_file = prod_dir + 'sourcedocuments' + '/' + '_includes' + '/' + 'qvdmaker.txt'
dest_file = local_dir + '/' + 'sourcedocuments' + '/' + '_includes' + '/' + 'qvdmaker.txt'    
shutil.copyfile(src=src_file,
                 dst=dest_file)
print('Copying {} to {}'.format(src_file, dest_file))

src_file = prod_dir + 'sourcedocuments' + '/' + '_includes' + '/' + 'qvdreloadtimes.txt'
dest_file = local_dir + '/' + 'sourcedocuments' + '/' + '_includes' + '/' + 'qvdreloadtimes.txt'    
shutil.copyfile(src=src_file,
                 dst=dest_file)
print('Copying {} to {}'.format(src_file, dest_file))

src_file = prod_dir + 'sourcedocuments' + '/' + '_includes' + '/' + 'qvdreloadtimessense.txt'
dest_file = local_dir + '/' + 'sourcedocuments' + '/' + '_includes' + '/' + 'qvdreloadtimessense.txt'    
shutil.copyfile(src=src_file,
                 dst=dest_file)
print('Copying {} to {}'.format(src_file, dest_file))

src_file = prod_dir + 'sourcedocuments' + '/' + '_includes' + '/' + 'sc_color_palettes.qvs'
dest_file = local_dir + '/' + 'sourcedocuments' + '/' + '_includes' + '/' + 'sc_color_palettes.qvs'    
shutil.copyfile(src=src_file,
                 dst=dest_file)
print('Copying {} to {}'.format(src_file, dest_file))

src_file = prod_dir + 'sourcedocuments' + '/' + '_includes' + '/' + 'sc_master_calendar.qvs'
dest_file = local_dir + '/' + 'sourcedocuments' + '/' + '_includes' + '/' + 'sc_master_calendar.qvs'    
shutil.copyfile(src=src_file,
                 dst=dest_file)
print('Copying {} to {}'.format(src_file, dest_file))

#sourcedata files    
src_file = prod_dir + 'sourcedata' + '/' + 'qlik content and security.xlsx'
dest_file = local_dir + '/' + 'sourcedata' + '/' + 'qlik content and security.xlsx'    
shutil.copyfile(src=src_file,
                 dst=dest_file)
print('Copying {} to {}'.format(src_file, dest_file))

#qvw file selected by user
os.mkdir(local_dir + '/' + 'sourcedocuments' + '/' + cat_folder)
src_file = prod_dir + 'sourcedocuments' + '/' + cat_folder + '/' + qvw_file
dest_file = local_dir +  'sourcedocuments' + '/' + cat_folder + '/' + qvw_file    
shutil.copyfile(src=src_file,
                 dst=dest_file)
print('Copying {} to {}'.format(src_file, dest_file))

    
print('Success')
