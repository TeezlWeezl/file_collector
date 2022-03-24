import pandas as pd
import numpy as np
import tkinter.filedialog as tk
import os
import shutil

inv_found = False
counter_found = 0
counter_missing = 0

def enter_format(list_input, io):
    io_new = ''

    i = 0
    while i < len(io):
        start = 0
        end = 0
        if io[i] == '{':
            start = i
            for j, value in enumerate(io[start+1:]):
                if value == '}':
                    end = j+start+1
                    break
            i = end+1
            io_new += list_input[int(io[start+1:end])]
        if i < len(io):
            io_new += io[i]     
        else: 
            break 
        i+=1
    return io_new

def find_and_copy(files, row):
    global counter_found
    for inv in files:
        if str(row[int(search_term)]) in inv:
            shutil.copyfile(f'{root}/{inv}', f'{save_dir}/export/{enter_format(xls_data[index_row], io)}.pdf')
            counter_found += 1
            return True

xls_data = pd.read_excel('./testing/xls_basis/xls_data.xlsx')
keys_list = xls_data.keys().to_list()

for i, key in enumerate(keys_list):
    print(f'[{i}]\t{key}')

search_term = input('Please enter the number of the column that yields the search term:\n>> ')
search_dir = tk.askdirectory(title='Open the directory to search through:')
save_dir = tk.askdirectory(title='Directory to save the files:')
io = input('Please enter the format you want to rename. Enter variable parts by {<column number>}:\n>> ')
#search_dir = './testing/inv_docs'
#save_dir = './testing/result'

try:
    os.makedirs(f'{save_dir}/export')
except(FileExistsError):
    pass

xls_data = xls_data.values.tolist()
for index_row, row in enumerate(xls_data):

    for index_col, col in enumerate(row):
        if type(col) == pd._libs.tslibs.timestamps.Timestamp:
            xls_data[index_row][index_col] = f'{col.year}-{col.month:02d}-{col.day:02d}'
        elif type(col) == float:
            xls_data[index_row][index_col] = '{:.2f}'.format(col).replace('.',',')
        elif type(col) == int:
            xls_data[index_row][index_col] = str(col)

    for root, dirs, files in os.walk(search_dir):
        inv_found = find_and_copy(files, row)
        if inv_found:
            break
    
    if not(inv_found):
        counter_missing += 1
        with open(f'{save_dir}/export/log_not_found.txt', 'a') as f:
            writer = ''
            for inv_attr in xls_data[index_row]:
                writer += inv_attr + '\t'
            f.write(writer + '\n')

print('Found files: ', counter_found)
print('Missing files: ', counter_missing)