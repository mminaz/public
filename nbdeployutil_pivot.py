# -*- coding: utf-8 -*-
"""
Created on Sat Feb 15 02:20:14 2020

@author: Mike Morgan
"""

# Import modules with shortcut names
import pandas as pd
import numpy as np

# Read in NBDeployUtils
pas_nb = pd.read_excel('report-capacity-pasnetbackup1-20200212_100136.xls', sheet_name='Itemization')
sc_nb = pd.read_excel('report-capacity-scnetbackup01-20200212_102722.xls', sheet_name='Itemization')
sp_nb = pd.read_excel('report-capacity-spwmsdp1-20200212_164030.xls', sheet_name='Itemization')

# Insert Site columns to index 0
pas_nb.insert(0, 'Site', 'Pasadena')
sc_nb.insert(0, 'Site', 'Santa Clara')
sp_nb.insert(0, 'Site', 'São Paulo')

# Create combined list of sites in preferred order
all_nb_list = [pas_nb, sc_nb, sp_nb]

# Concat list into appeded concat
append_nb = pd.concat(all_nb_list, sort=False)

# Create a new columns for TB
append_nb['Charged Size (TB)'] = append_nb['Charged Size (KB)'] / 1024 / 1024 / 1024
append_nb['Tape (TB)'] = append_nb['Tape (KB)'] / 1024 / 1024 / 1024

# Create pivot table
pivot_nb = pd.pivot_table(append_nb, index=['Site', 'Policy Type'], values=['Charged Size (TB)', 'Tape (TB)'], aggfunc=np.sum, fill_value=0)

# Round whole table to 2 decimals
pivot_nb = pivot_nb.round(decimals=2)
print(pivot_nb)

# Write to an excel file
pivot_nb.to_excel('WAM - NBDeployUtil All Sites 20200212.xlsx')
