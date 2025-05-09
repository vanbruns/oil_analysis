# -*- coding: utf-8 -*-
"""
Created on Fri Nov 26 11:41:44 2021

@author: Van
"""

import datetime
import re
import os

from xlwt import Workbook
from bs4 import BeautifulSoup as soup

from oildatatypes import Check
from oildatatypes import Well
from oildatatypes import WellData
from oildatatypes import Statement

# folder for saved web pages containing check data
CHECK_DATA_PATH = '<insert full path>'

checks = { }
wells = { }

# Create workbook for saving the data
wb = Workbook()

# Create sheets
check_sheet = wb.add_sheet('Checks')
well_sheet = wb.add_sheet('Wells')
well_data_sheet = wb.add_sheet('Well Data')

# Write headers to sheets
Check.writeHeaderToSheet(check_sheet)
Well.writeHeaderToSheet(well_sheet)
WellData.writeHeaderToSheet(well_data_sheet)

# Initialize well data row index (header first)
well_data_row_index = 1

# Initialize html text to empty
html_text = ''

# Get the check html files (should be the ones without 'Revenue Statement')
check_files = (f for f in os.listdir(CHECK_DATA_PATH) if (f.lower().endswith('.html') or f.lower().endswith('.htm')) and 'Revenue Statement' not in f)

for check_file in check_files:
    # Read html
    with open(CHECK_DATA_PATH + check_file, 'r') as f: html_text = f.read()
    check_soup = soup(html_text, 'html.parser')
    
    # Get check number and date from title
    check_title = check_soup.title.string
    check_title = check_title[check_title.find('Check ')+6:]
    check_values = check_title.split(' - ')
    check_number = check_values[0]
    check_date = datetime.datetime.strptime(check_values[1].strip(),'%b %d, %Y').strftime('%m/%e/%Y')
    
    if check_number in checks.keys():
        print('Attempting to add a check that already exists.')
        break
    
    # Get check revenue (3), tax (5), deductions (7), and total (9)
    grid_totals = check_soup.find('table', { 'id':'gridTotals' })
    grid_totals_tds = grid_totals.find_all('td')
    check_revenue = grid_totals_tds[5].text
    check_tax = grid_totals_tds[7].text
    check_deductions = grid_totals_tds[9].text
    check_total = grid_totals_tds[11].text
    
    # Add this check to checks list
    checks[check_number] = Check(check_number, check_date, check_revenue, check_tax, check_deductions, check_total)
    
    # Get the prefix for statement files associated with this check
    statement_prefix = ' - '.join(check_file.split(' - ', 2)[:2])
    statement_prefix += ' - Revenue Statement'
    
    # Get the statement html files
    statement_files = (f for f in os.listdir(CHECK_DATA_PATH) if (f.lower().endswith('.html') or f.lower().endswith('.htm')) and f.startswith(statement_prefix))
    
    print('On check: ' + check_date)
    
    for statement_file in statement_files:
        # Read html
        with open(CHECK_DATA_PATH + statement_file, 'r') as f: html_text = f.read()
        statement_soup = soup(html_text, 'html.parser')
        
        # Get well ID and name
        well_tds = statement_soup.find('span', { 'id':'cphMain_upPropertyInfo' }).find_all('td')
        well_id = well_tds[4].a.text
        well_name = well_tds[5].text.strip().removesuffix(' TX DEWITT')
        
        # Only add new well if it doesn't exist
        if well_id not in wells.keys():
            wells[well_id] = Well(well_id, well_name)
            
        statement_id = statement_file.split(' - ')[4].split('.')[0]
            
        # Create statement object to use
        statement = Statement(statement_id, checks[check_number], wells[well_id])
        
        print('  On statement: ' + statement_id)
        
        # Get statement table
        statement_table_trs = statement_soup.find_all('tr', { 'id':re.compile('cphGrid_gridDetailsRow.*') })

        for statement_tr in statement_table_trs:
            tr_tds = statement_tr.find_all('td')
            tds_0_split = tr_tds[0].a.text.split('.')
            
            well_data = WellData(statement,
                                 tds_0_split[0].strip(), # Product Type: OIL, GAS, NGL
                                 tds_0_split[1].strip(), # Int Type: NR, TRANS, TAX, PROC
                                 tr_tds[2].text.strip().replace(' ', ' 01, 20', 1), # Production Date
                                 tr_tds[3].text.strip(), # BTU/Gravity
                                 tr_tds[4].text.strip(), # Property Volume
                                 tr_tds[5].text.strip(), # Property Price
                                 tr_tds[6].text.strip(), # Property Value
                                 tr_tds[9].text.strip(), # Owner Volume
                                 tr_tds[10].text.strip()) # Owner Value
            
            # Set the well owner percentage
            well_data.statement.well.setOwnerPercent(tr_tds[7].text.strip())
            
            # If the percentages are different, we've made a bad assumption
            if statement.well.owner_percent != tr_tds[8].text.strip():
                print('Percent different for statement: ' + statement.ID)
                
            # Write well data to sheet and increment row
            well_data.writeToSheet(well_data_sheet, well_data_row_index)
            well_data_row_index += 1
            
# Write checks to check sheet
for row, check in enumerate(checks.values()):
    check.writeToSheet(check_sheet, row + 1)
    
# Write wells to well sheet
for row, well in enumerate(wells.values()):
    well.writeToSheet(well_sheet, row + 1)
    
# Save workbook to file
wb.save(CHECK_DATA_PATH + 'results.xls')