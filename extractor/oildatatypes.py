# -*- coding: utf-8 -*-
"""
Created on Fri Nov 26 12:32:49 2021

@author: Van
"""

import datetime

class DataConvert:
    @staticmethod
    def stringToFloat(strVal):
        if strVal == '':
            return strVal
        
        retVal = float(strVal.translate(str.maketrans('(','-', '),')))
        
        return retVal
    
    @staticmethod
    def stringPercentToFloat(strVal):
        if strVal == '':
            return strVal
        
        retVal = float(strVal.translate(str.maketrans('','', ' %')))
        retVal = retVal / 100
        
        return retVal
    
    @staticmethod
    def datetimeToString(date):
        return "{}/{}/{}".format(date.month,date.day,date.year)

class Check:
    def __init__(self, ID, date, revenue, tax, deductions, total):
        self.ID = ID
        self.date = date
        self.revenue = revenue
        self.tax = tax
        self.deductions = deductions
        self.total = total
        
    def writeToSheet(self, sheet, row_index):
        if row_index == 0:
            print('Attempt to write to check header.')
            return
        
        sheet.write(row_index, 0, float(self.ID))
        sheet.write(row_index, 1, DataConvert.datetimeToString(datetime.datetime.strptime(self.date, '%m/%d/%Y')))
        sheet.write(row_index, 2, DataConvert.stringToFloat(self.revenue))
        sheet.write(row_index, 3, DataConvert.stringToFloat(self.tax))
        sheet.write(row_index, 4, DataConvert.stringToFloat(self.deductions))
        sheet.write(row_index, 5, DataConvert.stringToFloat(self.total))
        
    @staticmethod
    def writeHeaderToSheet(sheet):
        sheet.write(0, 0, 'Check #')
        sheet.write(0, 1, 'Date')
        sheet.write(0, 2, 'Revenue')
        sheet.write(0, 3, 'Tax')
        sheet.write(0, 4, 'Deductions')
        sheet.write(0, 5, 'Total')

class Well:
    def __init__(self, ID, name):
        self.ID = ID
        self.name = name
        self.owner_percent = 0
        
    def setOwnerPercent(self, owner_percent):
        if self.owner_percent == 0:
            self.owner_percent = owner_percent
        elif self.owner_percent != owner_percent:
            print('Different owner percentage for ' + self.name + ': ' + self.owner_percent + ' != ' + owner_percent)
            
    def writeToSheet(self, sheet, row_index):
        if row_index == 0:
            print('Attempt to write to well header.')
            return
        
        sheet.write(row_index, 0, self.ID)
        sheet.write(row_index, 1, self.name)
        sheet.write(row_index, 2, DataConvert.stringPercentToFloat(self.owner_percent)) #0.19758400 %
        
    @staticmethod
    def writeHeaderToSheet(sheet):
        sheet.write(0, 0, 'Code')
        sheet.write(0, 1, 'Name')
        sheet.write(0, 2, 'Owner %')
    
class Statement:
    def __init__(self, ID, check, well):
        self.ID = ID
        self.check = check
        self.well = well
    
class WellData:
    def __init__(self, statement, product_type, int_type, production_date, btu_gravity, well_volume, well_price, well_value, owner_volume, owner_value):
        self.statement = statement
        self.product_type = product_type
        self.int_type = int_type
        self.production_date = production_date
        self.btu_gravity = btu_gravity
        self.well_volume = well_volume
        self.well_price = well_price
        self.well_value = well_value
        self.owner_volume = owner_volume
        self.owner_value = owner_value
        
    def writeToSheet(self, sheet, row_index):
        if row_index == 0:
            print('Attempt to write to well data header.')
            return
        
        sheet.write(row_index, 0, DataConvert.datetimeToString(datetime.datetime.strptime(self.statement.check.date, '%m/%d/%Y')))
        sheet.write(row_index, 1, self.statement.well.name)
        sheet.write(row_index, 2, DataConvert.datetimeToString(datetime.datetime.strptime(self.production_date, '%b %d, %Y')))
        sheet.write(row_index, 3, self.product_type)
        sheet.write(row_index, 4, self.int_type)
        sheet.write(row_index, 5, DataConvert.stringToFloat(self.btu_gravity))
        sheet.write(row_index, 6, DataConvert.stringToFloat(self.well_volume))
        sheet.write(row_index, 7, DataConvert.stringToFloat(self.well_price))
        sheet.write(row_index, 8, DataConvert.stringToFloat(self.well_value))
        sheet.write(row_index, 9, DataConvert.stringToFloat(self.owner_volume))
        sheet.write(row_index, 10, DataConvert.stringToFloat(self.owner_value))
        
    @staticmethod
    def writeHeaderToSheet(sheet):
        sheet.write(0, 0, 'Check Date')
        sheet.write(0, 1, 'Well Name')
        sheet.write(0, 2, 'Production Date')
        sheet.write(0, 3, 'Product Type')
        sheet.write(0, 4, 'Int Type')
        sheet.write(0, 5, 'BTU/Gravity')
        sheet.write(0, 6, 'Well Volume')
        sheet.write(0, 7, 'Well Price')
        sheet.write(0, 8, 'Well Value')
        sheet.write(0, 9, 'Owner Volume')
        sheet.write(0, 10, 'Owner Value')