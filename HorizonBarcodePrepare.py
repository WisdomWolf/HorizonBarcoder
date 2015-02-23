#!/usr/bin/python

import os
import sys
import re
import codecs
import pdb
from datetime import date
from BarcodeItem import BarcodeItem, calculateBarcodeChecksum
from tkinter import *
from tkinter import ttk, messagebox, filedialog
from xlwt import Workbook, easyxf
from xlrd import open_workbook

barcodeListSet = set()
newItemList = []
catChoice = None
priChoice = None

def openFile(options=None):
    if options == None:
        options = {}
        options['defaultextension'] = '.xls' 
        options['filetypes'] = [('Excel Files', '.xls')]
        options['title'] = 'Open...'
    file_opt = options
    file_path = filedialog.askopenfilename(**file_opt)
    options = None
    if file_path == None or file_path == "":
        print("No file selected")
        return
    else:
        readBarcodeRequest(file_path)
        
def oF():
    openFile()

def importBarcodeDatabase():
    file = 'barcodeList.txt'
    if os.path.exists(file):
        with codecs.open(file, 'r+', 'utf-8') as f:
            for barcode in f:
                barcode = str(barcode).strip('\r\n')
                if barcode not in barcodeListSet:
                    barcodeListSet.add(barcode)
        print('imported barcodes from ' + file)
    else:
        updateBarcodeDatabase()
    
def readBarcodeRequest(file=None):
    file = file or 'Copy of Add Menu Item Request Master Form.xls'
    print('reading barcode request')
    book = open_workbook(file)
    sheet = book.sheet_by_index(0)
    
    startRow = 0
    for r in range(sheet.nrows):
        if sheet.cell_value(r,0) == 1:
            startRow = r
            break
            
    for r in range(startRow, sheet.nrows):
        if str(sheet.cell_value(r,1)) == '' or sheet.cell_value(r,1) == None:
            continue
            
        row = sheet.row_values(r, 1)
        if str(row[3]) not in barcodeListSet:
            i = BarcodeItem(row[0], row[1], row[2], row[3], row[5])
            barcodeListSet.add(i.upc)
            newItemList.append(i)
        else:
            print(row[0] + ' has duplicate upc [' + str(row[3]) + ']')
            c = input('[S]kip, [N]ew, [A]bort -> ')
            c = c or 's'
            if str(c).casefold() == 'a'.casefold():
                print('Aborting')
                os._exit()
            elif str(c).casefold() == 'n'.casefold():
                print('generating new upc')
                i = BarcodeItem(row[0], row[1], row[2], row[3], row[5])
                i.updateUPC(generateUniqueBarcode(i.upc))
                barcodeListSet.add(i.upc)
                newItemList.append(i)
            else:
                print('skipping')
                continue
                
    print(str(file) + ' imported successfully')
    
def updateBarcodeDatabase():
    file = 'C:/Users/Ryan/workspace/Horizon Barcode Prepare/src/UploadTemp4 Master.xls'
    print('Extracting barcode data base from ' + str(file))
    sheet = open_workbook(file).sheet_by_index(3)
    print('workbook opened')
    for row in range(2, sheet.nrows):
        barcode = sheet.cell_value(row, 3)
        if row == int(sheet.nrows * 0.01):
            print('1%')
        elif row == int(sheet.nrows * 0.25):
            print('25%')
        elif row == int(sheet.nrows / 2):
            print('50%')
        elif row == int(sheet.nrows * 0.75):
            print('75%')
        elif row == int(sheet.nrows * 0.9):
            print('90%')
         
        if barcode not in barcodeListSet:
            barcodeListSet.add(barcode)
            
def outputBarcodeListToFile(file=None):
    file = file or 'barcodeList.txt'
    if barcodeListSet == None:
        print('Barcode List is empty. Aborting.')
        return
    else:
        print('Outputting barcode list to file.')
        with codecs.open(file, 'w+', 'utf8') as f:
            for x in barcodeListSet:
                try:
                    f.write(str(int(x)) + '\r\n')
                except ValueError:
                    f.write(str(x) + '\r\n')
                    
    print('Barcode output complete.')

def shortenName(name):
    print('Shorten\n' + name + ' (' + str(len(name)) + ')\n')
    newName = input('->')
    if len(newName) > 30:
        return shortenName(newName)
    else:
        return newName
    
def pickCategory(name):
    global catChoice
    try:
        print('Pick a Category for\n\n' + name + '\n')
    except UnicodeEncodeError:
        print('Unable to parse item name\n')
        
    catList = list(sorted(BarcodeItem.categories.keys()))
    for i, cat in zip(range(len(catList)), catList):
        print(str(i + 1) + '. ' + str(cat))
        
    if catChoice:
        c = input('\n[' + str(catChoice) + '] ->')
        c = c or catChoice
    else:
        c = input('\n->')
        
    c = int(c)
    
    if c == -1:
        print('Aborting')
        return
    elif c < 1 or c > len(BarcodeItem.categories):
        print('Invalid selection.  Try again.')
        return pickCategory(name)
    else:
        catChoice = c
        return catList[c - 1], BarcodeItem.categories[catList[c - 1]]
    
def pickPrimary(priList):
    global priChoice
    print('\nPick a subcategory')
    #pdb.set_trace()
    for i, cat in zip(range(len(priList)), priList):
        print(str(i + 1) + '. ' + str(cat).title())
    
    if priChoice:
        c = input('\n[' + str(priChoice) + '] ->')
        c = c or priChoice
    else:
        c = input('\n->')
        
    c = int(c)
    
    if c == -1:
        print('Aborting')
        return
    elif c < 1 or c > len(priList):
        print('Invalid selection. Try again.')
        return pickPrimary(priList)
    else:
        priChoice = c
        return priList[c -1].title()
    
def generateUniqueBarcode(barcode, leadingZero=False):
    #pdb.set_trace()
    barcode = str(barcode)
    if barcode in barcodeListSet:
        if barcode[0] == '0':
            leadingZero = True
        if len(barcode) == 12:
            barcode = int(barcode[:-1]) + 1
            if leadingZero:
                barcode = calculateBarcodeChecksum('0' + str(barcode))
            else:
                barcode = calculateBarcodeChecksum(str(barcode))
            return generateUniqueBarcode(barcode, leadingZero)
        else:
            barcode = int(barcode) + 1
            return generateUniqueBarcode(barcode, leadingZero)
    else:
        if leadingZero and len(str(barcode)) == 11:
            return '0' + barcode
        else:
            return barcode
    
def generatePreAccessFile(file=None):
    if newItemList == None:
        print('No items to process.')
        return
    
    file = file or 'Upload_to_Access.xls'
    book = Workbook()
        
    sheet = book.add_sheet(date.today().strftime('%m.%d.%Y'))
    row1 = sheet.row(0)
    headingList = ['Enterprise Number', 'Enterprise Name', 'Price',
                   'Cost', 'Source', 'Category', 'Primary', 'Secondary',
                   'Detail', 'Manufacturer', 'Size/Quantity', 'Station',
                   'Brand', 'UPC CODES']
    
    for col in range(len(headingList)):
        row1.write(col, headingList[col])
        
    for i, item in zip(range(1, len(newItemList) + 1), newItemList):
        print(str(i) + '/' + str(len(newItemList) + 1))
        item.category = 'temp'
        item.primary = 'placeholder'
        row = sheet.row(i)
        row.write(0, item.enterpriseNumber)
        row.write(1, item.name)
        row.write(3, item.cost)
        row.write(4, 'Prepackaged Items')
        row.write(5, item.category)
        row.write(6, item.primary)
        row.write(9, item.manufacturer)
        row.write(10, 'each')
        row.write(13, item.upc)
        
    try:
        book.save(file)
    except PermissionError:
        input('Please close the file and try again.')
        try:
            book.save(file)
        except PermissionError:
            print('You failed.')
            os._exit()
            
    outputBarcodeListToFile()
    print('Pre-Access XLS output completed.')
    
    
def gPAF():
    generatePreAccessFile()
    
importBarcodeDatabase()