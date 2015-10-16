#!/usr/bin/python3

import os, shutil
import sys
import re
import codecs
import pdb
import random
from datetime import date
from BarcodeItem import BarcodeItem, calculateBarcodeChecksum
from BarcodeUtilities import safePrint
from tkinter import *
from tkinter import ttk, messagebox, filedialog
from xlwt import Workbook, easyxf
from xlrd import open_workbook

barcodeListSet = set()
newItemList = []
catChoice = None
priChoice = None
itemCount = 0
itemImportCount = 0
itemIndex = 0
sheet = None
lastRow = None
book = None


def open_file(options=None):
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
        return None
    else:
        return file_path


def oF():
    open_file()


def open_directory(dir=None):
    if not dir:
        options = {}
        options['title'] = 'Choose parent directory'
        options['initialdir'] = os.getcwd()
        file_opt = options
        dir_path = filedialog.askdirectory(**file_opt)
        options = None
    else:
        dir_path = dir
    enumerate_files(dir_path)
    for file in os.listdir(dir_path):
        if file.endswith('.xls') and 'Upload_to_Access' not in file:
            if 'POS Export' in file:
                read_export_request(file)
            else:
                read_barcode_request(file)
            generate_pre_access_file()
            archive_file(file)
            
    generate_spot_check('Upload_to_Access.xls')


def oD(dir=None):
    dir = dir or os.getcwd()
    open_directory(dir)


def archive_file(file):
    newPath = os.getcwd() + '/Archive/' + date.today().strftime('%m.%d.%Y')
    if not os.access(newPath, os.F_OK):
        os.mkdir(newPath)
    shutil.move(file, newPath)


def enumerate_files(path):
    for file in os.listdir(path):
        if file.endswith('.xls') and 'Upload_to_Access' not in file:
            count_items(file)
    
    global itemIndex
    global itemCount
    itemIndex = itemCount


def count_items(file):
    book = open_workbook(file)
    sheet = book.sheet_by_index(0)
    
    startRow = 0
    for r in range(sheet.nrows):
        if sheet.cell_value(r,0) == 1:
            startRow = r
            break
    else:
        if sheet.cell_value(r,0) == 'Example':
            startRow = int(r) + 1
            
    for r in range(startRow, sheet.nrows):
        if str(sheet.cell_value(r,1)) == '' or sheet.cell_value(r,1) == None:
            continue
            
        global itemCount
        itemCount += 1
    

def import_barcode_database():
    file = 'barcodeList.txt'
    if os.path.exists(file):
        with codecs.open(file, 'r+', 'utf-8') as f:
            for barcode in f:
                barcode = str(barcode).strip('\r\n')
                if barcode not in barcodeListSet:
                    barcodeListSet.add(barcode)
        print('imported barcodes from ' + file)
    else:
        update_barcode_database()


def read_export_request(file):
    global itemImportCount
    print('reading barcode request')
    book = open_workbook(file)
    sheet = book.sheet_by_index(0)

    for r in range(1, sheet.nrows):
        if str(sheet.cell_value(r,1)) == '' or sheet.cell_value(r,1) == None:
            continue

        row = sheet.row_values(r, 1)
        enterprise = None
        itemImportCount += 1
        try:
            percent = str(int((itemImportCount / itemCount) * 100)) + '%'
        except ZeroDivisionError:
            percent = 'x%'

        print('Import Progress: {0}/{1} {2}\n'.format(itemImportCount, itemCount, percent))

        manufacturer = row[8]
        brand = row[8]
        i = BarcodeItem(row[2], manufacturer, brand, row[1], row[5],
                        row[6], row[7], enterprise=enterprise)
        if i.upc not in barcodeListSet:
            barcodeListSet.add(i.upc)
            newItemList.append(i)
        else:
            safePrint(row[0] + ' has duplicate upc [' + str(i.upc) + ']')
            c = input('[C]ontinue, [S]kip, [N]ew, [A]bort -> ')
            c = c or 'c'
            if (str(c)).casefold() == 'c'.casefold():
                print('continuing')
                barcodeListSet.add(i.upc)
                newItemList.append(i)
            elif str(c).casefold() == 'a'.casefold():
                print('Aborting')
                os._exit(0)
            elif str(c).casefold() == 'n'.casefold():
                print('generating new upc')
                i.updateUPC(generate_unique_barcode(i.upc))
                barcodeListSet.add(i.upc)
                newItemList.append(i)
            else:
                print('skipping')
                continue

    print(str(file) + ' imported successfully\n')


def read_barcode_request(file):
    global itemImportCount
    print('reading barcode request')
    book = open_workbook(file)
    sheet = book.sheet_by_index(0)
    
    startRow = 0
    for r in range(sheet.nrows):
        if sheet.cell_value(r,0) == 1:
            startRow = r
            break
        elif sheet.cell_value(r,0) == 'Example':
            startRow = int(r) + 1
            
    for r in range(startRow, sheet.nrows):
        if str(sheet.cell_value(r,1)) == '' or sheet.cell_value(r,1) == None:
            continue
            
        row = sheet.row_values(r, 1)
        enterprise = None
        itemImportCount += 1
        print('Import Progress: ' + str(itemImportCount) + '/' + str(itemCount) + '\n')
        try:
            print(str(int((itemImportCount / itemCount) * 100)) + '%\n')
        except ZeroDivisionError:
            pass
        if 'MMS' in row[1]: #Utilize predefined MMS number
            enterprise = row[1]
        i = BarcodeItem(row[0], row[1], row[2], row[3], row[5], enterprise=enterprise)
        if i.upc not in barcodeListSet:
            barcodeListSet.add(i.upc)
            newItemList.append(i)
        else:
            safePrint(row[0] + ' has duplicate upc [' + str(i.upc) + ']')
            c = input('[C]ontinue, [S]kip, [N]ew, [A]bort -> ')
            c = c or 'c'
            if (str(c)).casefold() == 'c'.casefold():
                print('continuing')
                #i = BarcodeItem(row[0], row[1], row[2], row[3], row[5])
                barcodeListSet.add(i.upc)
                newItemList.append(i)
            elif str(c).casefold() == 'a'.casefold():
                print('Aborting')
                os._exit(0)
            elif str(c).casefold() == 'n'.casefold():
                print('generating new upc')
               # i = BarcodeItem(row[0], row[1], row[2], row[3], row[5])
                i.updateUPC(generate_unique_barcode(i.upc))
                barcodeListSet.add(i.upc)
                newItemList.append(i)
            else:
                print('skipping')
                continue
                
    print(str(file) + ' imported successfully\n')


def update_barcode_database():
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


def output_barcode_list_to_file(file=None):
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


def shorten_name(name):
    try:
        safePrint('Shorten\n' + name + ' (' + str(len(name)) + ')\n')
    except UnicodeEncodeError:
        return '*!' + name + '*!'
    newName = input('->')
    if len(newName) > 30:
        return shorten_name(newName)
    else:
        return newName


def pick_category(name):
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
        return pick_category(name)
    else:
        catChoice = c
        return catList[c - 1], BarcodeItem.categories[catList[c - 1]]


def pick_primary(priList):
    global priChoice
    print('\nPick a subcategory')
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
        return pick_primary(priList)
    else:
        priChoice = c
        return priList[c -1].title()


def generate_unique_barcode(barcode, leadingZero=False):
    barcode = str(barcode)
    try:
        x = int(barcode)
    except ValueError:
        l = list(barcodeListSet)
        barcode = l[-1] 
        
    if barcode in barcodeListSet:
        if barcode[0] == '0':
            leadingZero = True
        if len(barcode) == 12:
            barcode = int(barcode[:-1]) + 1
            if leadingZero:
                barcode = calculateBarcodeChecksum('0' + str(barcode))
            else:
                barcode = calculateBarcodeChecksum(str(barcode))
            return generate_unique_barcode(barcode, leadingZero)
        else:
            barcode = int(barcode) + 1
            return generate_unique_barcode(barcode, leadingZero)
    else:
        if leadingZero and len(str(barcode)) == 11:
            return '0' + barcode
        else:
            return barcode


def generate_pre_access_file(file=None):
    if newItemList == None:
        print('No items to process.')
        return
    
    file = file or 'Upload_to_Access.xls'
    global sheet, lastRow, book
    if not sheet:
        book = Workbook()
        sheet = book.add_sheet('Sheet 1')
        row1 = sheet.row(0)
        lastRow = 0
        headingList = ['Enterprise Number', 'Enterprise Name', 'Price',
                       'Cost', 'Source', 'Category', 'Primary', 'Secondary',
                       'Detail', 'Manufacturer', 'Size/Quantity', 'Station',
                       'Brand', 'UPC CODES']
        
        for col in range(len(headingList)):
            row1.write(col, headingList[col])
        
    for i, item in zip(range((lastRow + 1), len(newItemList) + lastRow + 1), newItemList):
        print(str(i) + '/' + str(len(newItemList)) + ' ' + item.name)
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
        lastRow = i
        print('Last Row:', lastRow)
        
    try:
        book.save(file)
    except PermissionError:
        input('Please close the file and try again.')
        try:
            book.save(file)
        except PermissionError:
            print('You failed.')
            os._exit()
            
    output_barcode_list_to_file()
    newItemList.clear()
    print('Pre-Access XLS output completed.\n')


def gPAF():
    generate_pre_access_file()


def generate_spot_check(file=None):
    file = file or open_file()
    book = open_workbook(file)
    sheet = book.sheet_by_index(0)
    spotCheckMap = {}
    
    toBeCheckedList = random.sample(range(1, sheet.nrows), round(sheet.nrows * .2))
    
    for row in toBeCheckedList:
        spotCheckMap[sheet.cell_value(row, 1)] = sheet.cell_value(row, 13)
        
    with codecs.open('spot_check.csv', 'w+', 'utf8') as f:
            for item, sku in spotCheckMap.items():
                f.write(str(sku) + ',' + str(item) + '\r\n')
                
    print('Spot Check Generation Complete\n')
    os.startfile(file)

def gSC():
    generate_spot_check()
        
    
import_barcode_database()
if not sys.flags.interactive:
    open_directory()