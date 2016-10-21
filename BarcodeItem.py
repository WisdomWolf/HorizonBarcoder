#!/usr/bin/python
from BarcodeUtilities import safePrint
import pdb

class BarcodeItem:
    """An object to simplify item property assignment"""
    
    catBeverages = ['hot', 'cold']
    catHot = ['coffee', 'tea', 'other']
    catSnacks = ['gum', 'energy and health', 'candy', 'frozen/refrigerated',
                 'nuts', 'chocolate', 'chips', 'cookies', 'pastries', 'other']
    catCold = ['iced tea', 'juice', 'milk', 'soda', 'water', 'energy',
               'not carbonated', 'other']
    catGiftItems = ['cards', 'plants', 'flowers', 'plush', 'jewelery', 'seasonal',
                    'balloons', 'gift baskets', 'toys and games', 'other', 'baby',
                    'picture frames', 'home decor', 'stationery', 'scents']
    catHealthAndBeauty = ['medical', 'personal care', 'infant care', 'other']
    catElectronics = ['accessories', 'other']
    catApparel = ['logo items', "women's", "men's", 'children/baby', 'uniforms']
    catMedia = ['newspapers', 'books/magazines', 'dvd/music']
    categories = {'Beverages' : catBeverages, 'Hot' : catHot, 'Cold' : catCold, 
                  'Gift_Items' : catGiftItems, 'Health and Beauty' : catHealthAndBeauty,
                  'Electronics' : catElectronics, 'Apparel' : catApparel, 
                  'Media' : catMedia, 'Snacks' : catSnacks}
    
    def __init__(self, name, manufacturer, brand, upc, cost=0,
                cat='temp', pri='placeholder', enterprise=None):
                
        self.name = str(name).strip()
        if len(self.name) > 30:
            self.name = self.name[:30]
        self.manufacturer = str(manufacturer).strip()
        self.brand = str(brand).strip()
        self.upc = str(upc).strip().split(sep='.', maxsplit=1)[0]
        self.cost = cost
        if enterprise:
            self.enterpriseNumber = 'MMS-{0}'.format(enterprise.strip('MMS- '))
        else:
            self.enterpriseNumber =  'MMS-{0}'.format(self.upc.strip('MMS-'))
        self.has_defined_MMS = enterprise
        self.category = cat
        self.primary = pri
        if self.category != 'temp' or self.has_defined_MMS:
            self.source = 'Prepared Item Webtrition'
        else:
            self.source = 'Prepackaged Items'
        
    def updateUPC(self, upc):
        self.upc = str(upc).strip()
        if not self.has_defined_MMS:
            self.enterpriseNumber = 'MMS-{0}'.format(self.upc.strip('MMS-'))


def calculateBarcodeChecksum(barcode):
    barcode = str(barcode)
    if len(barcode) != 11:
        return barcode
    else:
        oddSum = 0
        evenSum = 0
        for i in range(len(barcode)):
            if i % 2 == 0:
                oddSum += int(barcode[i])
            else:
                evenSum += int(barcode[i])
        oddSum *= 3
        tempSum = (oddSum + evenSum) % 10
        if tempSum > 0:
            return barcode + str(10 - tempSum)
        else:
            return barcode + str(tempSum)


def shortenName(name):
    try:
        safePrint('Shorten\n' + name + ' (' + str(len(name)) + ')\n')
        print('->------------------------------')
        newName = input('->')
        if len(newName) > 30:
            return shortenName(newName)
        else:
            return newName
    except UnicodeEncodeError:
        print('There was an error processing this name')
        return '!' + name