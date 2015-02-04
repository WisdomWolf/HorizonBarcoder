#!/usr/bin/python

class BarcodeItem:
    """An object to simplify item property assignment"""
    
    catBeverages = ['hot', 'cold']
    catHot = ['coffee', 'tea', 'other']
    catSnacks = ['gum', 'energy and health', 'candy', 'frozen/refigerated',
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
    
    def __init__(self, name, manufacturer, brand, upc, cost, cat=None, pri=None):
        self.name = name.strip()
        if len(self.name) > 30:
            self.name = shortenName(self.name)
        self.manufacturer = manufacturer.strip()
        self.brand = brand.strip()
        self.upc = str(upc).strip().split(sep='.', maxsplit=1)[0]
        self.cost = cost
        self.enterpriseNumber = 'MMS-' + self.upc
        self.category = cat
        self.primary = pri
        
    def updateUPC(self, upc):
        self.upc = str(upc).strip()
        self.enterpriseNumber = 'MMS-' + self.upc
        
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
    print('Shorten\n' + name + ' (' + str(len(name)) + ')\n')
    newName = input('->')
    if len(newName) > 30:
        return shortenName(newName)
    else:
        return newName