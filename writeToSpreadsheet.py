from datetime import datetime
import xlwt
import time
from xlwt import Workbook
from scrape import scrape

def setFileName(nValue, category):
    time = ''

    dow = {
        1: 'monday',
        2: 'tuesday',
        3: 'wednesday',
        4: 'thursday',
        5: 'friday',
        6: 'saturday',
        7: 'sunday',
    }

    time += dow[datetime.now().isoweekday()]

    if nValue.startswith('N-1z0xcuuZ8t'):
        time += '_womens'
    else:
        time += '_mens'

    # Canadian data
    if nValue[-1] == '5':
        time += '_canada'

    return time

def parseToSpreadSheet(nValue, category):
    data = scrape(nValue, category)

    wb = Workbook()
    sheet = wb.add_sheet('Sheet 1')

    # label the spreadsheet
    sheet.write(0, 0, 'Product Name')
    sheet.write(0, 1, 'Price')
    sheet.write(0, 2, 'Origional Price')
    sheet.write(0, 3, 'Available Sizes')
    sheet.write(0, 4, 'Percent Off')
    sheet.write(0, 5, 'Product Category')
    sheet.write(0, 6, 'Colors Available')
    sheet.write(0, 7, 'Skus Available')
    sheet.write(0, 8, 'Link')

    def insertProduct(lineNum, product):
        # create sku and color lists
        skus = []
        colors = []
        for variant in product['skuStyleOrder']:
            skus.append(variant['sku'])
            colors.append(variant['colorName'])

        # calculate discount percentage
        avgReg = sum(float(e) for e in product['listPrice']) / len(product['listPrice'])
        avgDis = sum(float(e) for e in product['productSalePrice']) / len(product['productSalePrice'])
        discount = 1 - (avgDis / avgReg)

        sheet.write(lineNum, 0, product['displayName'])
        sheet.write(lineNum, 1, ', '.join('$' + str(e) for e in product['productSalePrice']))
        sheet.write(lineNum, 2, ', '.join('$' + str(e) for e in product['listPrice']))
        sheet.write(lineNum, 3, ', '.join(product['allAvailableSizes']))
        sheet.write(lineNum, 4, "{:.2%}".format(discount))
        sheet.write(lineNum, 5, product['parentCategoryUnifiedId'])
        sheet.write(lineNum, 6, ', '.join(colors))
        sheet.write(lineNum, 7, ', '.join(skus))
        sheet.write(lineNum, 8, "https://shop.lululemon.com" + product['pdpUrl'])

    # loop through all products
    counter = 1
    for product in data[2]['data']['categoryDetails']['products']:

        # for some reason this is parsing products that do not have a sale price
        # we want to exclude these so we execute this 
        if product['productSalePrice'] == []:
            continue

        insertProduct(counter, product)
        counter += 1

    # save spreadsheet
    fileData = setFileName(data[0], data[1])
    spreadsheetPath = fileData + '.xls'
    wb.save(spreadsheetPath)
