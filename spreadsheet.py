from datetime import datetime
import xlwt
from xlwt import Workbook
from scrape import scrape

def getFileData(nValue, category):
    time = str(datetime.now())
    time = time[:len(time) - 7]
    time = time.replace(' ', '_')
    time = time.replace(':', '-')
    return time + '_' + nValue + '_' + category

def parseToSpreadSheet(data):
    wb = Workbook()
    sheet = wb.add_sheet('Sheet 1')

    # label the spreadsheet
    sheet.write(0, 1, 'Product Name')
    sheet.write(0, 2, 'Price')
    sheet.write(0, 3, 'Origional Price')
    sheet.write(0, 4, 'Available Sizes')
    sheet.write(0, 5, 'Product Category')
    sheet.write(0, 6, 'Colors Available')
    sheet.write(0, 7, 'Skus Available')
    sheet.write(0, 8, 'Link')

    def insertProduct(lineNum, product):
        skus = []
        colors = []
        for variant in product['skuStyleOrder']:
            skus.append(variant['sku'])
            colors.append(variant['colorName'])

        sheet.write(lineNum, 0, lineNum)
        sheet.write(lineNum, 1, product['displayName'])
        sheet.write(lineNum, 2, str(product['productSalePrice']))
        sheet.write(lineNum, 3, str(product['listPrice']))
        sheet.write(lineNum, 4, str(product['allAvailableSizes']))
        sheet.write(lineNum, 5, product['parentCategoryUnifiedId'])
        sheet.write(lineNum, 6, str(colors))
        sheet.write(lineNum, 7, str(skus))
        sheet.write(lineNum, 8, "https://shop.lululemon.com" + product['pdpUrl'])

    # loop through all products
    for i, product in enumerate(data[2]['data']['categoryDetails']['products']):
        insertProduct(i+1, product)

    # save spreadsheet
    fileData = getFileData(data[0], data[1])
    spreadsheetPath = './spreadsheets/' + fileData + '.xls'
    wb.save(spreadsheetPath)



parseToSpreadSheet(scrape())