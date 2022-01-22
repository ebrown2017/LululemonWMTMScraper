import xlwt
import xlrd
from xlwt import Workbook
from datetime import datetime

wb = Workbook()
sheet = wb.add_sheet('Sheet 1')

def formatOutputSheet():
    # label the spreadsheet
    sheet.write(0, 1, 'Product Name')
    sheet.write(0, 2, 'Price')
    sheet.write(0, 3, 'Origional Price')
    sheet.write(0, 4, 'Available Sizes')
    sheet.write(0, 5, 'Percent Off')
    sheet.write(0, 6, 'Product Category')
    sheet.write(0, 7, 'Colors Available')
    sheet.write(0, 8, 'Skus Available')
    sheet.write(0, 9, 'Link')


def getNewItems(sheet1, sheet2):
    wb1 = xlrd.open_workbook(sheet1)
    wb2 = xlrd.open_workbook(sheet2)

    s1 = wb1.sheet_by_index(0)
    s2 = wb2.sheet_by_index(0)

    data1 = {}

    # store product names and colors for comparison
    for i in range(1, s1.nrows):
        data1[(s1.cell_value(i, 1))] = (s1.cell_value(i, 7))

    formatOutputSheet()

    def addDiffRow(i, index, colorList=None):
        sheet.write(index, 0, index)
        sheet.write(index, 1, s2.cell_value(i, 1))
        sheet.write(index, 2, s2.cell_value(i, 2))
        sheet.write(index, 3, s2.cell_value(i, 3))
        sheet.write(index, 4, s2.cell_value(i, 4))
        sheet.write(index, 5, s2.cell_value(i, 5))
        sheet.write(index, 6, s2.cell_value(i, 6))
        if colorList is not None:
            sheet.write(index, 7, ', '.join(newList))
        else:
            sheet.write(index, 7, s2.cell_value(i, 7))
        sheet.write(index, 8, s2.cell_value(i, 8))
        sheet.write(index, 9, s2.cell_value(i, 9))

    counter = 1
    for i in range(1, s2.nrows):
        if (s2.cell_value(i, 1)) in data1:
            # same product name found
            if data1[s2.cell_value(i, 1)] == s2.cell_value(i, 7):
                # same colorways found, do nothing
                continue
            else:
                # discrepency in colors found, doesn't necessarily mean
                # that a new color is here, could've sold out of old one
                list1 = list(data1[s2.cell_value(i, 1)].split(', '))
                list2 = list(s2.cell_value(i, 7).split(', '))
                newList = list(set(list2) - set(list1))

                if newList == []:
                    # no new colors, must have sold out of old one
                    continue
                else:
                    for color in newList:
                        print("---{}---- new color: {}".format(s2.cell_value(i, 1), color))
                    addDiffRow(i, counter, newList)
                    counter += 1
        else:
            # new product found, add to diff spreadsheet
            print("---{}--- newly added".format(s2.cell_value(i, 1)))
            addDiffRow(i, counter)
            counter += 1

    # all new stuff should be added, now we just have to save file
    path = './spreadsheets/' + 'new_items_in' + sheet2
    wb.save(path)