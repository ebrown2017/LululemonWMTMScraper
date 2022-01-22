import time
import os
import gspread
import pandas as pd
from datetime import datetime
from writeToSpreadsheet import parseToSpreadSheet
from spreadsheetDiff import getNewItems

dow = {
        1: 'monday',
        2: 'tuesday',
        3: 'wednesday',
        4: 'thursday',
        5: 'friday',
        6: 'saturday',
        7: 'sunday',
    }

sa = gspread.service_account(filename = 'api_credentials.json')

# main loop
while True:
    currtime = datetime.utcnow()
    # if it is 11:00am central time execute the scraper
    # 17 00
    if currtime.hour == 17 and currtime.minute == 00:
        print('---Fetching data for today---')

        dayNum = datetime.now().isoweekday()
        day = dow[dayNum]

        # get mens data
        print('Getting new mens data')
        parseToSpreadSheet('N-1z0xcmkZ8t6', 'sales')
        print('Getting new mens data - Done')

        # get womens data
        print('Getting new womens data')
        parseToSpreadSheet('N-1z0xcuuZ8t6', 'sales')
        print('Getting new womens data - Done')

        # get file names for previously scraped data
        newMens = day + '_mens.xls'
        newWomens = day + '_womens.xls'
        if dayNum == 1:
            oldMens = dow[7] + '_mens.xls'
            oldWomens = dow[7] + '_womens.xls'
        else:
            oldMens = dow[dayNum - 1] + '_mens.xls'
            oldWomens = dow[dayNum - 1] + '_womens.xls'

        # compare the new and old parsed data
        print('Creating sheet of new mens items')
        getNewItems(oldMens, newMens)
        print('Creating sheet of new mens items - Done')

        print('Creating sheet of new womens items')
        getNewItems(oldWomens, newWomens)
        print('Creating sheet of new womens items - Done')

        # post the new products to the corresponding gSlide
        print('Uploading the new data to the mens google slide')
        newMensStr = 'new_items_in_' + newMens
        mensContent = pd.read_excel(newMensStr).to_csv().encode()
        sa.import_csv('1ONngG-STAwnu7gpY59cvLyGmmmvcv6O24-7LlLzXwdA', mensContent)
        print('Uploading the new data to the mens google slide - Done')

        print('Uploading the new data to the womens google slide')
        newWomensStr = 'new_items_in_' + newWomens
        womensContent = pd.read_excel(newWomensStr).to_csv().encode()
        sa.import_csv('1oZ77SmNEUZIXB5CfiAQ2yMVioyAWeRfYVsMs6zd1WGE', womensContent)
        print('Uploading the new data to the womens google slide - Done')

        # post all products to the corresponding gslide
        print('Uploading the complete data to the womens google slide')
        allWomensContent = pd.read_excel(newWomens).to_csv().encode()
        sa.import_csv('1YVLizdXH6rR8IapHsr0TKiubMQi_8pIdEJBLmld1Vec', allWomensContent)
        print('Uploading the complete data to the womens google slide - Done')

        print('Uploading the complete data to the mens google slide')
        allMensContent = pd.read_excel(newMens).to_csv().encode()
        sa.import_csv('1DAYIgJBoJM4F5_W9JmQ-iYp0oxJRlWB3ecumsgOpFmE', allMensContent)
        print('Uploading the complete data to the mens google slide - Done')

        # delete old spreadsheets
        os.remove(oldMens)
        print('Removed file - {}'.format(oldMens))
        os.remove(oldWomens)
        print('Removed file - {}'.format(oldWomens))
        os.remove('new_items_in_' + oldMens)
        print('Removed file - {}'.format('new_items_in_' + oldMens))
        os.remove('new_items_in_' + oldWomens)
        print('Removed file - {}'.format('new_items_in_' + oldWomens))

        print('---Fetch complete---')

        time.sleep(60)
    else:
        time.sleep(30)