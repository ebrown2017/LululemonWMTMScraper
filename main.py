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
        parseToSpreadSheet('N-1z0xcmkZ8t6', 'sale')
        print('Getting new mens data - Done')

        # get womens data
        print('Getting new womens data')
        parseToSpreadSheet('N-1z0xcuuZ8t6', 'sale')
        print('Getting new womens data - Done')

        # get mens data - Canada
        print('Getting new mens data Canada')
        parseToSpreadSheet('N-1z0xcmkZ8t5', 'sale')
        print('Getting new mens data Canada - Done')

        # get womens data - Canada
        print('Getting new mens data Canada')
        parseToSpreadSheet('N-1z0xcuuZ8t5', 'sale')
        print('Getting new mens data Canada - Done')

        # get file names for previously scraped data
        newMens = day + '_mens.xls'
        newWomens = day + '_womens.xls'
        newMensCanada = day + '_mens_canada.xls'
        newWomensCanada = day + '_womens_canada.xls'
        if dayNum == 1:
            oldMens = dow[7] + '_mens.xls'
            oldWomens = dow[7] + '_womens.xls'
            oldMensCanada = dow[7] + '_mens_canada.xls'
            oldWomensCanada = dow[7] + '_womens_canada.xls'
        else:
            oldMens = dow[dayNum - 1] + '_mens.xls'
            oldWomens = dow[dayNum - 1] + '_womens.xls'
            oldMensCanada = dow[dayNum - 1] + '_mens_canada.xls'
            oldWomensCanada = dow[dayNum - 1] + '_womens_canada.xls'

        # compare the new and old parsed data
        print('Creating sheet of new mens items')
        getNewItems(oldMens, newMens)
        print('Creating sheet of new mens items - Done')

        print('Creating sheet of new womens items')
        getNewItems(oldWomens, newWomens)
        print('Creating sheet of new womens items - Done')

        print('Creating sheet of new mens items Canada')
        getNewItems(oldMensCanada, newMensCanada)
        print('Creating sheet of new mens items Canada - Done')

        print('Creating sheet of new womens items Canada')
        getNewItems(oldWomensCanada, newWomensCanada)
        print('Creating sheet of new womens items Canada - Done')

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
        print('Uploading the complete data to the mens google slide')
        allMensContent = pd.read_excel(newMens).to_csv().encode()
        sa.import_csv('1DAYIgJBoJM4F5_W9JmQ-iYp0oxJRlWB3ecumsgOpFmE', allMensContent)
        print('Uploading the complete data to the mens google slide - Done')

        print('Uploading the complete data to the womens google slide')
        allWomensContent = pd.read_excel(newWomens).to_csv().encode()
        sa.import_csv('1YVLizdXH6rR8IapHsr0TKiubMQi_8pIdEJBLmld1Vec', allWomensContent)
        print('Uploading the complete data to the womens google slide - Done')

        # ----------------CANADA---------------------

        # post the new products to the corresponding gSlide Canada
        print('Uploading the new data to the mens Canada google slide')
        newMensStrCanada = 'new_items_in_' + newMensCanada
        mensContentCanada = pd.read_excel(newMensStrCanada).to_csv().encode()
        sa.import_csv('19ceDZDRkjLjvJLUmrlg845zv6JmLNGgtlyX2-uI4Pks', mensContentCanada)
        print('Uploading the new data to the mens Canada google slide - Done')

        print('Uploading the new data to the womens Canada google slide')
        newWomensStrCanada = 'new_items_in_' + newWomensCanada
        womensContentCanada = pd.read_excel(newWomensStrCanada).to_csv().encode()
        sa.import_csv('1em6_opbMxpIp_I3USpzgXDBW21jwDHr2fKwMqPLt4qU', womensContentCanada)
        print('Uploading the new data to the womens Canada google slide - Done')

        # post all products to the corresponding gslide
        print('Uploading the complete data to the mens Canada google slide')
        allMensContentCanada = pd.read_excel(newMensCanada).to_csv().encode()
        sa.import_csv('1lw6Hd_c89MoCIDgha9J8GTQF7Cc5Kpe-T8Vp1w48uqY', allMensContentCanada)
        print('Uploading the complete data to the mens Canada google slide - Done')

        print('Uploading the complete data to the womens Canada google slide')
        allWomensContentCanada = pd.read_excel(newWomensCanada).to_csv().encode()
        sa.import_csv('1bJcQZb0svA7plRx-ghdzHmUhv3CRypLqjtwmmPihQqg', allWomensContentCanada)
        print('Uploading the complete data to the womens Canada google slide - Done')

        # delete old spreadsheets
        os.remove(oldMens)
        print('Removed file - {}'.format(oldMens))
        os.remove(oldWomens)
        print('Removed file - {}'.format(oldWomens))
        os.remove(oldMensCanada)
        print('Removed file - {}'.format(oldMensCanada))
        os.remove(oldWomensCanada)
        print('Removed file - {}'.format(oldWomensCanada))
        os.remove('new_items_in_' + oldMens)
        print('Removed file - {}'.format('new_items_in_' + oldMens))
        os.remove('new_items_in_' + oldWomens)
        print('Removed file - {}'.format('new_items_in_' + oldWomens))
        os.remove('new_items_in_' + oldMensCanada)
        print('Removed file - {}'.format('new_items_in_' + oldMensCanada))
        os.remove('new_items_in_' + oldWomensCanada)
        print('Removed file - {}'.format('new_items_in_' + oldWomensCanada))

        print('---Fetch complete---')

        time.sleep(60)
    else:
        time.sleep(30)