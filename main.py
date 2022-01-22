from datetime import datetime
import time
import os
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

# main loop, do nothing
while True:
    currtime = datetime.utcnow()
    # if it is 11:00am central time execute the scraper
    if currtime.hour == 17 and currtime.minute == 00:
        dayNum = datetime.now().isoweekday()
        day = dow[dayNum]

        # get womens data
        print('Getting new mens data')
        parseToSpreadSheet('N-1z0xcmkZ8t6', 'sales')
        print('Getting new mens data - Done')
        
        time.sleep(10)

        # get womens data
        print('Getting new womens data')
        parseToSpreadSheet('N-1z0xcuuZ8t6', 'sales')
        print('Getting new womens data - Done')

        time.sleep(10)

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

        time.sleep(3)

        print('Creating sheet of new womens items')
        getNewItems(oldWomens, newWomens)
        print('Creating sheet of new womens items - Done')

        # delete old spreadsheets
        os.remove(oldMens)
        print('Removed file - {}'.format(oldMens))
        os.remove(oldWomens)
        print('Removed file - {}'.format(oldWomens))
        os.remove('new_items_in_' + oldMens)
        print('Removed file - {}'.format('new_items_in_' + oldMens))
        os.remove('new_items_in_' + oldWomens)
        print('Removed file - {}'.format('new_items_in_' + oldWomens))

        time.sleep(60)
    else:
        time.sleep(30)