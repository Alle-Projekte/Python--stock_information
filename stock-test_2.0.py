'''
The goal of this program is to extract a list of stock symbols from
an Excel file, get it's information from Yahoo Finance, and store
it in a new Excel file. BTW idiot, this is the 1st-file. Learn how to
document next time -_-, and come up with better names for these god damned
variables. Some of them make no sense.
'''

import bs4, requests, pyperclip, os, openpyxl, time, re
from bs4 import BeautifulSoup as soup


def getURLtd(something):
    #Takes the symbol, extracts the page and stores
    #needed information in a variable.
    url = requests.get(r'https://finance.yahoo.com/quote/' + something + r'/key-statistics?p=' + something)
    html = soup(url.text, 'html.parser')
    starter_html = html.find('div', class_="Fl(start) smartphone_W(100%) W(100%)")
    start_html = html.find('div', class_="Fl(start) W(50%) smartphone_W(100%)")
    end_html = html.find('div', class_="Fl(end) W(50%) smartphone_W(100%)")
    dude = starter_html.findAll('td')
    dude1 = start_html.findAll('td')
    dude2 = end_html.findAll('td')

    h = 0
    for x in dude[0:2]:
        dude1.insert(h, x)
        h += 1

    dude1.extend(dude2)

    return dude1


def tdToText(something1):
    #Turns parsed information to text and stores it into a list
    list1 = []

    for x in something1:
        list1.append(x.text)

    return list1


def getList1Even(something2):
    #Extracts even indexes from list1 list and stores it into list2
    list2 = []

    for i in range(1,len(something2),2):
        list2.append(something2[i])

    return list2


def getList1Odd(something3):
    #Extracts odd indexes from list1 list and stores it into list3
    list3 = []

    for i in range(0,len(something3),2):
        list3.append(something3[i])

    return list3


#Changes the current file path to your desktop.
desktop = os.path.join(os.environ['USERPROFILE'], 'Desktop')
os.chdir(desktop)

#Creates an Excel workbook and gets its 1st sheet
wb = openpyxl.Workbook()
sheet = wb.get_sheet_by_name('Sheet')

#This should be a funxtion and will loop through an Excel file to extract stock
#symbols and store it in the 'symbol" list.
#------------------------------------------------------------------------------------------
fileName = input('Enter the name of the file: ')
wb1 = openpyxl.load_workbook(fileName.lower() + '.xlsx')
sheet1 = wb1.get_sheet_by_name('Basic View')
row = 1
symbol = []
#row < 1040
while row < sheet1.max_row:
    stock = sheet1['A' + str(row)].value
    symbol.append(stock)
    row += 1

#symbol1 = symbol[888:1041]
# symbol1 = ['AVCT', 'OCGN', 'OCSL', 'ODP', 'ONCY', 'ONTX', 'OPTT', 'ORMP', 'OSG', 'PDLI'] #['AMZN', 'DYNT', 'GOOGL',] 'AMZN', 'NVDA', 'NFLX']

url = requests.get(r'https://finance.yahoo.com/quote/GOOGL/key-statistics?p=GOOGL')
html = soup(url.text, 'html.parser')
starter_html = html.find('div', class_="Fl(start) smartphone_W(100%) W(100%)")
start_html = html.find('div', class_="Fl(start) W(50%) smartphone_W(100%)")
end_html = html.find('div', class_="Fl(end) W(50%) smartphone_W(100%)")
dude = starter_html.findAll('td')
dude1 = start_html.findAll('td')
dude2 = end_html.findAll('td')

h = 0
for x in dude[0:2]:
    dude1.insert(h, x)
    h += 1

dude1.extend(dude2)

list1 = []
for x in dude1:
    list1.append(x.text)

list3 = getList1Odd(list1)
list3_1 = [list3[x] for x in (0, 7, 15, 17, 23, 24, 33, 34, 35, 38)]

#Stores symbol and data name, on the left and top colunm and rows respectively,
#into the Excel workbook.
row, column = 1, 2
for a in list3_1:
    sheet.cell(row, column).value = a
    column += 1

row, column = 2, 1
for j in symbol:
    sheet.cell(row, column).value = j
    row += 1

#This part writes the actual data into the file
count = 2
lrv = []

for w in symbol:
    try:
        td = getURLtd(w)
        tdList = tdToText(td)
        listEven = getList1Even(tdList)
        list2 = [listEven[x] for x in (0, 7, 15, 17, 23, 24, 33, 34, 35, 38)]

        while bool(re.match(r'(\w+ \d+\, \d+)', list2[0])) or list2[0] == "N/A":
            print(w + " -X") #- For testing this loop
            time.sleep(150)
            td = getURLtd(w)
            tdList = tdToText(td)
            listEven = getList1Even(tdList)
            list2 = [listEven[x] for x in (0, 7, 15, 17, 23, 24, 33, 34, 35, 38)]
            
        #Specifies where to start and stores the list2 information into the
        #Excel workbook.
        row, column = count, 2
        for b in list2:                    
            sheet.cell(row, column).value = b
            column += 1

        count += 1

    except AttributeError:
        '''The AttributeError comes up when the webpage that you are pulling th request from
        doesn't exist. Sometimes Yahoo Finance does not have a "Statistics" tab for the
        stock, so it re-directs you to the summary page. When the program tries to get the
        information from the page, it doesn't find what is looking for and outputs the
        AttributeError. '''
        print(w)
        lrv.append(count)
        count += 1

counter = 0
for m in lrv:
    try:
        time.sleep(300)
        w = sheet["A" + str(m)].value
        td = getURLtd(w)
        tdList = tdToText(td)
        listEven = getList1Even(tdList)
        list2 = [listEven[x] for x in (0, 7, 15, 17, 23, 24, 33, 34, 35, 38)]

        while bool(re.match(r'(\w+ \d+\, \d+)', list2[0])) or list2[0] == "N/A":
            time.sleep(150)
            td = getURLtd(w)
            tdList = tdToText(td)
            listEven = getList1Even(tdList)
            list2 = [listEven[x] for x in (0, 7, 15, 17, 23, 24, 33, 34, 35, 38)]
            
        #Specifies where to start and stores the list2 information into the
        #Excel workbook.
        row, column = m, 2
        for b in list2:                    
            sheet.cell(row, column).value = b
            column += 1

    except AttributeError:
        sheet.delete_rows(m - counter, 1)
        counter += 1

#Saves the Excel workbook in the current directory
wb.save('Test.xlsx')
