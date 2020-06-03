'''
The goal of this program is to extract a list of stock symbols from
an Excel file, get it's information from Yahoo Finance, and store
it in a new Excel file. This is where the final form of the program
will be written in and it's also were it started. 
'''

import bs4, requests, pyperclip, os, openpyxl
from bs4 import BeautifulSoup as soup

#Changes the current file path to your desktop.
desktop = os.path.join(os.environ['USERPROFILE'], 'Desktop')
#desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
os.chdir(desktop) 

#This should be a funxtion and will loop through an Excel file to extract stock symbols
#and store it in the 'symbol" list. 
symbol = 'AMZN' #'GOOGL', 'AMZN', 'NVDA', 'NFLX']

#Takes the symbol, extracts the page and stores needed information in a variable.
url = requests.get(r'https://finance.yahoo.com/quote/' + symbol + r'/key-statistics?p=' + symbol)
html = soup(url.text, 'html.parser')
start_html = html.find('div', class_="Fl(start) W(50%) smartphone_W(100%)")
dude1 = start_html.findAll('td')

#Turns parsed information to text and stores it into a list
list1 = []  

for x in dude1:
    list1.append(x.text)

#Extracts even indexes from list1 list and stores it into list2
list2 = []

for i in range(5,len(list1),2):
    list2.append(list1[i])

#Extracts odd indexes from list1 list and stores it into list3
list3 = []

for i in range(4,len(list1),2):
    list3.append(list1[i])

#Creates an Excel workbook and gets its 1st sheet
wb = openpyxl.Workbook()
sheet = wb.get_sheet_by_name('Sheet')

#Specifies where to start and stores the list3 information into the Excel workbook.
row, column = 1, 2 

for a in list3:
    sheet.cell(row, column).value = a
    column += 1

#Specifies where to start and stores the list2 information into the Excel workbook.
row, column = 2, 2

for b in list2:
    sheet.cell(row, column).value = b
    column += 1

#Saves the Excel workbook in the current directory    
wb.save('example.xlsx')
