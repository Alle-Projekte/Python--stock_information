'''
this is the 3rd-file of the program, there are a total of 3 
that need to be put together. This checks the rows for specifi
rules, read "Rules.txt", and creates a new Excel (.xlsx) file
with the passing stocks. Originally it starts with the output file
of the previous 2 programs. It creates a 2nd Sheet with the preffered
stocks and a 3rd sheet with the accetable stocks. 
'''

import os, re, pyperclip, openpyxl

#Finds what 30% of the Total Cash is and adds it to it. 
def tCashPercent(something):
    percentage = sheet1['D' + str(something)].value * .3
    res = sheet1['D' + str(something)].value + percentage
    return res
    
desktop = os.path.join(os.environ['USERPROFILE'], 'Desktop')
os.chdir(desktop)
wb = openpyxl.load_workbook('test1.xlsx')
wb.create_sheet('2ND')
wb.create_sheet('3rd')
sheet1 = wb.get_sheet_by_name('Sheet')
sheet2 = wb.get_sheet_by_name('2ND')
sheet3 = wb.get_sheet_by_name('3rd')
a, x, counter, i = 1, 1, 1, 2

# This doesn't work but it should aid as the basis to perfect it and make it run. 
# Update 05/28/2020 - I don't remember if the comment directly on top of this still applies. Run the program again.
while i < sheet1.max_row:
    try:
	    if sheet1['B' + str(i)].value > 25000000 and '-' in str(sheet1['G' + str(i)].value) and sheet1['K' + str(i)].value < 0.1000:
		    if sheet1['E' + str(i)].value < tCashPercent(i):
			    while counter < sheet1.max_column + 1:
				    sheet2.cell(row = x, column = counter).value = sheet1.cell(row = i, column = counter).value
				    counter += 1
			    counter = 1
			    x += 1
			    i += 1
		    elif sheet1['E' + str(i)].value > tCashPercent(i) and sheet1['C' + str(i)].value > sheet1['E' + str(i)].value:
			    while counter < sheet1.max_column + 1:
				    sheet3.cell(row = a, column = counter).value = sheet1.cell(row = i, column = counter).value
				    counter += 1
			    counter = 1
			    a += 1
			    i += 1
		    else:
			    i += 1
			    continue
	    else:
		    i += 1
		    continue
    except TypeError:
	    break
	    
wb.save('maybe2.xlsx')
