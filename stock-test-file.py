'''
This is the 2nd-file of the program, there are a total of 3 
that need to be put together. This checks for faulty stocks 
that populated with a date in B-column (faultyStocks()), and stocks
that popu lated with "N/A" in b-column (noAnswer()). 
'''

import pyperclip, openpyxl, re, os

#Splits values with dots (.), ie.-(200.44M).
def stringSplitr(something):
	temp = re.compile('(\-*\d+\.\d+)(\w)')
	res = temp.match(something).groups()
	return res

#Splits values without dots (.), ie.-(200M).
def simpleSplitr(something):
	temp = re.compile('(\-*\d+)(\w)')
	res = temp.match(something).groups()
	return res

#Compares cell value to check if it populated with a following format date (Marc 02, 2020).
def compare(something):
	try:
		if something == re.findall(r'(\w+ \d+\, \d+)', something)[0]:
			return True
	except IndexError or TypeError:
		return False

#Performs the cell comparison through the compare() function to check for dates. 
def faultyStocks():
        stocks = []
        count = 2
        i = 1
        while i < sheet1.max_row:
                var = sheet1['B' + str(count)].value
                if var == 'N/A':
                        stocks.append(sheet1['A' + str(count)].value)
                        print(sheet1['A' + str(count)].value)
                elif compare(var) == True:
                        stocks.append(sheet1['A' + str(count)].value)
                        print(sheet1['A' + str(count)].value)
                else:
                        i += 1
                        count += 1
                        continue
                i += 1
                count += 1
        return stocks

#Performs the cell comparison to check if content is "N/A"
def noAnswer():
        letter = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']
        counter = 0
        i, count = 1, 2
        while counter < len(letter):
                while i < sheet1.max_row:
                        var1 = sheet1[letter[counter] + str(count)].value
                        if var1 == 'N/A':
                                sheet1[letter[counter] + str(count)].value = 0
                                count += 1
                                i += 1
                        else:
                                count += 1
                                i += 1
                counter += 1
                i, count = 1, 2

desktop = os.path.join(os.environ['USERPROFILE'], 'Desktop')
os.chdir(desktop)

#Need to make a function to check if filename exists
file = input('Enter the Excel filename without extension: ')
ext = '.xlsx'
wb = openpyxl.load_workbook(file + ext)
sheet1 = wb.get_sheet_by_name('Sheet')
letter = ['B', 'C', 'D', 'E', 'H']
lCounter = 0
i, count = 1, 2

if len(faultyStocks()) == 0:
        print('No faulty stocks found.')
else:
        print(faultyStocks())
        
while lCounter < 5:
	while i < sheet1.max_row:
		var1 = sheet1[letter[lCounter] + str(count)].value
		if var1 == 'N/A':
			count += 1
			i += 1
			continue
		elif compare(var1) == True:
			print(sheet1['A' + str(count)].value)
			count += 1
			i += 1
			continue
		if '.' in var1:
			var2 = stringSplitr(var1)
		else:
			var2 = simpleSplitr(var1)
		if 'k' in var2[1]:
			sheet1[letter[lCounter] + str(count)].value = float(var2[0]) * 1000
		elif 'M' in var2[1]:
			sheet1[letter[lCounter] + str(count)].value = float(var2[0]) * 1000000
		elif 'B' in var2[1]:
			sheet1[letter[lCounter] + str(count)].value = float(var2[0]) * 1000000000
		elif 'T' in var2[1]:
			sheet1[letter[lCounter] + str(count)].value = float(var2[0]) * 1000000000000
		else:
			print(var1)
		count += 1
		i += 1
	lCounter += 1
	i, count = 1, 2
'''Need to make a process to compile the numbers from column G, I and J,
separate the % sign from the number and multiply the number by 0.01'''

noAnswer()

wb.save('test1.xlsx')
