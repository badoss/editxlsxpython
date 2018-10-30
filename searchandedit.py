import xlwings as xw
import string
from xlwings.constants import DeleteShiftDirection,InsertShiftDirection
###################################################################

bookName = r'///test.xlsx'  #local file
wb = xw.Book(bookName)
sheetName = 'Sheet1'  #sheet name
sht = wb.sheets[sheetName]

###################################################################
print ('""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""')
print ('""""""""""""""""""""""""///  EDIT  EXCEL  ///""""""""""""""""""""""" :')
print ('""""""""""""""""""""""""""""""""""""""""""""""""""""""'' babobaboo""""')
print('Keyword Search :')
word = input()
print('word Edit:')
wordd = input()
###################################################################
search = word #search word to edit

myCell = wb.sheets[sheetName].api.UsedRange.Find(search)
A = myCell.address.replace('$', '', 1)
A = myCell.address.replace('$', '', 2)
print ('Address  :'+A)
print ('=== success ===')
sht.range(A).value = wordd #new word edit 


