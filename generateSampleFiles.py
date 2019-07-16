#! python3
# generateSampleFiles.py

# Generates records for two months for six employees

# import random and excel
import random, openpyxl

# create an array of 6 abbreviated names with 2 initials
employees = ['BW', 'MM', 'EW', 'JH', 'PH', 'ST']

print('Please wait while I generate the sample excel files...')

# loop through the list of employees
for emp in employees:

    # loop through a range of 1-50
    for week in range(1, 51):

        # create an excel file name, [emp]_[week]
        name = "%s_%d.xlsx" %(emp, week)

        # generate a number between 10 and 60 to represent work hours
        hours = random.randint(10, 61)

        # open a new workbook, select active sheet
        book = openpyxl.Workbook()
        sheet = book.active

        # print header
        sheet['A1'].value = 'Employee: '
        sheet['B1'].value = emp
        sheet['C1'].value = 'Week: '
        sheet['D1'].value = week

        # print hours in A2
        sheet['A2'].value = hours

        # save and close file
        book.save(name)
        book.close

print('File generation complete!')
        
        
        

        
