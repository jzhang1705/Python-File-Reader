# Authors: Ajay Desai, Jeffrey Zhang
# Date Created: 7/18/20


# Reading an excel file using Python 
import xlrd 

# Writing to an excel sheet
import xlwt
from xlwt import Workbook

# Data Collections
totalAverageLoansCollection = {}
undergradAverageLoansCollection = {}
gradAverageLoansCollection = {}
averageLoans = {}

# Calculating total amount of loans
def total_Amount_Of_Loans(loansRow,column, wb):
    totalLoans = 0
    sheet = wb.sheet_by_index(0) 
    for loansRow in range(loansRow, sheet.nrows):
        if type(sheet.cell_value(loansRow, column)) is float:
            # print(sheet.cell_value(loansRow, column))
            totalLoans += sheet.cell_value(loansRow,column)
    return totalLoans

# Calculating total number of recipients
def total_Number_Of_Recipients(recipientRow, column, wb):
    sheet = wb.sheet_by_index(0) 
    totalRecipients = 0
    for recipientRow in range(recipientRow, sheet.nrows):
        if type(sheet.cell_value(recipientRow, column)) is float:
            totalRecipients += sheet.cell_value(recipientRow, column)
    return totalRecipients

def operate(loc, key):
    # Workbook is created
    wb = Workbook()

    # To open Workbook 
    wb = xlrd.open_workbook(loc) 
    sheet = wb.sheet_by_index(0) 


    # Keeps track of the type of data file that's being used
    oldFile = False
    newFile = False

    # Global variables that will be needed for making necessary calculations
    currRow = 0
    currCol = 0
    # New Files Variables
    undergradLoans = 0
    undergradRecipients = 0
    gradRecipients = 0
    gradLoans = 0
    # Old Files Variables
    numRecipients = 0
    loanTotal = 0

    # find correct row, currCol will remain as zero
    for currRow in range(sheet.nrows):
        if sheet.cell_value(currRow, currCol) == "OPE ID":
            currRow -= 1
            break


    # find correct col
    for currCol in range(sheet.ncols):
        # Old Files Calculations
        if sheet.cell_value(currRow,currCol) == "DL UNSUBSIDIZED":
            oldFile = True
            # Finding the Recipients Row
            for currRow in range(sheet.nrows):
                if sheet.cell_value(currRow,currCol) == "Recipients":
                    numRecipients = total_Number_Of_Recipients(currRow, currCol, wb)
                    break
             # Finding the $ of Loans Originated Column
            currCol += 2
            loanTotal = total_Amount_Of_Loans(currRow, currCol, wb)
            
            print("Number of recipients is:{:,.2f}".format(numRecipients))
            print("Total Amount of loans is:${:,.2f}".format(loanTotal))
            totalAverageLoansCollection[key] = loanTotal/numRecipients
            break
        # New Files Calculations
        elif sheet.cell_value(currRow,currCol) == "DL UNSUBSIDIZED - UNDERGRADUATE" or sheet.cell_value(currRow, currCol) == "DL UNSUBSIDIZED- UNDERGRADUATE":
            newFile = True
            # Determines the loan cost and # of recipients for undergrad students
            for currRow in range(sheet.nrows):
                if sheet.cell_value(currRow,currCol) == "Recipients":
                    undergradRecipients = total_Number_Of_Recipients(currRow, currCol, wb)
                    break
            # Finding the $ of Loans Originated Column for undergrads
            currCol += 2
            undergradLoans =  total_Amount_Of_Loans(currRow, currCol, wb) 

            # Checking to make sure that we have the correct values   
            print("Number of undergrad recipients is:{:,.2f}".format(undergradRecipients))
            print("Total Amount of undergrad loans is:${:,.2f}".format(undergradLoans))

            undergradAverageLoansCollection[key] = undergradLoans/undergradRecipients

            # Determines the loan cost and # of recipients for grad students
            currRow -= 1
            for currCol in range(sheet.ncols):
                if sheet.cell_value(currRow,currCol) == "DL UNSUBSIDIZED - GRADUATE" or sheet.cell_value(currRow, currCol) == "DL UNSUBSIDIZED- GRADUATE":
                    # Finding number of grad recipients
                    for currRow in range(sheet.nrows):
                        if sheet.cell_value(currRow, currCol) == "Recipients":
                            gradRecipients = total_Number_Of_Recipients(currRow, currCol, wb)
                            break
                if gradRecipients > 0:
                    break
            # Finding total loan cost for grad recipients
            currCol += 2
            gradLoans = total_Amount_Of_Loans(currRow+1,currCol, wb)

        
            print("Number of grad recipients is:{:,.2f}".format(gradRecipients))
            print("Total Amount of grad loans is:${:,.2f}".format(gradLoans))
            gradAverageLoansCollection[key] = gradLoans/gradRecipients

    if newFile == True:
        averageLoan = (undergradLoans+gradLoans)/(undergradRecipients+gradRecipients)
    else:
        averageLoan = loanTotal/numRecipients
    print("The average loan is:${:,.2f}".format(averageLoan))
    averageLoans[key] = averageLoan

### Give the location of the file 
# 1999-2006
# location = '\Downloads\LoanData\'' 
for year in range(1999, 2006):
    location = "\Downloads\LoanData\DL_AwardYr_Summary_AY" + str(year) + "_" + str(year + 1) + "_All.xls"
    print(str(year) + "-" + str(year+1))
    operate(location, year)
    
for year in range(2006, 2018):
    for n in range(1,5):
        location = "\Downloads\LoanData\DL_Dashboard_AY" + str(year) + "_" + str(year + 1) + "_Q" + str(n) + ".xls"
        k = year + 0.1 * n
        print(str(year) + "-" + str(year+1) + " Q" + str(n))
        operate(location, k)

for year in range(2018, 2020):
    for n in range(1,5):
        if year == 2019 and n == 4:
            break
        location = "\Downloads\LoanData\dl-dashboard-ay" + str(year) + "-" + str(year + 1) + "-q" + str(n) + ".xls"
        k = year + 0.1 * n
        print(str(year) + "-" + str(year+1) + " Q" + str(n))
        operate(location, k)
print("Total average loans for all years")
for i in sorted (averageLoans) : 
    print((i, averageLoans[i]), end ="\n")

print("Total average loans for combined undergraduate and graduate years")
for i in sorted (totalAverageLoansCollection) : 
    print((i, totalAverageLoansCollection[i]), end ="\n")

print("The average loans for undergraduate years")
for i in sorted(undergradAverageLoansCollection):
    print((i, undergradAverageLoansCollection[i]), end = "\n")


print("The average loans for graduate years")
for i in sorted(gradAverageLoansCollection):
    print((i, gradAverageLoansCollection[i]), end = "\n")


'''
# Creates the excel sheet that will contain all of the data
sheet1 = wb.add_sheet('Sheet 1')
wrtRow = 0
wrtCol = 0

# Format for inserting data from a new data file
if newFile == True:
    sheet1.write(wrtRow, wrtCol, 'Year')
    sheet1.write(wrtRow, wrtCol+1, 'undergraduate Average Loan')
    sheet1.write(wrtRow, wrtCol+2, 'Graduate Average Loan')
    sheet1.write(wrtRow, wrtCol+3, 'Total Average Loan')
    wrtRow += 1
    wrtCol = 0
    sheet1.wrtie(wrtRow, wrtCol, year) # Figure out how to read year from excel file name
    sheet1.write(wrtRow, wrtCol, )
# Format for inserting data from an old data file
else:
    sheet1.write(wrtRow, wrtCol, 'Year')
    sheet1.write(wrtRow, wrtCol+1, 'Total Average Loan')
    wrtRow += 1
    wrtCol = 0

wb.save('proj2Data.xls')

# To ensure that we're using correct column and row (row, column)
print("The total number of rows for data values is:",sheet.nrows)
# print("The starting loan value is:${:,.2f}".format(sheet.cell_value(6,12)))



print("The total amount of loans is:${:,.2f}".format(numLoans))
print("The total number of recipients is:{:,.2f}".format(numRecipients))

# Finding the average loan per student for this quarter
averageLoan = numLoans/numRecipients
print()
print("The average loan for this quarter/year is:${:,.2f}".format(averageLoan))
'''