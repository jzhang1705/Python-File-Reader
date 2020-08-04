# REMINDER: Try to find a function to auto-adjust the size of the graph and try to merge this with newDataReader

# Reading an excel file using Python 
import xlrd, matplotlib.pyplot as plt

# Give the location of the file
location = "\Downloads\FinalCoding\dataStuff.xls"

# To open Workbook 
wb = xlrd.open_workbook(location) 
sheet = wb.sheet_by_index(0) 
years = []
currRow = 0
currCol = 0
for year in range(1, sheet.nrows):
    years.append(sheet.cell_value(year, currCol))

print("YEARS ", years)

print()

plt.xticks(years)


currCol = 1
loans = []
for loan in range(1, sheet.nrows):
    loans.append(sheet.cell_value(loan, currCol))

print("LOANS ", loans) 


gradLoans = []
for i in range(1, 10):
    print()
print("College Graduate Loans: ")

currRow = 1
currCol = 1
for i in range(1999, 2031):
    gradLoans.append(sheet.cell_value(currRow, currCol))
    currRow += 1
print(gradLoans)
print()
print()

gradSalaries = []
currRow = 1
currCol = 4
for currRow in range(currRow, sheet.nrows):
    gradSalaries.append(sheet.cell_value(currRow, currCol))


print("Graduating Salaries: ", gradSalaries)
print()

averageLoanTotals = []
for i in range(0, len(gradLoans)-4):
    averageLoanTotal = 0
    stoppingValue = i+4
    for z in range(i, stoppingValue):
        averageLoanTotal += gradLoans[z]
    averageLoanTotals.append(averageLoanTotal)

print("Average Loan Totals (Over 4 Years): ", averageLoanTotals)


print("College Graduate Salaries: ")
for i in range(2020, 2031):
    gradSalary= 670.46*i - 1301441.91
    gradSalaries.append(gradSalary)

graduateYears = []
for i in range(2003, 2031):
    graduateYears.append(i)

ratios = []
print("Length of graduate salaries: ")
print(len(gradSalaries))
print("Length of average loan totals: ")
print(len(averageLoanTotals))
for i in range(0, 28):
    ratio = (100/gradSalaries[i+4])*averageLoanTotals[i]
    ratios.append(ratio)
print("Ratios: ")
print(ratios)

plt.plot(years, gradLoans, color='red', linestyle='dashed', linewidth = 3, 
         marker='o', markerfacecolor='red', markersize=12) 

# naming the x-axis
plt.xlabel('Years')
# naming the y-axis
plt.ylabel('Average College Debt ($)')

# giving a title to the graph
plt.title('Student Loans Trend')

# function to show the plot
plt.xticks(years)
plt.savefig("image5.png",bbox_inches='tight',dpi=100)
plt.show()


