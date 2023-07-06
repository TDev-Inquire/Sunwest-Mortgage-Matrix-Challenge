import os
import time
import sys
import xlsxwriter
from re import A

#xx 01 02 03
#04 05 06 07 c4
#08 09 10 11 c5
#12 13 14 15 c6
#xx c1 c2 c3

#Granted I could create this a lot more simplified but I enjoyed being able to see it broken down

def clear_screen():
    if os.name == 'nt':
        _ = os.system('cls')
    else:
        _ = os.system('clear')
        
def changeValues():
    print(f"{'' : <10}{'col1' : ^10}{'col2' : ^10}{'col3' : ^10}")
    print(f"{'row1' : <10}{'x1' : ^10}{'x2' : ^10}{'x3' : ^10}")
    print(f"{'row2' : <10}{'y1' : ^10}{'y2' : ^10}{'y3': ^10}")
    print(f"{'row3' : <10}{'z1' : ^10}{'z2' : ^10}{'z3' : ^10}")
    print(f"{'Calculate' : <10}{'result1' : ^10}{'result2' : ^10}{'result3' : ^10}")
    print("===================================================")
    print("-Which of the fields above would you like to change? (Enter the ID)")
    print("-For results, value would be, -/+/*, or Subtract, add, multiply")
    print("-For col(umns) or row(s), the value can be letters/number")
    print("-For xyz values use digits/numbers")
    
    unitID = input("Which field would you like to change?")
        
def export_to_excel():
    # Define the values
    col1 = ''
    col2 = 'Col2'
    col3 = 'Col3'
    row1 = 'Row1'
    x1 = 1
    x2 = 2
    x3 = 3
    row2 = 'Row2'
    y1 = 4
    y2 = 5
    y3 = 6
    row3 = 'Row3'
    z1 = 7
    z2 = 8
    z3 = 9
    result1 = 'Result1'
    result2 = 'Result2'
    result3 = 'Result3'

    # Create a new workbook and add a worksheet
    workbook = xlsxwriter.Workbook('output.xlsx')
    worksheet = workbook.add_worksheet()

    # Set the formatting for the headers
    header_format = workbook.add_format({'bold': True, 'align': 'center'})
    worksheet.set_column('A:D', 10, None)

    # Write the headers to the spreadsheet
    worksheet.write('B1', col1, header_format)
    worksheet.write('C1', col2, header_format)
    worksheet.write('D1', col3, header_format)

    # Write the values to the spreadsheet
    worksheet.write('A2', row1, header_format)
    worksheet.write('B2', x1)
    worksheet.write('C2', x2)
    worksheet.write('D2', x3)

    worksheet.write('A3', row2, header_format)
    worksheet.write('B3', y1)
    worksheet.write('C3', y2)
    worksheet.write('D3', y3)

    worksheet.write('A4', row3, header_format)
    worksheet.write('B4', z1)
    worksheet.write('C4', z2)
    worksheet.write('D4', z3)

    worksheet.write('A5', 'Calculate', header_format)
    worksheet.write('B5', result1)
    worksheet.write('C5', result2)
    worksheet.write('D5', result3)

    # Close the workbook
    workbook.close()

clear_screen()

print('=====Time for a table=====')

#COLUMNS ==================================================================================================================================

print(f"{'' : <10}{'column 1' : ^10}{'column 2' : ^10}{'column 3' : ^10}")
print(f"{'row 1' : <10}{'0' : ^10}{'0' : ^10}{'0' : ^10}")
print(f"{'row 2' : <10}{'0' : ^10}{'0' : ^10}{'0' : ^10}")
print(f"{'row 3' : <10}{'0' : ^10}{'0' : ^10}{'0' : ^10}")
print(f"{'Calculate' : <10}{'-' : ^10}{'-' : ^10}{'-' : ^10}")

col1 = input('What would you like the first column title to be?: ')

clear_screen()

print(f"{'' : <10}{col1 : ^10}{'column 2' : ^10}{'column 3' : ^10}")
print(f"{'row 1' : <10}{'0' : ^10}{'0' : ^10}{'0' : ^10}")
print(f"{'row 2' : <10}{'0' : ^10}{'0' : ^10}{'0' : ^10}")
print(f"{'row 3' : <10}{'0' : ^10}{'0' : ^10}{'0' : ^10}")
print(f"{'Calculate' : <10}{'-' : ^10}{'-' : ^10}{'-' : ^10}")


col2 = input('What would you like the second column title to be?: ')

clear_screen()

print(f"{'' : <10}{col1 : ^10}{col2 : ^10}{'column 3' : ^10}")
print(f"{'row 1' : <10}{'0' : ^10}{'0' : ^10}{'0' : ^10}")
print(f"{'row 2' : <10}{'0' : ^10}{'0' : ^10}{'0' : ^10}")
print(f"{'row 3' : <10}{'0' : ^10}{'0' : ^10}{'0' : ^10}")
print(f"{'Calculate' : <10}{'-' : ^10}{'-' : ^10}{'-' : ^10}")


col3 = input('What would you like the third column title to be?: ')

clear_screen()

#ROWS ====================================================================================================================================

print(f"{'' : <10}{col1 : ^10}{col2 : ^10}{col3 : ^10}")
print(f"{'row 1' : <10}{'0' : ^10}{'0' : ^10}{'0' : ^10}")
print(f"{'row 2' : <10}{'0' : ^10}{'0' : ^10}{'0' : ^10}")
print(f"{'row 3' : <10}{'0' : ^10}{'0' : ^10}{'0' : ^10}")
print(f"{'Calculate' : <10}{'-' : ^10}{'-' : ^10}{'-' : ^10}")


row1 = input('What would you like the second row to be?: ')

clear_screen()

print(f"{'' : <10}{col1 : ^10}{col2 : ^10}{col3 : ^10}")
print(f"{row1 : <10}{'0' : ^10}{'0' : ^10}{'0' : ^10}")
print(f"{'row 2' : <10}{'0' : ^10}{'0' : ^10}{'0' : ^10}")
print(f"{'row 3' : <10}{'0' : ^10}{'0' : ^10}{'0' : ^10}")
print(f"{'Calculate' : <10}{'-' : ^10}{'-' : ^10}{'-' : ^10}")


row2 = input('What would you like the second row to be?: ')

clear_screen()

print(f"{'' : <10}{col1 : ^10}{col2 : ^10}{col3 : ^10}")
print(f"{row1 : <10}{'0' : ^10}{'0' : ^10}{'0' : ^10}")
print(f"{row2 : <10}{'0' : ^10}{'0' : ^10}{'0' : ^10}")
print(f"{'row 3' : <10}{'0' : ^10}{'0' : ^10}{'0' : ^10}")
print(f"{'Calculate' : <10}{'-' : ^10}{'-' : ^10}{'-' : ^10}")


row3 = input('What would you like the third row to be?: ')

clear_screen()

#Values ====================================================================================================================================

print(f"{'' : <10}{col1 : ^10}{col2 : ^10}{col3 : ^10}")
print(f"{row1 : <10}{'x' : ^10}{'0' : ^10}{'0' : ^10}")
print(f"{row2 : <10}{'0' : ^10}{'0' : ^10}{'0' : ^10}")
print(f"{row3 : <10}{'0' : ^10}{'0' : ^10}{'0' : ^10}")
print(f"{'Calculate' : <10}{'-' : ^10}{'-' : ^10}{'-' : ^10}")

x1 = int(input('What is the value?: '))

clear_screen()

print(f"{'' : <10}{col1 : ^10}{col2 : ^10}{col3 : ^10}")
print(f"{row1 : <10}{x1 : ^10}{'x' : ^10}{'0' : ^10}")
print(f"{row2 : <10}{'0' : ^10}{'0' : ^10}{'0' : ^10}")
print(f"{row3 : <10}{'0' : ^10}{'0' : ^10}{'0' : ^10}")
print(f"{'Calculate' : <10}{'-' : ^10}{'-' : ^10}{'-' : ^10}")

x2 = int(input('What is the value?: '))

clear_screen()

print(f"{'' : <10}{col1 : ^10}{col2 : ^10}{col3 : ^10}")
print(f"{row1 : <10}{x1 : ^10}{x2 : ^10}{'x' : ^10}")
print(f"{row2 : <10}{'0' : ^10}{'0' : ^10}{'0' : ^10}")
print(f"{row3 : <10}{'0' : ^10}{'0' : ^10}{'0' : ^10}")
print(f"{'Calculate' : <10}{'-' : ^10}{'-' : ^10}{'-' : ^10}")

x3 = int(input('What is the value?: '))

clear_screen()

#y Values ====================================================================================================================================

print(f"{'' : <10}{col1 : ^10}{col2 : ^10}{col3 : ^10}")
print(f"{row1 : <10}{x1 : ^10}{x2 : ^10}{x3 : ^10}")
print(f"{row2 : <10}{'x' : ^10}{'0' : ^10}{'0' : ^10}")
print(f"{row3 : <10}{'0' : ^10}{'0' : ^10}{'0' : ^10}")
print(f"{'Calculate' : <10}{'-' : ^10}{'-' : ^10}{'-' : ^10}")

y1 = int(input('What is the value?: '))

clear_screen()

print(f"{'' : <10}{col1 : ^10}{col2 : ^10}{col3 : ^10}")
print(f"{row1 : <10}{x1 : ^10}{x2 : ^10}{x3 : ^10}")
print(f"{row2 : <10}{y1 : ^10}{'x' : ^10}{'0' : ^10}")
print(f"{row3 : <10}{'0' : ^10}{'0' : ^10}{'0' : ^10}")
print(f"{'Calculate' : <10}{'-' : ^10}{'-' : ^10}{'-' : ^10}")


y2 = int(input('What is the value?: '))

clear_screen()

print(f"{'' : <10}{col1 : ^10}{col2 : ^10}{col3 : ^10}")
print(f"{row1 : <10}{x1 : ^10}{x2 : ^10}{x3 : ^10}")
print(f"{row2 : <10}{y1 : ^10}{y2 : ^10}{'x' : ^10}")
print(f"{row3 : <10}{'0' : ^10}{'0' : ^10}{'0' : ^10}")
print(f"{'Calculate' : <10}{'-' : ^10}{'-' : ^10}{'-' : ^10}")


y3 = int(input('What is the value?: '))

clear_screen()


#z Values ====================================================================================================================================
# add xy + y1, or subtract. 
print(f"{'' : <10}{col1 : ^10}{col2 : ^10}{col3 : ^10}")
print(f"{row1 : <10}{x1 : ^10}{x2 : ^10}{x3 : ^10}")
print(f"{row2 : <10}{y1 : ^10}{y2 : ^10}{y3 : ^10}")
print(f"{row3 : <10}{'x' : ^10}{'0' : ^10}{'0' : ^10}")
print(f"{'Calculate' : <10}{'-' : ^10}{'-' : ^10}{'-' : ^10}")

z1 = int(input('What is the value?: '))

clear_screen()

print(f"{'' : <10}{col1 : ^10}{col2 : ^10}{col3 : ^10}")
print(f"{row1 : <10}{x1 : ^10}{x2 : ^10}{x3 : ^10}")
print(f"{row2 : <10}{y1 : ^10}{y2 : ^10}{y3 : ^10}")
print(f"{row3 : <10}{z1 : ^10}{'x' : ^10}{'0' : ^10}")
print(f"{'Calculate' : <10}{'-' : ^10}{'-' : ^10}{'-' : ^10}")

z2 = int(input('What is the value?: '))

clear_screen()

print(f"{'' : <10}{col1 : ^10}{col2 : ^10}{col3 : ^10}")
print(f"{row1 : <10}{x1 : ^10}{x2 : ^10}{x3 : ^10}")
print(f"{row2 : <10}{y1 : ^10}{y2 : ^10}{y3 : ^10}")
print(f"{row3 : <10}{z1 : ^10}{z2 : ^10}{'x' : ^10}")
print(f"{'Calculate' : <10}{'-' : ^10}{'-' : ^10}{'-' : ^10}")

z3 = int(input('What is the value?: '))

clear_screen()

#function for calculation ===================================================================================================================

print(f"{'' : <10}{col1 : ^10}{col2 : ^10}{col3 : ^10}")
print(f"{row1 : <10}{x1 : ^10}{x2 : ^10}{x3 : ^10}")
print(f"{row2 : <10}{y1 : ^10}{y2 : ^10}{y3 : ^10}")
print(f"{row3 : <10}{z1 : ^10}{z2 : ^10}{z3 : ^10}")
print(f"{'Calculate' : <10}{'+/-/x' : ^10}{'-' : ^10}{'-' : ^10}")

operation1 = input('Add, Subtract, multiply columns?: ')

if operation1 == "multiply" or operation1 == "*" or operation1 == "X" or operation1 == "x":
    result1 = x1 * y1 * z1
elif operation1 == "subtract" or operation1 == "-":
    result1 = x1 - y1 - z1
elif operation1 == "add" or operation1 == "+":
    result1 = x1 + y1 + z1
else:
    result1 = "Invalid operation"

    


print(f"{'' : <10}{col1 : ^10}{col2 : ^10}{col3 : ^10}")
print(f"{row1 : <10}{x1 : ^10}{x2 : ^10}{x3 : ^10}")
print(f"{row2 : <10}{y1 : ^10}{y2 : ^10}{y3 : ^10}")
print(f"{row3 : <10}{z1 : ^10}{z2 : ^10}{z3 : ^10}")
print(f"{'Calculate' : <10}{result1 : ^10}{'+/-/x' : ^10}{'-' : ^10}")

operation2 = input('Add, Subtract, multiply columns?')

if operation2 == "multiply" or operation2 == "*" or operation2 == "X" or operation2 == "x":
    result2 = x2 * y2 * z2
elif operation2 == "subtract" or operation2 == "-":
    result2 = x2 - y2 - z2
elif operation2 == "add" or operation2 == "+":
    result2 = x2 + y2 + z2
else:
    result2 = "Invalid operation"
    


print(f"{'' : <10}{col1 : ^10}{col2 : ^10}{col3 : ^10}")
print(f"{row1 : <10}{x1 : ^10}{x2 : ^10}{x3 : ^10}")
print(f"{row2 : <10}{y1 : ^10}{y2 : ^10}{y3 : ^10}")
print(f"{row3 : <10}{z1 : ^10}{z2 : ^10}{z3 : ^10}")
print(f"{'Calculate' : <10}{result1 : ^10}{result2 : ^10}{'+/-/x' : ^10}")

operation3 = input('Add, Subtract, multiply columns?: ')

if operation3 == "multiply" or operation3 == "*" or operation3 == "X" or operation3 == "x":
    result3 = x3 * y3 * y3
elif operation3 == "subtract" or operation3 == "-":
    result3 = x3 - y3 - y3
elif operation3 == "add" or operation3 == "+":
    result3 = x3 + y3 + y3
else:
    result3 = "Invalid operation"
    


clear_screen()

    
clear_screen()
    
#final results ==================================================================================================================================


print(f"{'' : <10}{col1 : ^10}{col2 : ^10}{col3 : ^10}")
print(f"{row1 : <10}{x1 : ^10}{x2 : ^10}{x3 : ^10}")
print(f"{row2 : <10}{y1 : ^10}{y2 : ^10}{y3 : ^10}")
print(f"{row3 : <10}{z1 : ^10}{z2 : ^10}{z3 : ^10}")
print(f"{'Calculate' : <10}{result1 : ^10}{result2 : ^10}{result3 : ^10}")

def end():
    print('Task completed, would you like to:')
    print('-End')
    print('-Update (Change Values) (coming soon!)')
    print('-Export (to a spreadsheet)')

end()
choice = input('Answer...: ')



if choice == "End" or choice == "end":
    sys.exit()
elif choice == "Export" or choice == "export":
    export_to_excel()
elif choice == "Update" or choice == "update":
    change_variable = input("Enter the variable you want to change (col1, col2, col3, row1, row2, row3, x1, x2, x3, y1, y2, y3, z1, z2, z3, result1, result2, result3): ")
    new_value = input(f"Enter the new value for {change_variable}: ")
    exec(f"{change_variable} = new_value")
    end()
































































