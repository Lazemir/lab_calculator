from sympy import *
import openpyxl

wb = openpyxl.load_workbook(filename = 'vars.xlsx')
sheet = wb['Vars']

vars = []
values = []
errors = []
i = 2
while True:
    var = sheet.cell(row = i, column = 1).value
    if not var:
        break
    vars.append(var)
    values.append(sheet.cell(row = i, column = 2).value)
    errors.append(sheet.cell(row = i, column = 3).value)
    i += 1

sym = []
sum_square_errors = 0

formula = S(sheet.cell(row = 2, column = 4).value)

for i in range(len(vars)):
    sym.append(symbols(vars[i]))

for i in range(len(vars)):
    der = diff(formula, sym[i])
    err = der**2 * errors[i]**2
    sum_square_errors += err

for i in range(len(vars)):
    formula = formula.subs(sym[i], values[i])
    sum_square_errors = sum_square_errors.subs(sym[i], values[i])

error = sum_square_errors ** 0.5

print('Значение:', float(formula))
print('Абсолютная погрешность:', float(error))
print('Относительная погрешность: ', float(error / formula * 100), '%', sep = '')



