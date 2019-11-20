from sympy import *
import openpyxl

wb = openpyxl.load_workbook(filename = 'vars.xlsx')
sheet = wb['Vars']

var_row_index = 2
quant_of_calc = 0

while sheet.cell(row = var_row_index, column = 1).value:
    vars = []
    values = []
    errors = []

    if sheet.cell(row = var_row_index, column = 1).value[:3].lower() == 'new':
        res_name = sheet.cell(row = var_row_index, column = 1).value[4:]
        quant_of_calc += 1
        formula_row_index = var_row_index
        var_row_index += 1
        continue

    while sheet.cell(row = var_row_index, column = 1).value and sheet.cell(row = var_row_index, column = 1).value[:3].lower() != 'new':
        var = sheet.cell(row = var_row_index, column = 1).value
        vars.append(var)
        values.append(sheet.cell(row = var_row_index, column = 2).value)
        errors.append(sheet.cell(row = var_row_index, column = 3).value)
        var_row_index += 1

    sym = []
    sum_square_errors = 0

    if sheet.cell(row = formula_row_index, column = 4).value:
        formula = S(sheet.cell(row = formula_row_index, column = 4).value)
    else:
        formula = S(sheet.cell(row=2, column=5).value)

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

    sheet.cell(row= quant_of_calc + 1, column=7).value = res_name
    sheet.cell(row= quant_of_calc + 1, column=8).value = float(formula)
    sheet.cell(row= quant_of_calc + 1, column=9).value = float(error)
    sheet.cell(row= quant_of_calc + 1, column=10).value = float(error / formula)

    wb.save('vars.xlsx')

    print('Название:', res_name)
    print('Значение:', float(formula))
    print('Абсолютная погрешность:', float(error))
    print('Относительная погрешность: ', float(error / formula * 100), '%')
    print()



