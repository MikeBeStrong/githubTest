import xlwings as xw
def getXls():
    workbook = xw.books.open('输入：制造业清单.xlsx')
    worksheet = workbook.sheets[0]
    value = worksheet.range('A2').expand('table').value

    data = dict()
    for i in range(len(value)):
        product_name = value[i][1]
        if product_name not in data:
            data[product_name] = []
        data[product_name].append(value[i])
    for key,value in data.items():
        new_workbook = xw.books.add()
        new_worksheet = new_workbook.sheets.add(key)
        new_worksheet['A1'].value = worksheet['A1:H1'].value
        new_worksheet['A2'].value = value
        new_workbook.save('F:\{}.xlsx'.format(key))
