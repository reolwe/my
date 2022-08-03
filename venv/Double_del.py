import openpyxl as pxl
wbook = pxl.load_workbook('C:/Users/m.shcherbina/PycharmProjects/ITOIUItransfer/shtut_input.xlsx')
wSbook = wbook.active
for i in range(1, wSdict.max_row + 1):