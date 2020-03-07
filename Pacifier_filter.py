import xlrd
import xlwt

data_path = "D:\MCM2020\pacifier.xlsx"

f = xlwt.Workbook()
wb = xlrd.open_workbook(filename=data_path)
sheet1 = wb.sheet_by_index(0)  # input
sheet2 = f.add_sheet('page1', cell_overwrite_ok=True)  # for output
title_col = sheet1.col_values(5)

candidates = ['pacifier', 'dummy', 'binky', 'soother', 'teether', 'dodie']
j = 0
for i in range(0, len(title_col)):
    title_col[i] = title_col[i].replace(',', ' ')
    title_col[i] = title_col[i].replace('.', ' ')
    words = title_col[i].split(' ')
    for word in words:
        if word.lower() in candidates:
            for k in range(0, 16):
                sheet2.write(j, k, sheet1.row_values(i)[k])
            j = j + 1
            continue
    if not i%100:
        print(i)

f.save('D:\\MCM2020\\pacifier_filtered.xls')

