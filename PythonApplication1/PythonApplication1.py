from openpyxl import load_workbook

def isfloat(num):
    try:
        float(num)
        return True
    except ValueError:
        return False

load_wb = load_workbook("2022 Wholesale Delivered Prices.xlsx")

# grab the active worksheet
ws = load_wb.active

rows = 19
colums = 21

products_arr = [[0] * rows] * colums

with open("TestPB.txt", "r") as file1:
    for k in range(rows):
        
            for j in range(colums):
                for line in file1:
                    for i in line.split(','):
                        if isfloat(i.strip()):
                            products_arr[k][i]
                #if (isfloat(line.split(','))):
                  #  k += 1
                 #   continue;
                #else:
                   # proproducts_arr[k][j] = [float(i) for line in file1 for i in line.split(',') if i.strip()]

file1.close()


# Premium Bark
x = 0
for row in ws.iter_rows(min_col=5,min_row=7, max_col=25, max_row=7):
    for cell in row:
#        cell.value = format(lines[x], '.2f')
 #       print(cell.value)
        x += 1

# Save the file
load_wb.save("sample.xlsx")