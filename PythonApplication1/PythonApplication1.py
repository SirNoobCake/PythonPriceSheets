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

rows = 21
columns = 19

products_arr = [[0 for i in range(rows)] for j in range(columns)]

j = 0
k = -1 # Temp fix to make k == 0 for first iteration

with open("TestPB.txt", "r") as file1:
    for line in file1:
        for i in line.split(','):
            if isfloat(i.strip()):
                products_arr[k][j] = float(i)
                if k == 1 & j == 15 | k > 12 & j == 15:
                    continue
                j += 1
            else:
                k += 1
                j = 0

# k = product, 
#    0 = Delivery Price Mulch Poducts
#    1 = Delivery Price Soil Products
#    2 = Premium Bark
#    3 = Bark Blend
#    4 = Nature's Blend
#    5 = Beauty Bark
#    6 = Dyed Muclhes
#    7 = Safe Cover
#    8 = Clean Wood Chips
#    9 = Wood Chips
#    10 = Compost
#    11 = Leaf Compost
#    12 = Mushroom Soil
#  ****** Start of Changes in Yardages ****** 15
#    13 = Rain Garden Mix
#    14 = Screened Blend
#    15 = Screened Topsoil
#    16 = Regular Topsoil
#    17 = Fill Dirt
#    18 = Topsoil Overs
file1.close()


# Premium Bark
x = 0
for row in ws.iter_rows(min_col=5,min_row=7, max_col=25, max_row=7):
    for cell in row:
#        cell.value = format(lines[x], '.2f')
#        print(cell.value)
        x += 1

x = 6
for i in range(columns):
    j = 0
    if i == 1:
        for row in ws.iter_rows(min_col=5,min_row=29,max_col=25,max_row=29):
            for cell in row:
                cell.value = format(products_arr[i][j],'.2f')
                j += 1
                if j > 15:
                    continue
        x = 6
    else:
        for row in ws.iter_rows(min_col=5,min_row=x,max_col=25,max_row=x):
            for cell in row:
                cell.value = format(products_arr[i][j],'.2f')
                j += 1
                if i > 12:
                    if j > 15:
                        continue

    # This Makes the Program only interact with the desired collums
    if x == 6:
        x += 1
    else:
        if x == 27:
            x += 3
        else:
            x += 2
# k = 6
# if (j == 0) || (j == 23) k ++; else k + 2;
#
#6  Delivered Price Mulch
#7  Premium bark
#8  
#9  Bark Blend
#10 
#11 Nature's Blend
#12 
#13 Beauty Bark
#14 
#15 Dyed Mulches
#16 
#17 Safe Cover
#18 
#19 Clean Wood Chips
#20 
#21 Wood Chips
#22 
#23 Compost
#24 
#25 Leaf Compost
#26 
#27 Mushroom Soil
#28 
#29 Delivered Price Soil
#30 Rain Garden Mix
#31 
#32 Screened Blend
#33 
#34 Screened Topsoil
#35 
#36 Regular Topsoil
#37 
#38 Fill Dirt
#39 
#40 Topsoil Overs
# Save the file
load_wb.save("sample.xlsx")