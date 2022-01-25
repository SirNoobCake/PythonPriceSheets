from openpyxl import load_workbook

def isfloat(num):
    try:
        float(num)
        return True
    except ValueError:
        return False

# Setting Year
year = 0
with open("year.txt", "r") as file_year:
    for line in file_year:
        year = line
file_year.close()


# Setting up vars for WSDP
rows_WSDP = 21
columns_WSDP = 19

products_WSDP_arr = [[0 for i in range(rows_WSDP)] for j in range(columns_WSDP)]

j = 0
k = -1 # Temp fix to make k == 0 for first iteration

# Populating products_arr
with open("WSDP.txt", "r") as WSPD:
    for line in WSPD:
        for i in line.split(','):
            if isfloat(i.strip()):
                products_WSDP_arr[k][j] = float(i)
                if k == 1 & j == 15 | k > 12 & j == 15:
                    continue
                j += 1
            else:
                k += 1
                j = 0
WSPD.close()

# Setting up vars for WSP
rows_WSP = 20
columns_WSP = 34

products_WSP_list = [[0 for i in range(rows_WSP)] for j in range(columns_WSP)]

#   Populating products_WSP_list
j = 0
k = -1

with open("WSP.txt", "r") as WSP:
    for line in WSP:
        for i in line.split(','):
            if isfloat(i.strip()):
                products_WSP_list[k][j] = float(i)
                if k == 1 & j == 15 | k > 12 & j == 15:
                    continue
                j += 1
            else:
                k += 1
                j = 0
WSP.close()
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
#  ****** Start of Changes in Yardages ****** 17
#    13 = Rain Garden Mix
#    14 = Screened Blend
#    15 = Screened Topsoil
#    16 = Regular Topsoil
#    17 = Fill Dirt
#    18 = Topsoil Overs

#   Setting up for Wholesale Prices Sheet
load_WSP_wb = load_workbook("Wholesale PriceSheet Template.xlsx")
ws_WSP = load_WSP_wb.active

#   Set Year in sheet
ws_WSP['S1'] = year + " Wholesale Prices"

#   Setting Prices
x = 0
for i in range(columns_WSP):
    j = 0
    if i < 17:
        for row in ws_WSP.iter_rows(min_col=3,min_row=x + 5,max_col=6,max_row=x + 5):
            for cell in row:
                cell.value = format(products_WSP_list[i][j],'.2f') #    error?
                j += 1
                if j > 4:
                    break
    else:
        if i == 17:
            x = 0
        for row in ws_WSP.iter_rows(min_col=8,min_row=x + 5,max_col=27,max_row=x + 5):
            for cell in row:
                cell.value = format(products_WSP_list[i][j],'.2f')
                j += 1
                if i > 28:
                    if j > 19:
                        break
    if x == 3 or x == 5 or x == 12:
        x += 2
    else:
        x += 1
load_WSP_wb.save(year + "_Wholesale_Prices.xlsx")

#   Setting up for Wholesale Delivered Prices Sheet

#   grab the active worksheet
load_WSDP_wb = load_workbook("Wholesale Delivered Template.xlsx")
ws_WSDP = load_WSDP_wb.active

#   Set year in Excel
ws_WSDP['R2'] = year + " Wholesale Delivered Prices"

#   Setting Prices
x = 6
for i in range(columns_WSDP):
    j = 0
    if i == 1:
        for row in ws_WSDP.iter_rows(min_col=5,min_row=29,max_col=25,max_row=29):
            for cell in row:
                cell.value = format(products_WSDP_arr[i][j],'.2f')
                j += 1
                if j > 16:
                    break
        x = 6
    else:
        for row in ws_WSDP.iter_rows(min_col=5,min_row=x,max_col=25,max_row=x):
            for cell in row:
                cell.value = format(products_WSDP_arr[i][j],'.2f')
                j += 1
                if i > 12:
                    if j > 16:
                        break

    # This Makes the Program only interact with the desired collums
    if x == 6:
        x += 1
    else:
        if x == 27:
            x += 3
        else:
            x += 2

# Save Worksheet
load_WSDP_wb.save(year + "_Wholesale_Delivered_Prices.xlsx")

