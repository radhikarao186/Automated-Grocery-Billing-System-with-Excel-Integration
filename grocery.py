import sys
import math
import xlwt
import xlrd
import pandas as pd

# Update these variables to point to the grocery files
orderfile = "grocery orders.xlsx"
pricefile = "product list new.xlsx"
billsfolder = "./bills/"

def main():
    priceDict = fetch_price_list()
    orders = fetch_orders(priceDict)
    frame_bills(orders)

def frame_bills(orders):
    fileslist = []
    wbcombined = xlwt.Workbook()
    sheetcombined = wbcombined.add_sheet('Bill')
    for name, value in orders.items():
        filename = frameIndividualBill(name, value)
        fileslist.append(filename)
    
    outrow_idx = 0
    for f in fileslist:
        insheet = xlrd.open_workbook(f).sheets()[0]
        for row_idx in range(1, insheet.nrows):
            for col_idx in range(insheet.ncols):
                sheetcombined.write(outrow_idx, col_idx, insheet.cell_value(row_idx, col_idx))
            outrow_idx += 1
        outrow_idx += 2
    combined_filename = billsfolder + 'combined.xls'
    wbcombined.save(combined_filename)

def frameIndividualBill(name, value):
    wb = xlwt.Workbook()
    sheet = wb.add_sheet('Bill')
    total_amount = 0  # Initialize total amount for the bill

    for x in range(len(value)):
        key = list(value)[x]
        if key == 'Timestamp':
            date_style = xlwt.easyxf(num_format_str='YYYY-MM-DD HH-M-SS')
            keyval = value[key]
            sheet.write(x, 0, key)
            sheet.write(x, 1, keyval[0], date_style)
            sheet.write(x, 2, keyval[1])
            sheet.write(x, 3, keyval[2])
        else:
            keyval = value[key]
            sheet.write(x, 0, key)
            sheet.write(x, 1, keyval[0])
            sheet.write(x, 2, keyval[1])
            sheet.write(x, 3, keyval[2])

            # Ensure that `keyval[2]` is a number before adding it to `total_amount`
            if isinstance(keyval[2], (int, float)):
                total_amount += keyval[2]

    # Write the total amount at the end of the bill
    sheet.write(len(value), 0, 'Total Amount')
    sheet.write(len(value), 3, total_amount)

    filename = name + '.xls'
    file = billsfolder + filename
    wb.save(file)
    return file

def fetch_orders(priceDict):
    orderDict = {}
    ordersBook = pd.read_excel(orderfile)
    ordersBook.dropna(how='all', inplace=True)
    for row in ordersBook.itertuples(index=False):
        rowDict = {}
        for col in ordersBook.columns:
            colloc = ordersBook.columns.get_loc(col)
            if col in ['Timestamp', 'NAME', 'ADDRESS', 'AREA', 'MOBILE NO.', 'FLAT NO']:
                # Treat `FLAT NO` as a string
                rowDict[col] = [str(row[colloc]), '', ''] if pd.notna(row[colloc]) else ['', '', '']
            else:
                if pd.notna(row[colloc]):
                    proName = col
                    qty = int(row[colloc])  # Quantity should be an integer
                    price = int(priceDict.get(col.strip(), 0))
                    total = qty * price
                    valList = [qty, price, total]
                    rowDict[col] = valList
        orderDict[rowDict['NAME'][0]] = rowDict
    return orderDict

def fetch_price_list():
    priceDict = {}
    priceBook = pd.read_excel(pricefile, skiprows=3)  # Adjusted based on your product list
    for row in priceBook.itertuples(index=False):
        if pd.notna(row[1]):
            product_name = str(row[1]).strip()  # Assuming product names are in the second column
            price = row[4]  # Assuming MRP is in the third column
            priceDict[product_name] = price
    return priceDict

if __name__ == '__main__':
    main()
