import xlrd, datetime
import xlsxwriter

def parseDate(floatDate, mode):
    date = datetime.datetime(*xlrd.xldate_as_tuple(floatDate, mode))
    d = dict()
    d['month'] = (int)(date.month)
    d['year'] = (int)(date.year)
    d['quarter'] = (int)((date.month-1)/3)
    if d['quarter'] == 0 :
        d['quarter'] = 4
    return d

def addHeaders(sheet, formatMode, mode):
    sheet.write('A1', 'Year', formatMode)
    sheet.set_row(0, 50)
    numCol = 1
    if mode.lower() == "quarter" or mode.lower() == "month":
        sheet.write('B1', 'Quarter', formatMode)
        numCol = 2
    if mode == "month":
        sheet.write('C1', 'Month', formatMode)
        numCol = 3
    for i in range(numCol, numCol+5):
        sheet.write(0, i, '(PADD '+(str)(i-numCol+1)+') Refinery and Blender Net Input of Crude Oil', formatMode)
    sheet.write(0, numCol+5, 'Total US Refinery Net Input of Crude Oil', formatMode)
    sheet.set_column(numCol, numCol+5, 20)

def addSheetForD(book, datasheet, paddCols, jan16Row, datemode):
    sheet = book.add_worksheet('Monthly Table (D)')
    bold = book.add_format({'bold': 1, 'text_wrap': True})
    addHeaders(sheet, bold, "month")
    newRow = 1
    for i in range (jan16Row, datasheet.nrows):
        date = parseDate(datasheet.cell_value(rowx = i, colx = 0), datemode)
        sheet.write(newRow, 0, date['year'])
        sheet.write(newRow, 1, date['quarter'])
        sheet.write(newRow, 2, date['month'])
        colVal = 3
        sum = 0
        for j in paddCols:
            sheet.write(newRow, colVal, datasheet.cell_value(rowx = i, colx= j))
            sum = sum + datasheet.cell_value(rowx = i, colx= j)
            colVal = colVal + 1
        sheet.write(newRow, colVal, sum)
        newRow = newRow + 1


def addSheetForE(book, datasheet, paddCols, jan16Row, datemode):
    sheet = book.add_worksheet('Quarterly Data (E)')
    bold = book.add_format({'bold': 1, 'text_wrap': True})
    addHeaders(sheet, bold, "quarter")
    newRow = 1
    paddValues = [0,0,0,0,0,0]
    for i in range (jan16Row, datasheet.nrows):
        sum = 0
        for j in range(0,5):
            paddValues[j] = paddValues[j] + datasheet.cell_value(rowx = i, colx = paddCols[j])
            sum = sum + datasheet.cell_value(rowx = i, colx= paddCols[j])
        paddValues[5] = paddValues[5] + sum
        
        if (i-jan16Row)%3 == 2:
            date = parseDate(datasheet.cell_value(rowx = i, colx = 0), datemode)
            sheet.write(newRow, 0, date['year'])
            sheet.write(newRow, 1, date['quarter'])
            for j in range(2,8):
                sheet.write(newRow, j, paddValues[j-2])
            newRow = newRow + 1
            paddValues = [0,0,0,0,0,0]
    if (datasheet.nrows-1)%3 !=2 :
        date = parseDate(datasheet.cell_value(rowx = datasheet.nrows-1, colx = 0), datemode)
        sheet.write(newRow, 0, date['year'])
        sheet.write(newRow, 1, date['quarter'])
        for j in range(2,8):
            sheet.write(newRow, j, paddValues[j-2])

def addSheetForF(book, datasheet, paddCols, jan16Row, datemode):
    sheet = book.add_worksheet('Yearly Data (F)')
    bold = book.add_format({'bold': 1, 'text_wrap': True})
    addHeaders(sheet, bold, "year")
    newRow = 1
    paddValues = [0,0,0,0,0,0]
    for i in range (jan16Row, datasheet.nrows):
        sum = 0
        for j in range(0,5):
            paddValues[j] = paddValues[j] + datasheet.cell_value(rowx = i, colx = paddCols[j])
            sum = sum + datasheet.cell_value(rowx = i, colx= paddCols[j])
        paddValues[5] = paddValues[5] + sum
        
        if (i-jan16Row)%12 == 11:
            date = parseDate(datasheet.cell_value(rowx = i, colx = 0), datemode)
            sheet.write(newRow, 0, date['year'])
            for j in range(1,7):
                sheet.write(newRow, j, paddValues[j-1])
            newRow = newRow + 1
            paddValues = [0,0,0,0,0,0]
    if (datasheet.nrows-1)%12 != 11 :
        date = parseDate(datasheet.cell_value(rowx = datasheet.nrows-1, colx = 0), datemode)
        sheet.write(newRow, 0, date['year'])
        for j in range(1,7):
            sheet.write(newRow, j, paddValues[j-1])


# main code

book = xlrd.open_workbook("data.xls")
datasheet = book.sheet_by_index(1)

processedBook = xlsxwriter.Workbook('processedData.xlsx')
wrapText = processedBook.add_format({'text_wrap': True})

# Checking from which row is January Data beginning
jan16Row = 0
for i in range (3, datasheet.nrows):
    date = parseDate(datasheet.cell_value(rowx = i, colx = 0), book.datemode)
    if date['year'] > 2015:
        jan16Row = i
        break

# Checking whick columns have PADD info
paddCols = []
for ncol in range (1,datasheet.ncols) :
    if "PADD" in datasheet.cell_value(rowx=2, colx=ncol):
        paddCols.append(ncol)

addSheetForD(processedBook, datasheet, paddCols, jan16Row, book.datemode)
addSheetForE(processedBook, datasheet, paddCols, jan16Row, book.datemode)
addSheetForF(processedBook, datasheet, paddCols, jan16Row, book.datemode)

processedBook.close()