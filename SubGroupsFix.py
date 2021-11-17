from Try_to_find_reference import *

def GroupFix(filename):
    rb = xlrd.open_workbook(filename,formatting_info=True)
    sheet = rb.sheet_by_index(0)
    result = []
    for i in range(0,rowsCounter(filename,0)):
            result.append([replaceSympols(firstUpper(str(sheet.row_values(i)[0]).strip())), replaceSympols(firstUpper(str(sheet.row_values(i)[1]).strip())),
                           replaceSympols(firstUpper(str(sheet.row_values(i)[2]).strip())), replaceSympols(firstUpper(str(sheet.row_values(i)[3]).strip())),
                           firstUpper(str(sheet.row_values(i)[4]).strip()), firstUpper(str(sheet.row_values(i)[5]).strip()),
                           firstUpper(str(sheet.row_values(i)[6]).strip()), replaceSympols(firstUpper(str(sheet.row_values(i)[7]).strip())),
                           replaceSympols(firstUpper(str(sheet.row_values(i)[8]).strip()))])
    wb = Workbook()
    ws = wb.active

    for subarray in result:
        ws.append(subarray)

    wb.save('Імпорт підгруппи Ольжича.xlsx')

GroupFix('Імпорт подгруппы Ольжич.xls')
           
            
