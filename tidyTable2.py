import openpyxl

def tidyTable(): 
    wb = openpyxl.load_workbook('lengths3.xlsx')
    ws1 = wb['Sheet2']
    rows = tuple(ws1.rows)
    columns = tuple(ws1.columns)

    rowCount = 1
    colCount = 1

    for row in rows:
        for column in columns:
            currentCell = ws1.cell(row=rowCount, column=colCount)
            if currentCell.value == 0:
                currentCell.value = None
                colCount += 1
                #print('FOO')
            else:
                colCount +=1
        rowCount += 1
        colCount = 1

    wb.save('tidyTable.xlsx')
        
tidyTable()