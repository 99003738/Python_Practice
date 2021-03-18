import openpyxl as xl


path = "D:\Git_Python\Practice_problem\sheets.xlsx"
itemsearch = ['99003726', 'aman', 'g']

def sending_data_master(pathvariable, searchitem = []):

    # creating workbook from where data to be searched
    workbook1 = xl.load_workbook(pathvariable)
    workbooksheet = workbook1.sheetnames
    lengthworkbook1 = len(workbooksheet)

    # opening the destination worksheet
    pathMasterWorkbook = "C:\\Users\\99003738\\Desktop\\MasterBook.xlsx"
    masterbook = xl.load_workbook(pathMasterWorkbook)
    workSheetMaster = masterbook.active

    for i in range(0, lengthworkbook1):
        sheets = workbook1.worksheets[i]
        workSheetMaster = masterbook.active
        # calulating max row and max column

        wbMaxRow = sheets.max_row
        wbMaxColumn = sheets.max_column

        # taking length of the searchitem
        lengthSearch = len(searchitem)
        index =0
        rownum = 1
        colnum = 1
        for colnum in range(4, wbMaxColumn + 1):
            workSheetMaster.cell(row=rownum, column=colnum).value = sheets.cell(row=1, column=colnum).value
            colnum = colnum+1
        colnum = 1
        for i in range(1, wbMaxRow+1):
            for j in range(1, 4):
                data1 = searchitem[index]
                value1 =  sheets.cell(row= i , column= j)
                if data1 == value1:
                    rownum = rownum + 1
                    for k in range(4, wbMaxColumn+1):
                        workSheetMaster.cell(row= rownum, column= colnum).value = sheets.cell(row= i, column=k).value
                        colnum = colnum+1
                    colnum= 1
    masterbook.save(str(pathMasterWorkbook))


sending_data_master(path, itemsearch)
