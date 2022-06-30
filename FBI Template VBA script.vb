'Project Summary
    '1.) This code will import the BB FBI Template into the gl transactions template using a double for loop
    '2.) SHORTKEYS:
        'alt F11 to open VBA editor
        'F8 to run code line by line
        'F5 to run entire code

Public Sub import_FBI_data()


Dim FileToOpen As Variant
Dim OpenBook As Workbook

Dim wsSource As Worksheet
Dim wsDest As Worksheet

Dim i As Long
Dim j As Long
Dim lastrow1 As Long
Dim MY_LAST_ROW As Long


Application.ScreenUpdating = False


'Pop up to allow user to select which file they want to use as the source
 FileToOpen = Application.GetOpenFilename(Title:="Browse for your File & Import Range", FileFilter:="Excel Files (*.xls*),*xls*")
    If FileToOpen <> False Then
        'Set variables for the source and destination workbooks.worksheets
        Set OpenBook = Application.Workbooks.Open(FileToOpen)
        Set wsSource = OpenBook.Worksheets("Template")
        Set wsDest = ThisWorkbook.Worksheets("gl_transactions")
        
        'Clear contents of destination range before every import
        wsDest.Range("A2:A2000").EntireRow.ClearContents
        
        'Identify last row and column of the source workbook and starting point for the loop
        lastrow1 = wsSource.Cells(Rows.Count, 1).End(xlUp).Offset(-3, 0).Row
        lastcol = wsSource.Cells(5, Columns.Count).End(xlToLeft).Offset(0, -1).Column
            
            'Loop through every row and column starting in row 6 column 4 and grab all values skipping the blanks
            For i = 6 To lastrow1 Step 1
                For j = 4 To lastcol Step 1
                    If Not IsEmpty(Cells(i, j)) Then
                        amount = wsSource.Cells(i, j).Value
                        fund_code = wsSource.Cells(i, 2).Value
                        account_num = wsSource.Cells(4, j).Value
                        post_code = wsSource.Cells(3, 1).Value
                        description = wsSource.Cells(5, j).Value
                    
                        'Identify last row of destination workbook
                        MY_LAST_ROW = wsDest.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
                    
                        'Specify where FBI data will import to in destination workbook
                        With wsDest
                        wsDest.Cells(MY_LAST_ROW, 4).Value = amount
                        wsDest.Cells(MY_LAST_ROW, 2).Value = fund_code
                        wsDest.Cells(MY_LAST_ROW, 3).Value = account_num
                        wsDest.Cells(MY_LAST_ROW, 1).Value = post_code
                        wsDest.Cells(MY_LAST_ROW, 5).Value = description
                        End With
                    End If
                Next j
            Next i
    End If

OpenBook.Close False
    
Application.ScreenUpdating = True


End Sub
