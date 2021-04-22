Sub stock()

'defining variables
Dim i, j, k, s, f, l, nmbSheets, stock_counter As Integer
Dim lrow As Long
Dim mySheets()
Dim v As String
Dim open_price_beginning, closing_price_end As Double
Dim total_stock_volume As Double

total_stock_volume = 0


' OPTION 1: manual input names of sheets in the existing Workbook. Can we read in automatically like number of rows?
'mySheets = Array("A", "B", "C", "D", "E", "F", "P")

' OPTION 2: modified from http://www.vbaexpress.com/forum/showthread.php?62270-An-Array-of-Sheet-Names&p=377930&viewfull=1#post377930
nmbSheets = ActiveWorkbook.Sheets.Count
ReDim mySheets(1 To nmbSheets)
MsgBox ("Your active workbook has " + Str(nmbSheets) + " sheets")
  
For s = 1 To (nmbSheets)
    'ReDim Preserve mySheets(k): k = k + 1
    mySheets(s) = ActiveWorkbook.Sheets(s).Name
    'MsgBox ("Sheet name'" + mySheets(s) + "'")
Next s


' loop through each sheet
For s = 1 To (nmbSheets)
    ' calculate number of rows of data in a given Sheet
    'MsgBox ("Sheet name'" + mySheets(s) + "'")
    lrow = ActiveWorkbook.Sheets(mySheets(s)).Cells(Rows.Count, 1).End(xlUp).Row
    MsgBox ("Number of rows in the Sheet '" + (mySheets(s)) + "' is " + Str(lrow))
    
    'reset position of stock name listing start on each sheet
    k = 2
    
    
    
    ' loop through each value in <ticker> column and output a list of unique values
    For i = 1 To (lrow - 1)
        total_stock_volume = total_stock_volume + ActiveWorkbook.Sheets(mySheets(s)).Cells(i + 1, 7)
        v = ActiveWorkbook.Sheets(mySheets(s)).Cells((i + 1), 1).Value
        'open_price_beginning = ActiveWorkbook.Sheets(mySheets(s)).Cells((i + 1), 3).Value
        ' check if the value is the same as the above one, if different - preserve it
        ' as a new unique value
        If v = ActiveWorkbook.Sheets(mySheets(s)).Cells(i, 1).Value Then
            ' check year
        Else
            stock_counter = stock_counter + 1
            ActiveWorkbook.Sheets(mySheets(s)).Cells(k, 9).Value = v
            'closing_price_end = ActiveWorkbook.Sheets(mySheets(s)).Cells(i, 6).Value
            ActiveWorkbook.Sheets(mySheets(s)).Cells(k - 1, 12).Value = total_stock_volume
            k = k + 1
            total_stock_volume = 0
        End If
    Next i
    'Create and populate "Ticker" column
    ActiveWorkbook.Sheets(mySheets(s)).Range("I1").Value = "Ticker"
    ActiveWorkbook.Sheets(mySheets(s)).Range("J1").Value = "Yearly change"
    ActiveWorkbook.Sheets(mySheets(s)).Range("K1").Value = "Percent change"
    ActiveWorkbook.Sheets(mySheets(s)).Range("L1").Value = "Total stock volume"
Next s


End Sub

