Sub stock()

'defining variables
Dim i, j, k, s, f, l, nmbSheets, stock_counter As Integer
Dim lrow As Long
Dim mySheets()
Dim v As String
Dim open_price_beginning, closing_price_end As Double
Dim total_stock_volume As Double
Dim formattingRange As Range

total_stock_volume = 0


' OPTION 1: manual input names of sheets in the existing Workbook. Can we read in automatically like number of rows?
'mySheets = Array("A", "B", "C", "D", "E", "F", "P")

' OPTION 2: modified from http://www.vbaexpress.com/forum/showthread.php?62270-An-Array-of-Sheet-Names&p=377930&viewfull=1#post377930
nmbSheets = ActiveWorkbook.Sheets.Count

ReDim mySheets(1 To nmbSheets)
MsgBox ("Your active workbook has " + Str(nmbSheets) + " sheets")
  
For s = 1 To (nmbSheets)
    mySheets(s) = ActiveWorkbook.Sheets(s).Name
Next s


' loop through each sheet
For s = 1 To (nmbSheets)
    'Create and populate "Ticker" column
    ActiveWorkbook.Sheets(mySheets(s)).Range("I1").Value = "Ticker"
    ActiveWorkbook.Sheets(mySheets(s)).Range("J1").Value = "Yearly change"
    ActiveWorkbook.Sheets(mySheets(s)).Range("K1").Value = "Percent change"
    ActiveWorkbook.Sheets(mySheets(s)).Range("L1").Value = "Total stock volume"

    ' calculate number of rows of data in a given Sheet
    lrow = ActiveWorkbook.Sheets(mySheets(s)).Cells(Rows.Count, 1).End(xlUp).Row
    'MsgBox ("Number of rows in the Sheet '" + (mySheets(s)) + "' is " + Str(lrow))
    
    'reset position of stock name listing start on each sheet
    k = 2
    
    ' counter to determine opening price per ticker
    f = 0
    
    ' loop through each value in <ticker> column and output a list of unique values
    For i = 2 To (lrow)
        'Read in curren ticker name
        v = ActiveWorkbook.Sheets(mySheets(s)).Cells((i), 1).Value
        'If we are switching to a new ticker name, record opening price
        If f = 0 Then
            opening_price = ActiveWorkbook.Sheets(mySheets(s)).Cells((i), 3).Value
            f = f + 1
        End If
        
        total_stock_volume = total_stock_volume + ActiveWorkbook.Sheets(mySheets(s)).Cells(i, 7)
        
        ' check if the value is the same as the above one, if different - preserve it
        ' as a new unique value
        If v <> ActiveWorkbook.Sheets(mySheets(s)).Cells(i + 1, 1).Value Then
            'stock_counter = stock_counter + 1
            closing_price = ActiveWorkbook.Sheets(mySheets(s)).Cells((i), 6).Value
            ActiveWorkbook.Sheets(mySheets(s)).Cells(k, 9).Value = v
            ActiveWorkbook.Sheets(mySheets(s)).Cells(k, 10).Value = (closing_price - opening_price)
            If opening_price = 0 And closing_price = 0 Then
                ActiveWorkbook.Sheets(mySheets(s)).Cells(k, 11).Value = 0#
            Else
                ActiveWorkbook.Sheets(mySheets(s)).Cells(k, 11).Value = (closing_price - opening_price) / opening_price
            End If
            ActiveWorkbook.Sheets(mySheets(s)).Cells(k, 11).NumberFormat = "0.00%"
            
            ActiveWorkbook.Sheets(mySheets(s)).Cells(k, 12).Value = total_stock_volume
            
            
            
            'shifting summary output by one row
            k = k + 1
            ' resetting opening price per ticker flag
            f = 0
            ' resetting total stock volume per ticker value
            total_stock_volume = 0
        End If
    Next i
    'calculate number of rows of filled-in data for column J. It should be equal to a total number of ticker names + 1 for header
    lrow = ActiveWorkbook.Sheets(mySheets(s)).Cells(Rows.Count, 10).End(xlUp).Row
    
    ' Change color of cells in J column based on their value. Set color to green if >0 and red - if <0
    For i = 2 To lrow
    'MsgBox ("Cell value in row  " + Str(lrow) + " is " + Str(ActiveWorkbook.Sheets(mySheets(s)).Cells(i, 10).Value))
        If ActiveWorkbook.Sheets(mySheets(s)).Cells(i, 10).Value > 0 Then
            ActiveWorkbook.Sheets(mySheets(s)).Cells(i, 10).Interior.ColorIndex = 4
        ElseIf ActiveWorkbook.Sheets(mySheets(s)).Cells(i, 10).Value < 0 Then
            ActiveWorkbook.Sheets(mySheets(s)).Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i
        
Next s

End Sub
