Sub stock()

'defining variables
Dim i, j, k, s, f, l, nmbSheets, stock_counter As Integer
Dim lrow As Long
Dim mySheets()
Dim v As String
Dim open_price, closing_price, total_stock_volume As Double
Dim max_increase, max_decrease, max_total_stock_volume as double


' Counting of number of sheets within a workbook modified from solutions
' presented in http://www.vbaexpress.com/forum/showthread.php?62270-An-Array-of-Sheet-Names&p=377930&viewfull=1#post377930
' total number of sheets within the active Workbook
nmbSheets = ActiveWorkbook.Sheets.Count

' Assign a dynamic array containing references to Sheets within the workbook cycling from 1st to the nmbSheets
ReDim mySheets(1 To nmbSheets)
'MsgBox ("Your active workbook has " + Str(nmbSheets) + " sheets")
  
' Red in individual sheet names into the dynamic array
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
    
    ' reset total stock volume value as it will be determined for each individual ticker below
    total_stock_volume = 0
    
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
            If opening_price = 0 Then 'And closing_price = 0 Then
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
    
    ' Initialize reference values for maximum and minimum decrease in stock value calculations
    max_increase = ActiveWorkbook.Sheets(mySheets(s)).Cells(2, 11).Value
    max_decrease = max_increase
    max_total_stock_volume = ActiveWorkbook.Sheets(mySheets(s)).Cells(2, 12).Value
    
    ' Change color of cells in J column based on their value. Set color to green if >0 and red - if <0
    For i = 2 To lrow
    'MsgBox ("Cell value in row  " + Str(lrow) + " is " + Str(ActiveWorkbook.Sheets(mySheets(s)).Cells(i, 10).Value))
        If ActiveWorkbook.Sheets(mySheets(s)).Cells(i, 10).Value > 0 Then
            ActiveWorkbook.Sheets(mySheets(s)).Cells(i, 10).Interior.ColorIndex = 4
        ElseIf ActiveWorkbook.Sheets(mySheets(s)).Cells(i, 10).Value < 0 Then
            ActiveWorkbook.Sheets(mySheets(s)).Cells(i, 10).Interior.ColorIndex = 3
        End If
        
    ' Check to see if current is the max/min % decrease
        If ActiveWorkbook.Sheets(mySheets(s)).Cells(i, 11).Value > max_increase Then
            max_increase = ActiveWorkbook.Sheets(mySheets(s)).Cells(i, 11).Value
            max_increase_stock = ActiveWorkbook.Sheets(mySheets(s)).Cells(i, 9).Value
        End If
        If ActiveWorkbook.Sheets(mySheets(s)).Cells(i, 11).Value < max_decrease Then
            max_decrease = ActiveWorkbook.Sheets(mySheets(s)).Cells(i, 11).Value
            max_decrease_stock = ActiveWorkbook.Sheets(mySheets(s)).Cells(i, 9).Value
        End If
        If ActiveWorkbook.Sheets(mySheets(s)).Cells(i, 12).Value > max_total_stock_volume Then
            max_total_stock_volume = ActiveWorkbook.Sheets(mySheets(s)).Cells(i, 12).Value
            max_total_stock_volume_stock = ActiveWorkbook.Sheets(mySheets(s)).Cells(i, 9).Value
        End If
    Next i
    
    ' report min/max % change statistics
    ' row headers
    ActiveWorkbook.Sheets(mySheets(s)).Range("O2").Value = "Greatest % increase"
    ActiveWorkbook.Sheets(mySheets(s)).Range("O3").Value = "Greatest % decrease"
    ActiveWorkbook.Sheets(mySheets(s)).Range("O4").Value = "Greatest Total Volume"
    ' column headers
    ActiveWorkbook.Sheets(mySheets(s)).Range("P1").Value = "Ticker"
    ActiveWorkbook.Sheets(mySheets(s)).Range("Q1").Value = "Value"
    
    ActiveWorkbook.Sheets(mySheets(s)).Range("P2").Value = max_increase_stock
    ActiveWorkbook.Sheets(mySheets(s)).Range("P3").Value = max_decrease_stock
    ActiveWorkbook.Sheets(mySheets(s)).Range("P4").Value = max_total_stock_volume_stock
    '
    ActiveWorkbook.Sheets(mySheets(s)).Range("Q2").Value = max_increase
        ActiveWorkbook.Sheets(mySheets(s)).Range("Q2").NumberFormat = "0.00%"
    ActiveWorkbook.Sheets(mySheets(s)).Range("Q3").Value = max_decrease
        ActiveWorkbook.Sheets(mySheets(s)).Range("Q3").NumberFormat = "0.00%"
    ActiveWorkbook.Sheets(mySheets(s)).Range("Q4").Value = max_total_stock_volume
Next s

End Sub
