Attribute VB_Name = "Module2"
Sub stockAnalysis()
For Each ws In Worksheets
    ' Set ticker as a string
    Dim ticker As String
    ' Set stock volume (vol) as double, since it is a decimal number
    Dim vol As Double
    ' Set starting vol at 0, because we want it to start with no value
    vol = 0
    ' Summary Table
    Dim summaryTableRow As Integer
    summaryTableRow = 2
    'Go to Prompt #2 and track yearly change from opening price at the beginning of a year
    ' to the closing proce at the end of a year
    ' yearlyChange = closePrice - openPrice
    'Set openPrice as Double since it is a decimal number
    Dim openPrice As Double
    ' Tell computer where openPrice values are
    openPrice = ws.Cells(2, 3).Value
    
    'Set closingPrice, yearlyChange, and percentChange as Double because they are decimals
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    
    
    ' Insert summary_table into workbook
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Stock Volume"
    
    'Count all the rows in column 1
    RowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Start my for loop to loop through all rows by ticker
    'Start at row 2 and to to the bottom
    For I = 2 To RowCount
        ' If the ticker value below is not equal to the current ticker values then
        If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
        ' print the ticker value into the ticker column on summary table
        ticker = ws.Cells(I, 1).Value
        ' Add vol of each ticker as long as above condition remains true
        ' ticker volume = current volume + volume of cell below in column 7
        vol = vol + ws.Cells(I, 7).Value
        
        ' Define closing price values
        closingPrice = ws.Cells(I, 6).Value
        ' Print values into summaryTableRow
        ws.Range("I" & summaryTableRow).Value = ticker
        
        'find yearly change and print it into J
        yearlyChange = (closingPrice - openPrice)
        ws.Range("J" & summaryTableRow).Value = yearlyChange
        
        ' Find total stock Volume and print it into L
        ws.Range("L" & summaryTableRow).Value = vol
        
        'I was getting an error 6 Overflow when I tried to run
        'After some research, I need to include a process so the code doesn't divide by zero
        If openPrice = 0 Then
            percentChange = 0
        Else
            percentChange = yearlyChange / openPrice
    End If
        'find percent change and print it into "K"
        ws.Range("K" & summaryTableRow).Value = percentChange
        ws.Range("K" & summaryTableRow).NumberFormat = "0.00%"
        
        ' Reset everything
        ' Add 1 to the summary table row
        summaryTableRow = summaryTableRow + 1
        
        vol = 0
        
        openPrice = ws.Cells(I + 1, 3)
        
        Else
            vol = vol + ws.Cells(I, 7).Value
      
        End If
    Next I
    ' conditional formatting for the cells in summary table
    lastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
        For I = 2 To lastRow
            If ws.Cells(I, 10).Value > 0 Then
                ws.Cells(I, 10).Interior.ColorIndex = 10
            Else
                ws.Cells(I, 10).Interior.ColorIndex = 3
            End If
        Next I
        
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        For I = 2 To lastRow
        
            If ws.Cells(I, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastRow)) Then
            ws.Cells(2, 16).Value = ws.Cells(I, 9).Value
            ws.Cells(2, 17).Value = ws.Cells(I, 11).Value
            ws.Cells(3, 17).NumberFormat = "0.00%"
            
            ElseIf ws.Cells(I, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastRow)) Then
                ws.Cells(3, 16).Value = ws.Cells(I, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(I, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
                
            ElseIf ws.Cells(I, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastRow)) Then
                ws.Cells(4, 16).Value = ws.Cells(I, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(I, 12).Value
                
            End If
    
    
    Next I
    
    
    Next ws
End Sub
