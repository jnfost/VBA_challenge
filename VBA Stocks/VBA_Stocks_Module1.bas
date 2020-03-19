Attribute VB_Name = "Module1"
Sub Stock_Market2()
'Loop through all stocks for 1 year and return the following
    'Ticker ("J")
    'Yearly change from opening price at beginning to closing at end ("K")
    '% change from opening at beginning to closing at end ("L")
    'Total stock volume ("M")
'Loop through all tickers on a worksheet
    'If ticker symbol is different, stop looking and do analysis
    
Dim ws As Worksheet

For Each ws In Worksheets
    
    Dim ticker As String
    Dim Summary_Table_Row As Double
    Summary_Table_Row = 2
    Dim open_price As Double
    Dim close_price As Double
    Dim price_change As Double
    Dim percent_change As Double
    Dim volume As Double
    
    volume = 0
    
    
    'Find last row of sheet
    Dim lrow As Double
    lrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    For r = 2 To lrow
        If ws.Cells(r + 1, 1) <> ws.Cells(r, 1).Value Then
        
        'Set ticker symbol
        ticker = ws.Cells(r, 1).Value
        
        'Print ticker in summary table
        ws.Range("J" & Summary_Table_Row).Value = ticker
        
        'Add 1 to summary table row
        Summary_Table_Row = Summary_Table_Row + 1
        
        End If
    Next r
    
    'Reset summary table row
        Summary_Table_Row = 2
        
    'Find opening price for each ticker value
    For r = 2 To lrow
        If ws.Cells(r + 1, 1) <> ws.Cells(r, 1).Value And IsEmpty(ws.Cells(r + 1, 1).Value) = False Then
        
        open_price = ws.Cells(r + 1, 3).Value
        ws.Range("N" & Summary_Table_Row + 1).Value = open_price
        ws.Range("N2") = ws.Cells(2, 3).Value
        
        'Add 1 to summary table row
        Summary_Table_Row = Summary_Table_Row + 1
    
        End If
    Next r
    
    'Reset Summary table row
        Summary_Table_Row = 2
    
    For r = 2 To lrow
        If ws.Cells(r + 1, 1) <> ws.Cells(r, 1).Value Then
        
        'Find closing price and store in summary table
        close_price = ws.Cells(r, 6).Value
        ws.Range("O" & Summary_Table_Row).Value = close_price
        
        'Find total Volume
        volume = volume + ws.Cells(r, 7)
        ws.Range("M" & Summary_Table_Row).Value = volume
        
        'Reset volume
        volume = 0
        
        'Add 1 to summary table row
        Summary_Table_Row = Summary_Table_Row + 1
        
        Else
        'Add to total volume
        volume = volume + ws.Cells(r, 7).Value
        
        End If
    Next r
    
    'Reset Summary table row
    Summary_Table_Row = 2
    
    'Find yearly change
    Dim slrow As Double
    slrow = ws.Cells(Rows.Count, "J").End(xlUp).Row
    For t = 2 To slrow
        price_change = ws.Cells(t, 15).Value - ws.Cells(t, 14).Value
        ws.Range("K" & Summary_Table_Row).Value = price_change
    
    'Find percent change
    'Convert to percentage
        If ws.Cells(t, 14).Value = 0 Then
            ws.Range("L" & Summary_Table_Row).Value = 0
            ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
        Else
            percent_change = price_change / ws.Cells(t, 14).Value
            ws.Range("L" & Summary_Table_Row).Value = percent_change
            ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
        End If
        
    'Add 1 to summary table row
        Summary_Table_Row = Summary_Table_Row + 1
    
    Next t
    
    'Set Headers for Summary Table
    ws.Range("J1") = "Ticker"
    ws.Range("K1") = "Yearly Change"
    ws.Range("L1") = "Percent Change"
    ws.Range("M1") = "Total Stock Volume"
    ws.Range("N1") = "Opening Price"
    ws.Range("O1") = "Closing Price"
    
    'Conditional formatting for yearly change (red - neg and green - pos)
    Dim lsumrow As Double
    lsumrow = ws.Cells(Rows.Count, "K").End(xlUp).Row
    For r = 2 To lsumrow
        If ws.Cells(r, 11).Value > 0 Then
            ws.Cells(r, 11).Interior.ColorIndex = 4
        ElseIf ws.Cells(r, 11).Value < 0 Then
            ws.Cells(r, 11).Interior.ColorIndex = 3
        Else
        End If
    Next r
    
ws.Columns("J:O").EntireColumn.AutoFit

Next ws

End Sub
