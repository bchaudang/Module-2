

Sub sortStock():

    'variables
    
    Dim Ticker As String
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim TotalVolume As Double
    Dim PercentChange As Double
    Dim YearChange As Integer
    Dim TableSummary As Double
    
    Dim Increase As String
    Dim Decrease As String
    Dim Volume As Integer
     
        
    TableSummary = 2
    YearlyChange = 0
    TickerNew = 2
   
    
    Dim ws As Worksheet
    
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    For Each ws In Worksheets
    
    'insert new headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    
    
    For i = 2 To lastRow
    
        'Sort Ticker Symbols
     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
     
        'Set Tickers and Total Volume
     Ticker = ws.Cells(i, 1).Value
     TotalVolume = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(TickerNew, 7), ws.Cells(i, 7)))
     
        'Opening and Closing Stock Prices
     OpeningPrice = ws.Cells(TickerNew, 3).Value
     ClosingPrice = ws.Cells(i, 6).Value
     TickerNew = (i + 1)
     
        'Find Yearly and Percent Changes
     YearlyChange = ClosingPrice - OpeningPrice
     PercentChange = YearlyChange / OpeningPrice
        'Set Decimal Point for Percentage
     ws.Range("K" & TableSummary).NumberFormat = "0.00%"
     
        'Output of data into columns
     ws.Range("I" & TableSummary).Value = Ticker
     ws.Range("J" & TableSummary).Value = YearlyChange
     ws.Range("K" & TableSummary).Value = PercentChange
     ws.Range("L" & TableSummary).Value = TotalVolume
     
        'Conditional Formatting
    If ws.Range("J" & TableSummary).Value > 0 Then
        ws.Range("J" & TableSummary).Interior.ColorIndex = 4
    
    ElseIf ws.Range("J" & TableSummary).Value < 0 Then
            ws.Range("J" & TableSummary).Interior.ColorIndex = 3
    End If
    
    If ws.Range("K" & TableSummary).Value > 0 Then
        ws.Range("K" & TableSummary).Interior.ColorIndex = 10
    
    ElseIf ws.Range("K" & TableSummary).Value < 0 Then
            ws.Range("K" & TableSummary).Interior.ColorIndex = 9
    End If
             
         'Greatest Increase, Decrease and Volume
    ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & lastRow)) * 100
    ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & lastRow)) * 100
    ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & lastRow))
    
    Increase = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastRow)), ws.Range("K2:K" & lastRow), 0)
    ws.Range("P2") = Cells(Increase + 1, 9)
    Decrease = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastRow)), ws.Range("K2:K" & lastRow), 0)
    ws.Range("P3") = Cells(Decrease + 1, 9)
    Volume = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastRow)), ws.Range("L2:L" & lastRow), 0)
    ws.Range("P4") = Cells(Volume + 1, 9)
    
           
                        
     TableSummary = TableSummary + 1
             
        
     End If
     
    Next i
    
    TableSummary = 2
    YearlyChange = 0
    Next ws
    
    End Sub
    
     



