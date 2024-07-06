Sub VBA_Module_2()

'Loop for all worksheetes in workbook
For Each ws In Worksheets

    'Setting dims
    Dim Ticker As String

    Dim QuarterlyChange As Double
        QuarterlyChange = 0
    
    Dim PercentageChange As Double
        PercentageChange = 0
    
    Dim TotalStockVolume As Double
        TotalStockVolume = 0

    Dim OpenPrice As Double
        OpenPrice = ws.Cells(2, 3).Value
    
    Dim ClosePrice As Double
        ClosePrice = ws.Cells(2, 6).Value

    Dim Output As Integer
        Output = 2

    'Headers & Lable
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"

    'Count rows in the first column to stop at end
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastrow

    'If value of the next cell is same than that of the current cell keep going
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'Ticker name
            Ticker = ws.Cells(i, 1).Value
        
            'Add ticker name found
            ws.Range("I" & Output).Value = Ticker
        
            'Calculate QuarterlyChange
            ClosePrice = ws.Cells(i, 6).Value
            QuarterlyChange = (ClosePrice - OpenPrice)
        
            'Add calculated QuarterlyChange & formating
            ws.Range("J" & Output).Value = QuarterlyChange
            ws.Range("J" & Output).NumberFormat = "0.00"
        
            'Calculate PercentageChange
            PercentageChange = QuarterlyChange / OpenPrice
        
            'Add calculated PercentageChange & formating
            ws.Range("K" & Output).Value = PercentageChange
            ws.Range("K" & Output).NumberFormat = "0.00%"
        
            'Calculate TotalStockVolume
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        
            'Add calculated TotalStockVolume
            ws.Range("L" & Output).Value = TotalStockVolume
        
            'Reset TotalStockVolume for next ticker
            TotalStockVolume = 0
        
            'Reset OpenPrice for next % in output
            OpenPrice = ws.Cells(i + 1, 3)
        
            'Reset row and add one row for next output
            Output = Output + 1
        
        Else
            'Add TotalStockVolume
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        
    
        End If
    
    Next i

    'Count the unique tickers for color
    CountRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Color for QuarterlyChange
    For i = 2 To CountRow
        If ws.Cells(i, 10).Value > 0 Then
            'Green
            ws.Cells(i, 10).Interior.ColorIndex = 4
        Else
            'Red
            ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i

    For i = 2 To CountRow
            
    'Get Greatest % Increase
        If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & CountRow)) Then
            'Ticker name
            ws.Range("P2").Value = ws.Cells(i, 9).Value
            'Max % Value
            ws.Range("Q2").Value = ws.Cells(i, 11).Value
            'Formating cell
            ws.Range("Q2").NumberFormat = "0.00%"

    'Get Greatest % Decrease
        ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & CountRow)) Then
            'Ticker name
            ws.Range("P3").Value = ws.Cells(i, 9).Value
            'Min % Value
            ws.Range("Q3").Value = ws.Cells(i, 11).Value
            'Formating cell
            ws.Range("Q3").NumberFormat = "0.00%"
            
    'Get Greatest Total Volume %
        ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & CountRow)) Then
            'Ticker name
            ws.Range("P4").Value = ws.Cells(i, 9).Value
            'Max Volumn Value
            ws.Range("Q4").Value = ws.Cells(i, 12).Value
            'Formating cell
            ws.Range("Q4").NumberFormat = "0"
        End If
        
    Next i
ws.Columns("I:Q").AutoFit

Next ws

End Sub
