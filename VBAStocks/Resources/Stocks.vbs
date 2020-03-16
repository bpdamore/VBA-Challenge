Sub stockTicker():

    'Keep these out of the ws loop. Only going on the first ws
    Dim greatInc As Double
    Dim greatIncTic As String
    Dim greatDec As Double
    Dim greatDecTic As String
    Dim greatVol As Double
    Dim greatVolTic As String

    'Set variables for reference outside the ws loop
    greatInc = 0
    greatDec = 0
    greatVol = 0

    For Each ws In Worksheets
        'Headers for each sheet
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    
        'This auto-detects the last row.
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        Dim ticker As String
        Dim stockVol As Double
        Dim yearChange As Double
        Dim percentChange As Double
        Dim openStock As Double
        Dim closeStock As Double

        'Set variables for reference, outside the i loop
        stockVol = 0
        openStock = ws.Cells(2, 3).Value
        k = 2
    
        For i = 2 To lastrow
    
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                closeStock = ws.Cells(i, 6).Value
                yearChange = (closeStock - openStock)
                
                'Prevents any DIV/0 errors
                If openStock = 0 Then
                    percentChange = 0
                Else
                    percentChange = ((closeStock - openStock) / openStock)
                End If
                
                stockVol = stockVol + ws.Cells(i, 7).Value
                ws.Cells(k, 9).Value = ticker
                ws.Cells(k, 10).Value = yearChange
                
                'Conditional formatting
                If yearChange > 0 Then
                    ws.Cells(k, 10).Interior.ColorIndex = 4
                ElseIf yearChange < 0 Then
                    ws.Cells(k, 10).Interior.ColorIndex = 3
                End If
                    
                ws.Cells(k, 11).Value = Format(percentChange, "Percent")
                ws.Cells(k, 12).Value = stockVol

                openStock = ws.Cells(i + 1, 3).Value
                stockVol = 0
                k = k + 1

            'Adding total stock volume when cells(i,1) and cells(i+1) match.
            Else
                stockVol = stockVol + ws.Cells(i, 7).Value
            
            End If
            
        Next i
        
        For j = 2 To lastrow
        
            'Scrapes through each ws for the overall analysis on the first sheet
            If ws.Cells(j, 11).Value > greatInc Then
                greatInc = ws.Cells(j, 11).Value
                greatIncTic = ws.Cells(j, 9).Value
            End If
            If ws.Cells(j, 11).Value < greatDec Then
                greatDec = ws.Cells(j, 11).Value
                greatDecTic = ws.Cells(j, 9).Value
            End If
            If ws.Cells(j, 12).Value > greatVol Then
                greatVol = ws.Cells(j, 12).Value
                greatVolTic = ws.Cells(j, 9)
            End If
            
        Next j
        
        ws.Columns("I:L").AutoFit
        
    Next ws
    
    
    
    'Headers for the overall analysis
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Volume"
    
    'Overall analysis for the first sheet.
    'Outside the ws loop so the variables have been fully calculated.
    Cells(2, 16).Value = greatIncTic
    Cells(3, 16).Value = greatDecTic
    Cells(4, 16).Value = greatVolTic
    Cells(2, 17).Value = Format(greatInc, "Percent")
    Cells(3, 17).Value = Format(greatDec, "Percent")
    Cells(4, 17).Value = greatVol
    
    Columns("O:Q").AutoFit
    
End Sub


