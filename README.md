# VBA-challenge

Sub MultipleYearStockData():

    For Each ws In Worksheets
    
        Dim WorksheetName As String
        Dim i As Long
        Dim row As Long
        Dim TickCount As Long
        Dim rowCount As Long
        Dim rowCount2 As Long
        Dim PerChange As Double
        Dim GreatInc As Double
        Dim GreatDec As Double
        Dim GreatTotVol As Double
        
        WorksheetName = ws.Name
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        TickerCount = 2
        
        row = 2
        
        rowCount = ws.Cells(Rows.Count, 1).End(xlUp).row
            
    For i = 2 To rowCount
            
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(TickerCount, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(TickerCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(row, 3).Value
                    If ws.Cells(TickerCount, 10).Value < 0 Then
                    ws.Cells(TickerCount, 10).Interior.ColorIndex = 3
                
                    Else
                
                    ws.Cells(TickerCount, 10).Interior.ColorIndex = 4
                
                    End If
                    
                    If ws.Cells(row, 3).Value <> 0 Then
                    PerChange = ((ws.Cells(i, 6).Value - ws.Cells(row, 3).Value) / ws.Cells(row, 3).Value)
                    ws.Cells(TickerCount, 11).Value = Format(PerChange, "Percent")
                    
                    Else
                    
                    ws.Cells(TickerCount, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                ws.Cells(TickerCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(row, 7), ws.Cells(i, 7)))
                
                TickerCount = TickerCount + 1
                
                row = i + 1
                
                End If
            
            Next i
            
        rowCount2 = ws.Cells(Rows.Count, 9).End(xlUp).row
       
        GreatTotVol = ws.Cells(2, 12).Value
        GreatInc = ws.Cells(2, 11).Value
        GreatDec = ws.Cells(2, 11).Value
        
            For i = 2 To rowCount2
                
                If ws.Cells(i, 12).Value > GreatTotVol Then
                GreatTotVol = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatTotVol = GreatTotVol
                
                End If
                    
                If ws.Cells(i, 11).Value > GreatInc Then
                GreatInc = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatInc = GreatInc
                
                End If
                
                If ws.Cells(i, 11).Value < GreatDec Then
                GreatDec = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatDec = GreatDec
                
                End If
                           
            ws.Cells(2, 17).Value = Format(GreatInc, "Percent")
            ws.Cells(3, 17).Value = Format(GreatDec, "Percent")
            ws.Cells(4, 17).Value = Format(GreatTotVol, "Scientific")
            
            Next i
            
        
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
            
    Next ws
        
End Sub
