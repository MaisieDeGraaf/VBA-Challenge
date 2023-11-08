Attribute VB_Name = "Module1"
Sub StockTicker():

    Dim Counter As Double
    Dim SummaryTable As Integer
    Dim DateOpen As Long
    Dim DateClosed As Long
    Dim OpenValue As Double
    Dim ClosingValue As Double
    Dim GreatestDecrease As Double
    Dim GreatestIncrease As Double
    Dim GreatestVolume As Double
    Dim rw As Long


    For Each ws In Worksheets
    
        ws.Range("A1:Q1, O2, O3, O4").Font.Bold = True
        ws.Cells(1, 9).Value = "Ticker"

        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        Counter = 0
        SummaryTable = 2
        DateOpen = ws.Cells(2, 2).Value
        DateClosed = ws.Cells(2, 2).Value
        
        Range("A1").Select
        rw = Range("A1").End(xlDown).Row
        'select to the end of the data
        ws.Range("A1", ws.Range("G1").End(xlDown)).Sort Key1:=ws.Range("A1"), Order1:=xlAscending, Header:=xlYes
        'In case the stocks have not been sorted, this will auto sort them
    
        For i = 2 To rw
        
            If (ws.Cells(i, 2).Value <= DateOpen) Then
        
                DateOpen = ws.Cells(i, 2).Value
                OpenValue = ws.Cells(i, 3).Value
                
            End If
        
            If (ws.Cells(i, 2).Value >= DateClosed) Then
            
                DateClosed = ws.Cells(i, 2).Value
                ClosingValue = ws.Cells(i, 6).Value
            
            End If
            
         If (ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value) Then
            
            Counter = Counter + ws.Cells(i, 7).Value
        
            
        Else
            Counter = Counter + ws.Cells(i, 7).Value
            
            ws.Cells(SummaryTable, 9).Value = ws.Cells(i, 1).Value
            ws.Cells(SummaryTable, 10).Value = ClosingValue - OpenValue
            ws.Cells(SummaryTable, 11).Value = Format((ClosingValue - OpenValue) / OpenValue, "Percent")
            ws.Cells(SummaryTable, 12).Value = Counter
            
                
            If ws.Cells(SummaryTable, 10).Value > 0 Then
                
                ws.Cells(SummaryTable, 10).Interior.Color = RGB(124, 252, 0)
                
            Else
                    
                ws.Cells(SummaryTable, 10).Interior.Color = RGB(255, 0, 0)
                
            End If
                
            Counter = 0
            SummaryTable = SummaryTable + 1
            DateOpen = ws.Cells(i + 1, 2).Value
            DateClosed = ws.Cells(i + 1, 2).Value
    
        End If

     Next i
     
     GreatestDecrease = ws.Cells(2, 11).Value
     GreatestIncrease = ws.Cells(2, 11).Value
     GreatestVolume = ws.Cells(2, 12).Value
     
     For i = 2 To rw

        
        If ws.Cells(i, 11).Value >= GreatestIncrease Then
            
            GreatestIncrease = ws.Cells(i, 11).Value
            ws.Cells(2, 17).Value = Format(GreatestIncrease, "Percent")
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            
        End If

     
        If ws.Cells(i, 11).Value <= GreatestDecrease Then
            
            GreatestDecrease = ws.Cells(i, 11).Value
            ws.Cells(3, 17).Value = Format(GreatestDecrease, "Percent")
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            
        End If

        If ws.Cells(i, 12).Value >= GreatestVolume Then
            
            GreatestVolume = ws.Cells(i, 12).Value
            ws.Cells(4, 17).Value = GreatestVolume
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            
        End If
        
    Next i
    
    ws.Range("A:P").Columns.AutoFit
    'Make all columns fit the width of the contents
        
    Next ws
        

End Sub

