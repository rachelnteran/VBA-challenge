Sub greatest()
    Dim ws As Worksheet
    Dim LastRow As Long
    For Each ws In Worksheets
    LastRow = ws.Cells(ws.Rows.Count, 8).End(xlUp).Row
        ws.Cells(2, 14).Value = "Greatest % increase"
        ws.Cells(3, 14).Value = "Greatest % decrease"
        ws.Cells(4, 14).Value = "Greatest total Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        Max_increase = ws.Cells(2, 10)
        Min_decrease = 0
        greatestVolume = ws.Cells(2, 11)
        For i = 2 To LastRow
        If ws.Cells(i, 10).Value > Max_increase Then
        Max_increase = ws.Cells(i, 10)
        greatestincname = ws.Cells(i, 8)
        End If
        If ws.Cells(i, 10).Value < Min_decrease And ws.Cells(i, 10) < ws.Cells(i + 1, 10).Value Then
        Min_decrease = ws.Cells(i, 10)
        GreatestDecName = ws.Cells(i, 8)
        End If
        If ws.Cells(i, 11).Value > greatestVolume Then
        greatestVolume = ws.Cells(i, 11).Value
        greatestvolname = ws.Cells(i, 8)
        End If
        ws.Cells(2, 15).Value = greatestincname
        ws.Cells(3, 15).Value = GreatestDecName
        ws.Cells(4, 15).Value = greatestvolname
        ws.Cells(2, 16).Value = Max_increase * 100 & "%"
        ws.Cells(3, 16).Value = Min_decrease * 100 & "%"
        ws.Cells(4, 16).Value = greatestVolume
        
        Next i
    Next ws
    
End Sub
