

Sub stocksi()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim ticker As String
    Dim opening As Double
  ' Added to store opening date
    Dim closing As Double
    Dim summary_table_row As Integer

    For Each ws In Worksheets
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        ws.Range("H1:K1").EntireColumn.Insert
        ws.Cells(1, 8).Value = "Ticker"
        ws.Cells(1, 9).Value = "Yearly change"
        ws.Cells(1, 10).Value = "Percent change"
        ws.Cells(1, 11).Value = "Total stock value"

        summary_table_row = 2
        Ticker_Total = 0

        For i = 2 To LastRow
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                opening = ws.Cells(i, 3).Value
            End If
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                closing = ws.Cells(i, 3).Value
                Ticker_Total = Ticker_Total + Cells(i, 7).Value

                ' Assuming you want to calculate Yearly change, Percent change, and Total stock value here
                ws.Cells(summary_table_row, 8).Value = ticker
                ws.Cells(summary_table_row, 9).Value = closing - opening
                ws.Cells(summary_table_row, 10).Value = Round(100 * ((closing - opening) / opening), 4) & "%"
        
              

                summary_table_row = summary_table_row + 1
                opening = ws.Cells(i + 1, 6).Value
                
            Else
                Ticker_Total = Ticker_Total + Cells(i, 7).Value
            End If
            ws.Cells(summary_table_row, 11).Value = Ticker_Total
        Next i
    Next ws
    
        
        
End Sub
