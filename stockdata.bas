Attribute VB_Name = "Module1"
Sub stockdata()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Dim column As Integer
        Dim LastRow As Long
        Dim j As Integer
        Dim h As Double
        Dim v As Double
        Dim increase As Double
        Dim decrease As Double
        Dim greatestVolume As Double

        column = 1
        j = 2
        h = 2
        v = 0
        
        LastRow = ws.Cells(ws.Rows.Count, column).End(xlUp).Row
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        For i = 2 To LastRow
            v = v + ws.Cells(i, 7).Value
            If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
                ' Ticker
                ws.Cells(j, 9).Value = ws.Cells(i, column).Value
                ' Yearly Change
                ws.Cells(j, 10).Value = ws.Cells(i, 6).Value - ws.Cells(h, 3).Value
                ' If Yearly Change cell negative, format to red color
                ' Otherwise, format to green color
                If ws.Cells(j, 10).Value < 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(j, 10).Interior.ColorIndex = 4
                End If
                ' Percent Change
                ws.Cells(j, 11).Value = ws.Cells(j, 10).Value / ws.Cells(h, 3).Value
                ws.Cells(j, 11).NumberFormat = "0.00%"
                ' Total Stock Volume
                ws.Cells(j, 12).Value = v
                v = 0
                j = j + 1
                h = i + 1
            End If
        Next i
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ' Greatest % Increase
       ws.Range("Q2") = WorksheetFunction.Max(ws.Range("K2:K" & LastRow))
       ws.Range("Q2").NumberFormat = "0.00%"
       increase = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & LastRow)), ws.Range("K2:K" & LastRow), 0)
       ws.Range("P2") = ws.Cells(increase + 1, 9)
       
       ' Greatest % Decrease
       ws.Range("Q3") = WorksheetFunction.Min(ws.Range("K2:K" & LastRow))
       ws.Range("Q3").NumberFormat = "0.00%"
       decrease = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & LastRow)), ws.Range("K2:K" & LastRow), 0)
       ws.Range("P3") = ws.Cells(decrease + 1, 9)
       
       ' Greatest Total Volume
       ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & LastRow))
       greatestVolume = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & LastRow)), ws.Range("L2:L" & LastRow), 0)
       ws.Range("P4") = ws.Cells(greatestVolume + 1, 9)
    Next ws
End Sub
