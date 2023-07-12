# VBA-Challenge

stockdata.bas contains the code.
The image files are the screenshots of the output.

Worked with a Learning Assistant to get the following code:

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
