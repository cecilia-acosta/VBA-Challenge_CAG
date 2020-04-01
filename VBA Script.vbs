Sub Stocks()

  Dim column, ref As Integer
  Dim OpenPrice, ClosePrice, YearChange, PercentChange, Total As Double
  Dim lastrow, Sumvol As Long
  Dim Ticker As String
  Dim ws As Worksheet
    
  Total = 0
  
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
column = 1
ref = 2
  
For Each ws In Worksheets

OpenPrice = ws.Range("C" & ref).Value

  For i = 2 To lastrow
    If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
                       
        Ticker = ws.Cells(i, column).Value
        ws.Range("I" & ref).Value = Ticker
                     
        ClosePrice = ws.Cells(i, 6).Value
        YearChange = ClosePrice - OpenPrice
            
            If OpenPrice = 0 Then
                PercentChange = 0
            Else
                PercentChange = (ClosePrice / OpenPrice) - 1
            End If
        
        ws.Range("J" & ref).Value = YearChange
        ws.Range("K" & ref).Value = PercentChange
        
            If PercentChange > 0 Then
                ws.Range("K" & ref).Interior.ColorIndex = 10
                ws.Range("K" & ref).Font.ColorIndex = 2
            Else
                ws.Range("K" & ref).Interior.ColorIndex = 3
                ws.Range("K" & ref).Font.ColorIndex = 2
            End If
                 
        Total = Total + ws.Cells(i, 7).Value
        Sumvol = ws.Cells(i, 7).Value + ws.Cells(i + 1, 7).Value
        ws.Range("L" & ref).Value = Total
        ref = ref + 1
        Total = 0
        
        OpenPrice = ws.Cells(i + 1, 3).Value
          
    Else
        Total = Total + ws.Cells(i, 7).Value
    
    End If
   Next i

Next ws

End Sub
