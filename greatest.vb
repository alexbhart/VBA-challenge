Sub greatest()

Dim i As Long
Dim lastrow As Long

lastrow = Cells(Rows.Count, "J").End(xlUp).Row


Range("o2").Value = "Greatest % Increase"
Range("o3").Value = "Greatest % Decrease"
Range("o4").Value = "Greatest Total Volume"
Range("p1").Value = "Ticker"
Range("q1").Value = "Value"


For i = 2 To lastrow
    
    Range("Q2").Value = Application.WorksheetFunction.Max(Range("L2:L" & lastrow))
    Range("Q3").Value = Application.WorksheetFunction.Min(Range("L2:L" & lastrow))
    Range("Q4").Value = Application.WorksheetFunction.Max(Range("M2:M" & lastrow))

Next i


For i = 2 To lastrow

    If Range("Q2").Value = Cells(i, 12).Value Then
        Range("P2").Value = Cells(i, 10).Value
    End If
    If Range("Q3").Value = Cells(i, 12).Value Then
        Range("P3").Value = Cells(i, 10).Value
    End If
    If Range("Q4").Value = Cells(i, 13).Value Then
        Range("P4").Value = Cells(i, 10).Value
    End If
      
Next i


Range("Q2:Q3").NumberFormat = "0.00%"

End Sub
