Sub totalstockvolume()

Dim lastrow1 As Long
Dim lastrow2 As Long
Dim arg2 As String
Dim i As Long



lastrow1 = Cells(Rows.Count, "J").End(xlUp).Row
lastrow2 = Cells(Rows.Count, "A").End(xlUp).Row


Range("M2:M" & lastrow2).ClearContents


For i = 2 To lastrow1

    Cells(i, 13).Value = Application.WorksheetFunction.SumIf(Range("A2:A" & lastrow2), Range("J" & i), Range("G2:G" & lastrow2))
    
  
Next i

End Sub




