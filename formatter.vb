Sub formatting()

For i = 2 To Cells(Rows.Count, "L").End(xlUp).Row
    
    Cells(i, 12).NumberFormat = "0.00%"
    Cells(i, 11).NumberFormat = "$0.00"
    If Cells(i, 12).Value > 0 Then
        Cells(i, 11).Interior.ColorIndex = 4
        Else
        Cells(i, 11).Interior.ColorIndex = 3
    
    End If

Next i

        
End Sub
