Sub ticker():

Dim i As Long
Dim j As Long
Dim st_open As Double
Dim st_close As Double



Range("J1").Value = "Ticker Symbol"
Range("K1").Value = "Yearly Change"
Range("L1").Value = "Percent Change"
Range("M1").Value = "Total Stock Volume"

j = 2

  ' Loop through rows in the column
For i = 2 To Cells(Rows.Count, "A").End(xlUp).Row

    ' Searches for when the value of the next cell is different than that of the current cell
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    'sends last found value to new column
        Cells(j, 10).Value = Cells(i, 1).Value
        j = j + 1
            
    End If
 

Next i


j = 2



j = 2


For i = 2 To Cells(Rows.Count, "A").End(xlUp).Row
    ' grabs first opening value
     If Cells(j, 10).Value = Cells(i, 1).Value And Cells(i + 1, 1).Value = Cells(i, 1).Value Then
          st_open = Cells(i, 3).Value
        '  Cells(j, 8).Value = st_open
          j = j + 1
     End If
    ' grabs end of year close
     If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        st_close = Cells(i, 6).Value
   
      '  Cells(j, 9).Value = st_close
        Cells(j - 1, 11).Value = st_open - st_close
        Cells(j - 1, 12).Value = (st_open - st_close) / st_close
      
    
    End If


Next i
    

End Sub






