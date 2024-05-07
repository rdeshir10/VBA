Sub ColorPercentchange()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim cell As Range

    Set wb = ThisWorkbook
    For Each ws In wb.Worksheets
       
        For Each cell In ws.Range("J2", ws.Cells(ws.Rows.count, 10).End(xlUp))
            
            If cell.Value > 0 Then
                
                cell.Interior.Color = RGB(0, 255, 0)
            ElseIf cell.Value < 0 Then
              
                cell.Interior.Color = RGB(255, 0, 0)
            Else
                cell.Interior.ColorIndex = xlNone
            End If
        Next cell
    Next ws
End Sub


