Sub HighestTotalVolume()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim worksheetNames As Variant
    Dim rowNum As Long
    Dim highestValue As Double
    Dim highestTicker As String
    
    highestValue = -1E+308
    highestTicker = ""
    
    Set wb = ThisWorkbook
    
    worksheetNames = Array("Q1", "Q2", "Q3", "Q4")
    
    For i = LBound(worksheetNames) To UBound(worksheetNames)
        
        Set ws = wb.Worksheets(worksheetNames(i))
        
     rowNum = 2
        Do While ws.Cells(rowNum, 12).Value <> ""
            Dim currentValue As Double
            currentValue = ws.Cells(rowNum, 12).Value
            
            If currentValue > highestValue Then
                highestValue = currentValue
                highestTicker = ws.Cells(rowNum, 1).Value
            End If
 
    rowNum = rowNum + 1
        Loop
    Next i
    
    wb.Sheets(1).Cells(4, 16).Value = highestValue
    wb.Sheets(1).Cells(4, 15).Value = highestTicker
    wb.Sheets(1).Cells(4, 14).Value = "Greatest Total Volume"
    
End Sub


