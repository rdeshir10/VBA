Sub TickerTotalSum()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Dim ticker As String
    Dim dict As Object
    Dim totalSum As Double
    Dim worksheetsToProcess As Variant
    Dim row As Long
    Dim tickerRow As Long

    Set wb = ThisWorkbook

    worksheetsToProcess = Array("Q1", "Q2", "Q3", "Q4")

    For i = LBound(worksheetsToProcess) To UBound(worksheetsToProcess)
        Set ws = wb.Worksheets(worksheetsToProcess(i))
        Set dict = CreateObject("Scripting.Dictionary")
        
        row = 2
        Do While ws.Cells(row, 1).Value <> ""
            ticker = ws.Cells(row, 1).Value
            If dict.Exists(ticker) Then

                dict(ticker) = dict(ticker) + ws.Cells(row, 7).Value
            Else
                dict.Add ticker, ws.Cells(row, 7).Value
            End If
            row = row + 1
        Loop
        
        Set rng = ws.Range("I2:I" & ws.Cells(ws.Rows.count, 9).End(xlUp).row)

        For Each cell In rng
            ticker = cell.Value
            If dict.Exists(ticker) Then
                cell.Offset(0, 3).Value = dict(ticker)
            End If
        Next cell

        ws.Cells(1, 12).Value = "Total Stock volume"
        dict.RemoveAll
    Next i
    
End Sub

