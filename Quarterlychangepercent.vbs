Sub QuarterlyChangePercent()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim worksheetNames As Variant
    Dim dictFirst As Object
    Dim dictLast As Object
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim change As Double
    Dim percentChange As Double
    Dim rowNum As Long
    Dim outputRowNum As Long

    Set wb = ThisWorkbook
    
    worksheetNames = Array("Q1", "Q2", "Q3", "Q4")
    
    For i = LBound(worksheetNames) To UBound(worksheetNames)
        
        Set ws = wb.Worksheets(worksheetNames(i))
        Set dictFirst = CreateObject("Scripting.Dictionary")
        Set dictLast = CreateObject("Scripting.Dictionary")
        
        rowNum = 2
        outputRowNum = 2
    
        Do While ws.Cells(rowNum, 1).Value <> ""
            ticker = ws.Cells(rowNum, 1).Value
            openingPrice = ws.Cells(rowNum, 3).Value
            closingPrice = ws.Cells(rowNum, 6).Value

            If Not dictFirst.Exists(ticker) Then
                dictFirst.Add ticker, openingPrice
            End If
            
            dictLast(ticker) = closingPrice
            rowNum = rowNum + 1
        Loop
        
        For Each key In dictFirst.Keys
            
            change = dictLast(key) - dictFirst(key)
            percentChange = (change / dictFirst(key)) * 100 / 100
            ws.Cells(outputRowNum, 11).Value = percentChange
            outputRowNum = outputRowNum + 1
        Next key
        ws.Cells(1, 10).Value = "Quaterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
    Next i
    
End Sub


