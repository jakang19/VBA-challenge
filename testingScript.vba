Sub StockAnalysis()
    ' Loop through all worksheet
    For Each WS In Worksheets
        lastRow = WS.Cells(Rows.Count, 1).End(xlUp).row
        
        ' Add Heading for summary
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
        ' Create Variables
        Dim ticker As String
        Dim yearlyChange As Double
        Dim percentChange As Double
        Dim totalVol As Double
        Dim openPrice As Double
        Dim closePrice As Double
        totalVol = 0
        
        
        Dim i As Long
        
        openPrice = Cells(2, 3).Value
        
        For i = 2 To lastRow
            If Cells(i + 1, 2).Value <> Cells(i, 2).Value Then
                ticker = Cells(i, 2).Value
                Cells(2, 9).Value = ticker
                
                closePrice = Cells(i, 6).Value
                yearlyChange = closePrice - openPrice
                Cells(2, 10).Value = yearlyChange
                If (openPrice = 0 And closePrice = 0) Then
                    percentChange = 0
                ElseIf (openPrice = 0 And closePrice <> 0) Then
                    percentChange = 1
                Else
                    percentChange = yearlyChange / openPrice
                    Cells(2, 11).Value = percentChange
                    Cells(2, 11).NumberFormat = "0.00%"
                End If
                
    
End Sub
