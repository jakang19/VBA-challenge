Sub StockAnalysis()

    ' Loop through all worksheet
    For Each WS In Worksheets
        lastRow = WS.Cells(Rows.Count, 1).End(xlUp).row
        
        ' Headers
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
        
        Dim row As Integer
        row = 2
        
        ' initialize openPrice
        openPrice = Cells(2, 3).Value
        
        ' set ticker symbol
        For i = 2 To lastRow
            ' check if ticker symbol has changed
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                ' Generate ticker name
                ticker = Cells(i, 1).Value
                Cells(row, 9).Value = ticker
                closePrice = Cells(i, 6).Value
                ' Calculate yearlyChange
                yearlyChange = closePrice - openPrice
                Cells(row, 10).Value = yearlyChange
                ' Calculate percentChange
                If (openPrice = 0 And closePrice = 0) Then
                    percentChange = 0
                ElseIf (openPrice = 0 And closePrice <> 0) Then
                    percentChange = 1
                Else
                    percentChange = yearlyChange / openPrice
                    Cells(row, 11).Value = percentChange
                    Cells(row, 11).NumberFormat = "0.00%"
                End If
                ' Calculate total Volume
                totalVol = totalVol + Cells(i, 7).Value
                Cells(row, 12).Value = totalVol
                ' go to next row
                row = row + 1
                ' reset variables to recalculate
                openPrice = Cells(i + 1, 3)
                totalVol = 0
            ' if ticker symbol hasn't changed
            Else
                totalVol = totalVol + Cells(i, 7).Value
            End If
        Next i
        
        ' Conditional Formatting for Yearly Change
        ' Last row of yearlyChange for each WS
        lastRowYC = WS.Cells(Rows.Count, 9).End(xlUp).row
        ' Colors
        For i = 2 To lastRowYC
            If (Cells(i, 10).Value > 0 Or Cells(i, 10).Value = 0) Then
                Cells(i, 10).Interior.ColorIndex = 10
            ElseIf Cells(i, 10).Value < 0 Then
                Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i
        
    Next WS
End Sub
