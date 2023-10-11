Attribute VB_Name = "TickerCompile_2"
Sub CalculateStockData()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim SummaryRow As Long
    Dim YearlyChangeColor As Long
    Dim PercentageChangeColor As Long
    Dim MaxPercentIncrease As Double
    Dim MaxPercentDecrease As Double
    Dim MaxTotalVolume As Double
    Dim MaxPercentIncreaseTicker As String
    Dim MaxPercentDecreaseTicker As String
    Dim MaxTotalVolumeTicker As String
    
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row in the current sheet
        LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
       
	   ' Set initial values for maximum values for each worksheet
		MaxPercentIncrease = 0
		MaxPercentDecrease = 0
		MaxTotalVolume = 0
        MaxPercentIncreaseTicker = ""
        MaxPercentDecreaseTicker = ""
        MaxTotalVolumeTicker = ""
        
        ' Set initial values
        Ticker = ws.Cells(2, 1).Value
        OpenPrice = ws.Cells(2, 3).Value
        TotalVolume = 0
        SummaryRow = 2
        
        ' Add headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Loop through the rows in the current sheet
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> Ticker Then
                ' Calculate YearlyChange, PercentChange, and TotalVolume
                ClosePrice = ws.Cells(i, 6).Value
                YearlyChange = ClosePrice - OpenPrice
                If OpenPrice <> 0 Then
                    PercentChange = (YearlyChange / OpenPrice) * 100
                Else
                    PercentChange = 0
                End If
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value ' Accumulate TotalVolume only for the same ticker
                
                ' Determine cell color based on YearlyChange
                If YearlyChange > 0 Then
                    YearlyChangeColor = RGB(0, 255, 0) ' Green for positive YearlyChange
                ElseIf YearlyChange < 0 Then
                    YearlyChangeColor = RGB(255, 0, 0) ' Red for negative YearlyChange
                Else
                    YearlyChangeColor = RGB(255, 255, 255) ' White for zero YearlyChange
                End If
                
                ' Determine cell color based on PercentageChange
                If PercentChange > 0 Then
                    PercentageChangeColor = RGB(0, 255, 0) ' Green for positive YearlyChange
                ElseIf PercentChange < 0 Then
                    PercentageChangeColor = RGB(255, 0, 0) ' Red for negative YearlyChange
                Else
                    PercentageChangeColor = RGB(255, 255, 255) ' White for zero YearlyChange
                End If
                
                ' Output data to summary table with cell color formatting
                ws.Cells(SummaryRow, 9).Value = Ticker
                ws.Cells(SummaryRow, 10).Value = YearlyChange
                ws.Cells(SummaryRow, 10).Interior.Color = YearlyChangeColor
                ws.Cells(SummaryRow, 11).Interior.Color = PercentageChangeColor
                ws.Cells(SummaryRow, 11).Value = Format(PercentChange, "0.00") & "%"
                ws.Cells(SummaryRow, 12).Value = TotalVolume
                
                ' Check for Greatest % Increase, Greatest % Decrease, and Greatest Total Volume
                If PercentChange > MaxPercentIncrease Then
                    MaxPercentIncrease = PercentChange
                    MaxPercentIncreaseTicker = Ticker
                ElseIf PercentChange < MaxPercentDecrease Then
                    MaxPercentDecrease = PercentChange
                    MaxPercentDecreaseTicker = Ticker
                End If
                
                If TotalVolume > MaxTotalVolume Then
                    MaxTotalVolume = TotalVolume
                    MaxTotalVolumeTicker = Ticker
                End If
                
                ' Move to the next row in the summary table
                SummaryRow = SummaryRow + 1
                
                ' Reset values for the next Ticker
                Ticker = ws.Cells(i + 1, 1).Value
                OpenPrice = ws.Cells(i + 1, 3).Value
                TotalVolume = 0 ' Reset TotalVolume for the new ticker
            Else
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            End If
        Next i
        
        ' Add headers
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        
		' Output Greatest % Increase, Greatest % Decrease, and Greatest Total Volume
        ws.Cells(2, 15).Value = MaxPercentIncreaseTicker
        ws.Cells(3, 15).Value = MaxPercentDecreaseTicker
        ws.Cells(4, 15).Value = MaxTotalVolumeTicker
        ws.Cells(2, 16).Value = Format(MaxPercentIncrease, "0.00") & "%"
        ws.Cells(3, 16).Value = Format(MaxPercentDecrease, "0.00") & "%"
        ws.Cells(4, 16).Value = MaxTotalVolume
    Next ws
End Sub

