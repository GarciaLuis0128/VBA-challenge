Attribute VB_Name = "Module1"
Sub LoopforStocks()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim outputRow As Long
    MaxPercentIncrease = 0
    MaxPercentDecrease = 0
    MaxTotalVolume = 0
    MaxPercentIncreaseTicker = ""
    MaxPercentDecreaseTicker = ""
    MaxTotalVolumeTicker = ""
    
    'Worksheet I'm using
    For Each ws In ThisWorkbook.Worksheets
        ' Initialize variables for the current worksheet
        SummaryRow = 2
        totalVolume = 0
        openPrice = ws.Cells(2, 3).Value
        ' Find the last row in the current worksheet
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            
    'Last row of data
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Where I want it to start the results
    outputRow = 2
    
    'Headers I want in this sheet
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    'Data for the 1st Ticker
    openPrice = ws.Cells(2, 3).Value
    totalVolume = 0
    
    'Loop through data
        For i = 2 To lastRow
            ticker = ws.Cells(i, 1).Value
            closePrice = ws.Cells(i, 6).Value
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
    'When ticker symbol changes
    If ticker <> ws.Cells(i + 1, 1).Value Then
        yearlyChange = closePrice - openPrice
        If openPrice <> 0 Then
            percentChange = (yearlyChange / openPrice) * 100
        Else
            percentChange = 0
        End If
        
   'Put info into new rows
    ws.Cells(outputRow, 9).Value = ticker
    ws.Cells(outputRow, 10).Value = yearlyChange
    ws.Cells(outputRow, 11).Value = percentChange
    ws.Cells(outputRow, 12).Value = totalVolume
    
    'Move on to next row
    outputRow = outputRow + 1
    
    'Start values for new ticker
    openPrice = ws.Cells(i + 1, 3).Value
    totalVolume = 0
    
    'Little Graph
     If percentChange > MaxPercentIncrease Then
        MaxPercentIncrease = percentChange
        MaxPercentIncreaseTicker = ticker
    End If
    
    If percentChange < MaxPercentDecrease Then
                    MaxPercentDecrease = percentChange
                    MaxPercentDecreaseTicker = ticker
    End If
    
    If totalVolume > MaxTotalVolume Then
                     MaxTotalVolume = totalVolume
                     MaxTotalVolumeTicker = ticker
    End If
    
End If

'Update total volume for the current ticker
totalVolume = totalVolume + volume

Next i

ws.Cells(2, 16) = MaxPercentIncreaseTicker
ws.Cells(3, 16) = MaxPercentDecreaseTicker
ws.Cells(4, 16) = MaxTotalVolumeTicker
ws.Cells(2, 17) = MaxPercentIncrease
ws.Cells(3, 17) = MaxPercentDecrease
ws.Cells(4, 17) = MaxTotalVolume

MaxPercentIncrease = 0
MaxPercentDecrease = 0
MaxTotalVolume = 0
MaxPercentIncreaseTicker = ""
MaxPercentDecreaseTicker = ""
MaxTotalVolumeTicker = ""

Next ws


End Sub
