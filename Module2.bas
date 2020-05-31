Attribute VB_Name = "Module2"
Sub stockMarket():

Dim ws As Worksheet
Set ws = Worksheets("2014")

Dim Ticker As String
Dim openVal As Double
Dim closeVal As Double
Dim stockVol As String
Dim yearChange As Double
Dim percentChange As Double

lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

tickerCounter = 2
stockCounter = 2
yearCounter = 2
percentCounter = 2

yearChange = 0
stockVol = 0

Cells(1, 12).Value = "Ticker"
Cells(1, 13).Value = "Yearly Change"
Cells(1, 14).Value = "Percent Change"
Cells(1, 15).Value = "Total Stock Volume"

For i = 2 To lastRow

    If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then

        'Gets open Value
        openVal = Cells(i, 3).Value

    End If


    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        'Gets Ticker symbol
        Ticker = Cells(i, 1).Value
        Cells(tickerCounter, 12).Value = Ticker
        tickerCounter = tickerCounter + 1

        'Gets close Value
        closeVal = Cells(i, 6).Value

        'Gets year Change
        yearChange = closeVal - openVal
        Cells(yearCounter, 13).Value = yearChange
        yearCounter = yearCounter + 1

        'Gets colors for year Change
        If yearChange >= 0 Then
            Cells(yearCounter - 1, 13).Interior.ColorIndex = 4
        Else
            Cells(yearCounter - 1, 13).Interior.ColorIndex = 3
        End If

        'Gets percent Change
        If openVal = 0 Or yearChange = 0 Then
            percentChange = 0
            Cells(percentCounter, 14).Value = percentChange
            Cells(percentCounter, 14).NumberFormat = "0.00%"
            percentCounter = percentCounter + 1

        Else
            percentChange = yearChange / openVal
            Cells(percentCounter, 14).Value = percentChange
            Cells(percentCounter, 14).NumberFormat = "0.00%"
            percentCounter = percentCounter + 1
        End If

    End If
    
    'Gets total stock Volume
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        stockVol = stockVol + Cells(i, 7).Value
        Cells(stockCounter, 15).Value = stockVol
        stockCounter = stockCounter + 1
        stockVol = 0
    Else
        stockVol = stockVol + Cells(i, 7).Value

    End If
               
    Next i
    
End Sub
