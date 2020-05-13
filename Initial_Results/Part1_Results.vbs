Attribute VB_Name = "Module1"
'1. Loop through Column A and compare Below and Above cells to find the ticker differece.
'2. Generate a TotalTicker and Print it into Column J
'3. Generate a TotalChange based on the Open and Close Dates per ticker, and print it into column K.
'4. Calculate a PercentangeChange based on the Open and Close Dates per ticker, and print it into column L.
'5. Calculate a TotalStockVol per ticker, and print it into column M.

Sub VBA_HomeWork()

'Variables Declaration

Dim i As Long
Dim LastRow As Long
Dim CountTicker As Long
Dim OpenStock As Double
Dim CloseStock As Double
Dim TotalChange As Double
Dim PercentageChange As Double
Dim TotalStockVol As Double

'Variables Definition
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
CountTicker = 1
OpenStock = Range("c2").Value
TotalStockVol = 0

'Setting two decimals to PercentageChange

Columns("K").NumberFormat = "0.00%"
Cells(2, 17).NumberFormat = "0.00%"
Cells(3, 17).NumberFormat = "0.00%"

' Headers Print 1

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'The Cell Adjustment is missing

'Loop

For i = 2 To LastRow

'TotalStockVol Cum Per Ticker

    TotalStockVol = TotalStockVol + Cells(i, 7).Value

    If Cells(i, 1).Value <> Cells(i + 1, 1) Then
        CountTicker = CountTicker + 1
    
'Ticker Print
        Range("I" & CountTicker).Value = Cells(i, 1).Value
    
'Stock Annual Change
        CloseStock = Cells(i, 6).Value
        TotalChange = CloseStock - OpenStock
        Range("J" & CountTicker).Value = TotalChange
    
'Annual Percentage Change
            If OpenStock <> 0 Then
                PercentageChange = (TotalChange / OpenStock)
            Else
                PercentageChange = 0
            End If
        Range("K" & CountTicker).Value = PercentageChange

        'Color Formatting
            If PercentageChange < 0 Then
                Range("J" & CountTicker).Interior.ColorIndex = 3
            Else
                Range("J" & CountTicker).Interior.ColorIndex = 4
            End If
        
'Open Stock Next Period
        OpenStock = Cells(i + 1, 3).Value
    
'Stock Total Volume
        Range("L" & CountTicker).Value = TotalStockVol
     
'TotalStock Vol reset
        TotalStockVol = 0
    End If
Next i

Columns("j:l").AutoFit

End Sub
