Attribute VB_Name = "Module1"
'1. Loop through Column A and compare Below and Above cells to find the ticker differece.
'2. Generate a TotalTicker and Print it into Column J
'3. Generate a TotalChange based on the Open and Close Dates per ticker, and print it into column K.
'4. Calculate a PercentangeChange based on the Open and Close Dates per ticker, and print it into column L.
'5. Calculate a TotalStockVol per ticker, and print it into column M.
'6. Create Arrays and assigned calculated PercentageChange and TotalStockVol in each one. Its necessary to create the last row for each column in order to define the limit
'7. Sort the Arrays from the highest to the lowest value. Use the bubble sort algorithm.
'8. Seek for the Tikcer that corresponds to the first and last value for each sorted arrays. In this step I need to compare the sorted "Imaginary" Array with Column K and Colum L, and Print Ticker from Column I.
'9. Adjust the code to can run all existing sheets.

Sub VBA_HomeWork()

'ws Declaration

Dim ws As Variant

'Loop through all Worksheets

For Each ws In Worksheets

'Variables Declaration

Dim i As Long
Dim LastRow As Long
Dim CountTicker As Long
Dim OpenStock As Double
Dim CloseStock As Double
Dim TotalChange As Double
Dim PercentageChange As Double
Dim TotalStockVol As Double
Dim x As Long, y As Long
Dim Temp As Variant
Dim Percentage_range() As Variant
Dim TotalVolume_range() As Variant
Dim c As Long
Dim d As Integer
Dim z As Integer
Dim k As Integer
Dim p As Integer
Dim col_sizepercent As Long
Dim col_sizeVolume As Long


'Variables Definition
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
CountTicker = 1
OpenStock = ws.Range("c2").Value
TotalStockVol = 0

'Setting two decimals to PercentageChange

ws.Columns("K").NumberFormat = "0.00%"
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 17).NumberFormat = "0.00%"

' Headers Print 1

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'The Cell Adjustment is missing

'Loop

For i = 2 To LastRow

'TotalStockVol Cum Per Ticker

    TotalStockVol = TotalStockVol + ws.Cells(i, 7).Value

    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1) Then
        CountTicker = CountTicker + 1
    
'Ticker Print
        ws.Range("I" & CountTicker).Value = ws.Cells(i, 1).Value
    
'Stock Annual Change
        CloseStock = ws.Cells(i, 6).Value
        TotalChange = CloseStock - OpenStock
        ws.Range("J" & CountTicker).Value = TotalChange
    
'Annual Percentage Change
            If OpenStock <> 0 Then
                PercentageChange = (TotalChange / OpenStock)
            Else
                PercentageChange = 0
            End If
        ws.Range("K" & CountTicker).Value = PercentageChange

        'Color Formatting
            If PercentageChange < 0 Then
                ws.Range("J" & CountTicker).Interior.ColorIndex = 3
            Else
                ws.Range("J" & CountTicker).Interior.ColorIndex = 4
            End If
        
'Open Stock Next Period
        OpenStock = ws.Cells(i + 1, 3).Value
    
'Stock Total Volume
        ws.Range("L" & CountTicker).Value = TotalStockVol
     
'TotalStock Vol reset
        TotalStockVol = 0
    End If
Next i

'Headers Print 2

ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

' Maximum, Minimum Percentage Change + Maximum Total Volume + Ticker Print

'Last Row Percentage Definition + Total Stock Volume

col_sizepercent = ws.Cells(Rows.Count, "K").End(xlUp).Row
col_sizeVolume = ws.Cells(Rows.Count, "L").End(xlUp).Row

'Set Array to col_sizepercent

ReDim Percentage_range(col_sizepercent)
ReDim TotalVolume_range(col_sizeVolume)

' Percentage & Total Volume Array Creation

For c = 1 To col_sizepercent
    Percentage_range(c) = ws.Cells(c, "K").Value
Next c

For d = 1 To col_sizeVolume
    TotalVolume_range(d) = ws.Cells(d, "L").Value
Next d

' Percentage and Total Volume Change Array Sorting

For x = LBound(Percentage_range) + 1 To UBound(Percentage_range) - 1
    For y = x + 1 To UBound(Percentage_range)
        If Percentage_range(x) < Percentage_range(y) Then
            Temp = Percentage_range(y)
            Percentage_range(y) = Percentage_range(x)
            Percentage_range(x) = Temp
        End If
    Next y
Next x

For z = LBound(TotalVolume_range) + 1 To UBound(TotalVolume_range) - 1
    For k = z + 1 To UBound(TotalVolume_range)
        If TotalVolume_range(z) < TotalVolume_range(k) Then
            Temp = TotalVolume_range(k)
            TotalVolume_range(k) = TotalVolume_range(z)
            TotalVolume_range(z) = Temp
        End If
    Next k
Next z

'--------------------------------------------
'For a = 1 To col_sizepercent
'   Cells(a, 13).Value = Percentage_range(a)
'Next a
'For a = 1 To col_sizeVolume
'   Cells(a, 14).Value = TotalVolume_range(a)
'Next a
'--------------------------------------------

'Percentage Min and Max Print + Max Total Stock Volume

ws.Cells(2, 17).Value = Percentage_range(2)
ws.Cells(3, 17).Value = Percentage_range(col_sizepercent)
ws.Cells(4, 17).Value = TotalVolume_range(2)
ws.Range("Q4").NumberFormat = "0.0000E+00"

' Percentage and Total Stock Volume Ticker Print Max

For c = 1 To col_sizepercent
    If ws.Cells(c, 11).Value = Percentage_range(2) Then
        ws.Cells(2, 16).Value = ws.Cells(c, 9).Value
    Exit For
    End If
Next c

For p = 1 To col_sizeVolume
    If ws.Cells(p, 12).Value = TotalVolume_range(2) Then
        ws.Cells(4, 16).Value = ws.Cells(p, 9).Value
    Exit For
    End If
Next p

'Ticker Print Percentage Min

For c = 1 To col_sizepercent
    If ws.Cells(c, 11).Value = Percentage_range(col_sizepercent) Then
        ws.Cells(3, 16).Value = ws.Cells(c, 9).Value
    Exit For
    End If
Next c

ws.Columns("j:q").AutoFit

Next ws


End Sub
