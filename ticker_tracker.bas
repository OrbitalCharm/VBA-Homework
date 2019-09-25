Sub ticker_tracker()

'Variable definitions

Dim ticker As String

Dim year_change As Double
Dim percent_change As Double
Dim total_volume As Double

Dim y_o As Double
Dim y_f As Double
total_volume = 0

Dim summary_row As Integer
    summary_row = 2
Dim last_row As Long
    last_row = Cells(Rows.Count, 1).End(xlUp).Row

Dim maxPerc As Double
Dim minPerc As Double
Dim maxVol As Double
Dim maxTicker As String
maxPerc = 0
minPerc = 0
maxVol = 0

'Header lablels
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

For i = 2 To last_row

    y_o = Cells(i, 3).Value

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

    ticker = Cells(i, 1).Value
    y_f = Cells(i, 6).Value
    year_change = y_f - y_o
    percent_change = (year_change / y_o) * 100
    If y_o <> 0 Then
    percent_change = (year_change / y_o) * 100
    ElseIf y_f = 0 Then
    percent_change = 0
    yearl_change = 0
    End If
    total_volume = total_volume + Cells(i, 7).Value


    Range("I" & summary_row).Value = ticker
    Range("J" & summary_row).Value = year_change
    Range("K" & summary_row).Value = percent_change
    Range("L" & summary_row).Value = total_volume
    total_volume = 0

    If (yearlyChange > 0) Then
    Range("J" & summary_row).Interior.ColorIndex = 4
    ElseIf (year_change <= 0) Then
    Range("J" & summary_row).Interior.ColorIndex = 3

    End If

    summary_row = summary_row + 1
    y_o = 0
    y_f = 0
    y_o = Cells(i + 1, 3).Value

    If percent_change > maxPerc Then
    maxPerc = percent_change
    maxTicker = ticker
    End If
    If percent_change < minPerc Then
    minPerc = percent_change
    minTicker = ticker
    End If
    If total_Vol > maxVol Then
    maxVol = total_Vol
    maxVol_ticker = ticker
    End If

    total_Vol = 0
            
    ElseIf Cells(i, 1).Value = Cells(i + 1, 1).Value Then
           
    total_Vol = total_Vol + Cells(i, 7).Value
    If y_o = 0 Then
    y_o = Cells(i + 1, 3).Value
    End If
End If

Next i

Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    
    Dim last_row_sumTable As Long
    last_row_sumTable = Cells(Rows.Count, 9).End(xlUp).Row

    
    maxPerc = 0
    minPerc = 0
    maxVol = 0
    For j = 2 To last_row_sumTable
       If Cells(j, 11).Value > maxPerc Then
        maxPerc = Cells(j, 11).Value
        maxTicker = Cells(j, 9).Value
        End If
        If Cells(j, 11).Value < minPerc Then
        minPerc = Cells(j, 11).Value
        minTicker = Cells(j, 9).Value
        End If
        If Cells(j, 12).Value > maxVol Then
        maxVol = Cells(j, 12).Value
        maxVol_ticker = Cells(j, 9).Value
        End If
    
    Next j
    Range("P2").Value = maxTicker
    Range("Q2").Value = Str(maxPerc) & "%"
    Range("P3").Value = minTicker
    Range("Q3").Value = Str(minPerc) & "%"
    Range("P4").Value = maxVol_ticker
    Range("Q4").Value = maxVol
    
    Columns("I:Q").AutoFit



End Sub
