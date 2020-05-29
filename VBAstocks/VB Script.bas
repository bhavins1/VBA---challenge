Attribute VB_Name = "Module1"
Sub StockAnalysis()

Dim ws As Worksheet
For Each ws In Worksheets
ws.Activate


'Creating header
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"

ws.Range("i1:l1", "n1:p1").Borders(xlEdgeBottom).LineStyle = xlContinuous
ws.Range("i1:l1", "n1:p1").Borders(xlEdgeBottom).Weight = xlThick


Dim Ticker As String
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalVolume As LongLong
Dim BegPrice As Double
Dim EndPrice As Double

TotalVolume = 0
BegPrice = 0

Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Dim i As Long

Dim k As Long

For i = 2 To lastrow

    If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
                
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
        Do Until BegPrice > 0
                    
            For k = i To lastrow
                If BegPrice = 0 Then
                BegPrice = BegPrice + Cells(k, 6).Value
                End If
            Next k
                
        Loop
        
        Else
        'Display Ticker
            Ticker = ws.Cells(i, 1).Value
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            EndPrice = ws.Cells(i, 6).Value
            ws.Range("I" & Summary_Table_Row).Value = Ticker

        'Display total volume, yearly change, percent change
            ws.Range("L" & Summary_Table_Row).Value = TotalVolume
            YearlyChange = EndPrice - BegPrice
            ws.Range("J" & Summary_Table_Row).Value = YearlyChange
            PercentChange = YearlyChange / BegPrice
            ws.Range("K" & Summary_Table_Row).Value = PercentChange
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
      ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Total Volume and Beginning Price
            TotalVolume = 0
            BegPrice = 0
                        
    End If
            
Next i


Dim j As Integer

Lastrowcolor = Cells(Rows.Count, 10).End(xlUp).Row

    For j = 2 To Lastrowcolor

        If ws.Cells(j, 10).Value < 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 3
            Else
            ws.Cells(j, 10).Interior.ColorIndex = 4
        End If
        
    Next j
    
    
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"
ws.Columns("N").AutoFit

Dim Max As Double
Dim Min As Double
Dim MaxVol As LongLong
Dim TickerMax As String
Dim TickerMin As String
Dim TickerVol As String

Max = WorksheetFunction.Max(Columns("K"))
Min = WorksheetFunction.Min(Columns("K"))
MaxVol = WorksheetFunction.Max(Columns("L"))


Dim n As Integer

lastrow2 = Cells(Rows.Count, 9).End(xlUp).Row

For n = 2 To lastrow2

    If Cells(n, 11) = Max Then
        TickerMax = ws.Cells(n, 9).Value
        ws.Range("O2").Value = TickerMax
        ws.Range("P2").Value = Max
    End If
    
    If Cells(n, 11) = Min Then
        TickerMin = ws.Cells(n, 9).Value
        ws.Range("O3").Value = TickerMin
        ws.Range("P3").Value = Min
        ws.Range("P2:P3").NumberFormat = "0.00%"
    End If
    
    If Cells(n, 12) = MaxVol Then
        TickerVol = ws.Cells(n, 9).Value
        ws.Range("O4").Value = TickerVol
        ws.Range("P4").Value = MaxVol
    End If
    
Next n
    
ws.Columns("P").AutoFit
      

Next ws


End Sub
