Sub Looptest1()
    For j = 1 To Sheets.Count
        Sheets(j).Activate
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Changes"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("I2").Value = "A"
        Dim endNum As Long
        Dim firstCount As Integer
        firstCount = 2
        Dim beginValue As Double
        beginValue = Cells(2, 3).Value
        endNum = Cells(Rows.Count, "A").End(xlUp).Row
        For i = 2 To endNum
            If Cells(i, 1).Value = Cells(firstCount, 9).Value Then
                Cells(firstCount, 12).Value = Cells(firstCount, 12).Value + Cells(i, 7).Value
            Else
                Cells(firstCount + 1, 9).Value = Cells(i, 1).Value
                Cells(firstCount, 10).Value = Cells(i - 1, 6).Value - beginValue
                If beginValue = 0 Then
                    Cells(firstCount, 11).Value = 0
                Else
                    Cells(firstCount, 11).Value = Cells(firstCount, 10) / beginValue
                End If
                beginValue = Cells(i, 3).Value
                firstCount = firstCount + 1
            End If
        Next i
        Cells(firstCount, 11).Value = CDbl(Cells(firstCount, 10).Value) / CDbl(Cells(endNum, 6))
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("Q2:Q4").Value = 0
        For i = 2 To Cells(Rows.Count, "I").End(xlUp).Row
            If Cells(i, 10).Value >= 0 Then
                Cells(i, 10).Interior.ColorIndex = 4
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
            If Range("Q2").Value < Cells(i, 11).Value Then
                Range("Q2").Value = Cells(i, 11).Value
                Range("P2").Value = Cells(i, 9).Value
            End If
            If Range("Q3").Value > Cells(i, 11).Value Then
                Range("Q3").Value = Cells(i, 11).Value
                Range("P3").Value = Cells(i, 9).Value
            End If
            If Range("Q4").Value < Cells(i, 12).Value Then
                Range("Q4").Value = Cells(i, 12).Value
                Range("P4").Value = Cells(i, 9).Value
            End If
            Cells(i, 11).NumberFormat = "0.00%"
            Cells(i, 10).NumberFormat = "0.000000000"
        Next i
        Range("Q2:Q3").NumberFormat = "0.00%"
    Next j
End Sub
