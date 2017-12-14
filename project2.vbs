Sub main():
    ' declare variables
    Dim sol_row As Long
    For Each Sheet In Worksheets
        sol_row = 2
        With Sheet
            ' put in ticker header
            .Range("I1").Value = "Ticker"
            ' put in yearly change header
            .Range("J1").Value = "Yearly Change"
            ' put in percent change header
            .Range("K1").Value = "Percent Change"
            ' put in total volume header
            .Range("L1").Value = "Total Stock Volume"
            ' put in the other ticker header
            .Range("p1").Value = "Ticker"
            ' put in value header
            .Range("q1").Value = "Value"
            ' put in greatest % increase label
            .Range("O2").Value = "Greatest % Increase"
            ' put in greatest % decrease label
            .Range("O3").Value = "Greatest % Decrease"
            ' put in greatest total volume label
            .Range("O4").Value = "Greatest Total Volume"
            ' put in new ticker value
            .Range("I2").Value = .Range("a2").Value
            ' set new yearly change to opening price
            .Range("J2").Value = .Range("C2").Value
            ' set new percent change to the same
            .Range("K2").Value = .Range("C2").Value
            ' set new total vol to vol
            .Range("L2").Value = .Range("G2").Value
            For Row = 3 To .Cells(.Rows.Count, "A").End(xlUp).Row
                If .Cells(Row, 1).Value = .Cells(Row + 1, 1).Value Then
                    ' add this row's volume to the total
                    .Cells(sol_row, "L").Value = (.Cells(sol_row, "L").Value + .Cells(Row, "G").Value)
                Else
                    ' add this row's volume to the total
                    .Cells(sol_row, "L").Value = (.Cells(sol_row, "L").Value + .Cells(Row, "G").Value)
                    ' set yearly change to current value of yearly change - current row's closing price
                    .Cells(sol_row, "J").Value = (.Cells(sol_row, "J").Value - .Cells(Row, "F").Value)
                    ' check if opening price is 0
                    If .Cells(sol_row, "K").Value = 0 Then

                    ' elif closing price > opening price
                    ElseIf .Cells(sol_row, "F").Value >= .Cells(sol_row, "K").Value Then
                        ' set percent change to (current row's closing price - current value of percent change) / current vlaue of percent change
                        .Cells(sol_row, "K").Value = ((.Cells(Row, "F").Value - .Cells(sol_row, "K").Value) / .Cells(sol_row, "K").Value)
                    Else
                        ' set percent change to (current value of percent change - current row's closing price) / current vlaue of percent change
                        .Cells(sol_row, "K").Value = ((.Cells(sol_row, "K").Value - .Cells(Row, "F").Value) / .Cells(sol_row, "K").Value)
                    End If
                    ' if yearly change > max yearly change
                    If .Cells(sol_row, "K").Value > .Range("Q2").Value Then
                        ' update max yearly change
                        .Range("Q2").Value = .Cells(sol_row, "K").Value
                        ' update ticker
                        .Range("P2").Value = .Cells(sol_row, "I").Value
                    ' else if yearly change < min yearly change
                    ElseIf .Cells(sol_row, "K").Value < .Range("Q3").Value Then
                        ' update min yearly change
                        .Range("q3").Value = .Cells(sol_row, "K").Value
                        ' update ticker
                        .Range("P3").Value = .Cells(sol_row, "I").Value
                    End If
                    ' if total volume > max total volume
                    If .Cells(sol_row, "L").Value > .Range("q4").Value Then
                        ' update max total volume
                        .Range("q4").Value = .Cells(sol_row, "L").Value
                        ' update ticker
                        .Range("p4").Value = .Cells(sol_row, "I").Value
                    End If
                    ' increment sol_row
                    sol_row = (sol_row + 1)
                    ' put in new ticker value
                    .Cells(sol_row, "I").Value = .Cells(Row + 1, "A").Value
                    ' set new yearly change to the next line's opening price
                    .Cells(sol_row, "J").Value = .Cells(Row + 1, "C").Value
                    ' set new percent change to the same
                    .Cells(sol_row, "K").Value = .Cells(Row + 1, "C").Value
                    ' shouldn't need to set new total vol to 0
                    ' .Cells(sol_row, "L").Value = 0
                End If
            Next Row
        End With
    Next Sheet
End Sub