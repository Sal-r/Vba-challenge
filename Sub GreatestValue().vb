Sub GreatestValue()

Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"
Cells(2, 14).Value = "Greatest % Increase"
Cells(3, 14).Value = "Greatest % Decrease"
Cells(4, 14).Value = "Greatest Total Volume"

Dim MaxColumnI As Long
MaxColumnI = WorksheetFunction.CountA(Columns("I:I"))

Dim TickerLetter2 As String
Dim BigtIncrease, BigDecrease As Double
BigIncrease = Application.WorksheetFunction.Max(Range("k:k"))
BigDecrease = Application.WorksheetFunction.Min(Range("k:k"))

Dim BigTotal As LongPtr
BigTotal = Application.WorksheetFunction.Max(Range("L:L"))

    For Z = 2 To MaxColumnI

        If Cells(Z + 1, 9).Value <> Cells(Z, 9).Value Then
            TickerLetter2 = Cells(Z, 9).Value
            
            'Cells(2, 15).Value = "Greatest % Increase"
            Cells(2, 15).Value = TickerLetter2
            Cells(2, 16).Value = BigIncrease
            'Cells(3, 15).Value = "Greatest % Decrease"
            Cells(3, 15).Value = TickerLetter2
            Cells(3, 16).Value = BigDecrease
            'Cells(4, 15).Value = "Greatest Total Volume"
            Cells(4, 15).Value = TickerLetter2
            Cells(4, 16).Value = BigTotal
            
        End If
        
    Next Z
    
End Sub
