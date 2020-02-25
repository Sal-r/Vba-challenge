Sub ColorPercentChange()

Dim MaxColumnI As Long
MaxColumnI = WorksheetFunction.CountA(Columns("I:I"))
For Y = 2 To MaxColumnI
    
        'color percent change
        If Cells(Y, 11).Value < 0 Then
            Cells(Y, 11).Interior.ColorIndex = 3

        ElseIf Cells(Y, 11).Value > 0 Then
            Cells(Y, 11).Interior.ColorIndex = 4

        Else
            Cells(Y, 11).Interior.ColorIndex = 0
        End If
Next Y
        
End Sub