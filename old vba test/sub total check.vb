Sub Loop_Ticker_Yearly_Percent_Total()

    'name columns
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

Dim ws_num, x As Integer

ws_num = ActiveWorkbook.Worksheets.Count

For x = 1 To ws_num

    'define tick letter, year change, % change, total vol
    Dim TL As String
    
    Dim YC As Double
    
    Dim PC As Double
    
    Dim TV As LongPtr
    TV = 0
    
    Dim ST As Integer
    ST = 1
    
    'Loop through all ticker values
    For y = 2 To 99999

        'Check for different ticker
        If Cells(y + 1, 1).Value <> Cells(y, 1).Value Then
         
            'set ticker letter
            TL = Cells(y, 1).Value
            
            'total volume
            TV = TV + Cells(y, 7).Value
            
            'next item
            ST = ST + 1
            
            'print to sheet
            Range("I" & ST).Value = TL
            Range("L" & ST).Value = TV
            
            
            'color changer
            Dim YR As Range
            Set YR = Range("$J:$J")
            
            For Each cell In YR
                If YR > 0 Then
                    cell.Interior.ColorIndex = 4
                ElseIf YR < 0 Then
                    cell.Interior.ColorIndex = 3
                Else
                    cell.Interior.ColorIndex = 0
                End If
            
                    
        End If
  
    Next y
    

Next x

End Sub