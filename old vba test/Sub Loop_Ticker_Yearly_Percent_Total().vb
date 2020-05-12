Sub Loop_Ticker_Yearly_Percent_Total()

    'name columns
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

Dim ws_num, x As Integer

Dim Ticker_Letter As String
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Volume As LongPtr
Dim Same_Ticker As Integer

Dim Yearly_Range As Range


ws_num = ActiveWorkbook.Worksheets.Count

For x = 1 To ws_num

    'define tick letter, year change, % change, total vol

    Total_Volume = 0
    
    Same_Ticker = 1
    
    'Loop through all ticker values
    For y = 2 To 99999

        'Check for different ticker
        If Cells(y + 1, 1).Value <> Cells(y, 1).Value Then
         
            'set ticker letter
            Same_Ticker = Cells(y, 1).Value

            '
            
            'total volume
            Total_Volume = Total_Volume + Cells(y, 7).Value
            
            'next item
            Same_Ticker = Same_Ticker + 1
            
            'print to sheet
            Range("I" & Same_Ticker).Value = Ticker_Letter
            Range("L" & Same_Ticker).Value = Total_Volume
            
            
            'color changer
            Set Yearly_Range = Range("$J:$J")
            
            For Each cell In YR
                If Yearly_Range > 0 Then
                    cell.Interior.ColorIndex = 4
                ElseIf Yearly_Range < 0 Then
                    cell.Interior.ColorIndex = 3
                Else
                    cell.Interior.ColorIndex = 0
                End If
            
                    
        End If
  
    Next y
    

Next x

End Sub