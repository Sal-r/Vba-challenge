Sub Ticker_Yearly_Percent_Total()

'Name Columns & Rows
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Range("J:J").NumberFormat = "$#,##0.00"
Range("K:K").NumberFormat = "0.00%"

'define tick letter, year change, % change, total vol, first open, last close, last y, maxcolumns
Dim TickerLetter As String

Dim YearlyChange As Double

Dim PercentChange As Double

Dim TotalValue As LongPtr
TotalValue = 0

Dim FirstOpen As Double
FirstOpen = 0

Dim LastClose As Double
LastClose = 0

Dim PreviousX As Long
PreviousX = 2

Dim MaxColumnA As Long
MaxColumnA = WorksheetFunction.CountA(Columns("A:A"))

    
Dim SameTicker As Double
SameTicker = 1
    
    'Loop through all ticker values
    For x = 2 To MaxColumnA

        'Check for different ticker
        If Cells(x + 1, 1).Value <> Cells(x, 1).Value Then
            
            'set ticker letter
            TickerLetter = Cells(x, 1).Value
            
            'FirstOpen
            FirstOpen = Cells(PreviousX, 3).Value
            LastClose = Cells(x, 6).Value
                        
            'Yearly Change = (lastclose - firstopen)
            YearlyChange = LastClose - FirstOpen
            
            'percent change = ((lastclose - firstopen)/firstopen)
            PercentChange = ((LastClose - FirstOpen) / FirstOpen)
            
            'total volume
            TotalValue = TotalValue + Cells(x, 7).Value
            
            'next item
            SameTicker = SameTicker + 1
            
            'print to sheet
            Range("I" & SameTicker).Value = TickerLetter
            Range("J" & SameTicker).Value = YearlyChange
            Range("K" & SameTicker).Value = PercentChange
            Range("L" & SameTicker).Value = TotalValue
                        
            TotalValue = 0
            'save last y
            PreviousX = x + 1
            
        Else
        
            TotalValue = TotalValue + Cells(x, 7).Value
            
        End If
         
    Next x

End Sub