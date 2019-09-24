
Sub stocks()

    For Each ws In Worksheets
    
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        Dim tickerName As String
        Dim lastRow As Long
        Dim totaltickerVolume As Double
        totaltickerVolume = 0
        Dim summarytableRow As Long
        summarytableRow = 2
        Dim yearlyOpen As Double
        Dim yearlyClose As Double
        Dim yearlyChange As Double
        Dim previousAmount As Long
        previousAmount = 2
        Dim percentChange As Double
        Dim greatestIncrease As Double
        greatestIncrease = 0
        Dim greatestDecrease As Double
        greatestDecrease = 0
        Dim lastRowValue As Long
        Dim greatesttotalVolume As Double
        greatesttotalVolume = 0

        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastRow

          totaltickerVolume = totaltickerVolume + ws.Cells(i, 7).Value
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            
           tickerName = ws.Cells(i, 1).Value
           ws.Range("I" & summarytableRow).Value = tickerName
           ws.Range("L" & summarytableRow).Value = totaltickerVolume
           totaltickerVolume = 0
            yearlyOpen = ws.Range("C" & previousAmount)
            yearlyClose = ws.Range("F" & i)
            yearlyChange = yearlyClose - yearlyOpen
            ws.Range("J" & summarytableRow).Value = yearlyChange

         ' Percent Change
            If yearlyOpen = 0 Then
            percentChange = 0
            Else
            yearlyOpen = ws.Range("C" & previousAmount)
            percentChange = yearlyChange / yearlyOpen
            End If
        
            ws.Range("K" & summarytableRow).NumberFormat = "0.00%"
            ws.Range("K" & summarytableRow).Value = percentChange
           ' Conditional Formatting
           If ws.Range("J" & summarytableRow).Value >= 0 Then
           ws.Range("J" & summarytableRow).Interior.ColorIndex = 4
             Else
            ws.Range("J" & summarytableRow).Interior.ColorIndex = 3
                End If

           summarytableRow = summarytableRow + 1
           previousAmount = i + 1
                
                End If
            Next i

            ' Greatest Section % Increase, Decrease and  Total Volume
            lastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
            For i = 2 To lastRow
                If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
           ws.Range("Q2").Value = ws.Range("K" & i).Value
           ws.Range("P2").Value = ws.Range("I" & i).Value
                End If

          If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
             ws.Range("Q3").Value = ws.Range("K" & i).Value
             ws.Range("P3").Value = ws.Range("I" & i).Value
               End If

           If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
             ws.Range("Q4").Value = ws.Range("L" & i).Value
             ws.Range("P4").Value = ws.Range("I" & i).Value
                End If

            Next i
    
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Columns("I:Q").AutoFit

    Next ws

End Sub

Sub reset_button()

For Each ws In Worksheets
        Dim WorksheetName As String
        WorksheetName = ws.Name
        
        Sheets(ws.Name).Select

Columns("I:Q").Select
Selection.Clear
Columns("I:Q").EntireColumn.AutoFit
    Cells(1, 1).Select
    
        Next ws
    
End Sub
