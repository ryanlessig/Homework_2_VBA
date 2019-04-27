Sub stockChecker_Easy()

    'Declare all variables
    Dim tickerSymbol As String
    Dim tickerSymbolCurrent As String
    Dim tickerSymbolNext As String
    Dim volumeTotal As Double
    Dim totalTableRow As Integer
    Dim maxDataRow As Long
    Dim i As Long

    'Initialize volumeTotal equal to 0. It wil be reset again in each loop.
    volumeTotal = 0

    'Loop through all active worksheets within the workbook
    For Each currentWS In Worksheets

            'Set the variable named maxRow equal to the count of rows with data in them.
            maxDataRow = currentWS.Cells(Rows.Count, 1).End(xlUp).Row

            'Set the variable named totalTableRow for the total table equal to row 2 for each sheet. This resets the totalTableRow counter on each sheet.
            totalTableRow = 2

            'Set the column headings for the easy solution to their appropriate names.
            currentWS.Range("I1").Value = "Ticker"
            currentWS.Range("J1").Value = "Total Stock Volume"

        'On each sheet, set the counter i = 2, then start the for loop which will go from i to the last row in the table or maxDataRow as defined above. i = 2 refers to starting at row 2 for this loop.
        For i = 2 To maxDataRow

            'Set the variable tickerSymbolCurrent equal to the value in cell A2. Then set the variable tickerSymbolNext equal to the value in cell A3.
            tickerSymbolCurrent = currentWS.Cells(i, 1).Value
            tickerSymbolNext = currentWS.Cells(i + 1, 1).Value

            'Check if cell A2 is NOT equal to cell A3. If that is true, enter the then statement. If that is false, go to the else statement. Next time the loop runs it'll check if A3 is NOT equal to A4, and so on, for every sheet.
            If tickerSymbolCurrent <> tickerSymbolNext Then

                'If the value A2 is NOT equal to A3, then set the variable tickerSymbol equal to the variable tickerSymbolCurrent (as defined above, it'll reference a cells value). And set tickerSymbolCurrent equal to tickerSymbolNext. This increments the variables as the loop goes down the rows.
                tickerSymbol = tickerSymbolCurrent
                tickerSymbolCurrent = tickerSymbolNext
                
                'Set volumeTotal equal to volumeTotal (which starts at 0) plus the value of the current cell in the loop from column G. This creates a rolling sum by ticker.
                volumeTotal = volumeTotal + currentWS.Cells(i, 7).Value

                'Set the values in I through J equal to all the variables defined in the loop.
                currentWS.Range("I" & totalTableRow).Value = tickerSymbol
                currentWS.Range("J" & totalTableRow).Value = volumeTotal

                'Reset volumeTotal equal to 0 for the next ticker name. Go to the next row in the totalTable and set the opening price to the current row plus 1 (which is the first opening price for the next ticker sybmol.)
                volumeTotal = 0
                totalTableRow = totalTableRow + 1
         
            Else

                volumeTotal = volumeTotal + currentWS.Cells(i, 7).Value

            End If

        Next i

        'Format each sheet. Make the header row bold, center the text in all cells and autosize all columns to fit the corresponding data inside them.
        currentWS.Range("A1:J1").Font.Bold = True
        currentWS.Range("A:J").HorizontalAlignment = xlCenter
        currentWS.Columns.AutoFit

    'After you run all the scripts above, go to the next sheet and run them all again until you get through every worksheet in the workbook.
    Next currentWS
    
End Sub

Sub clear()

'Loop through each sheet and clear all the calculated values that the sub above runs. Easier than doing it sheet by sheet. Then you can run the script again to verify it all works.
For Each currentWS In Worksheets

    currentWS.Range("I:Q").clear
    
Next currentWS

End Sub