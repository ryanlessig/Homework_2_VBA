Sub stockChecker_Hard()

    'Declare all variables
    Dim tickerSymbol As String
    Dim tickerSymbolCurrent As String
    Dim tickerSymbolNext As String
    Dim volumeTotal As Double
    Dim totalTableRow As Integer
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim maxDataRow As Long
    Dim maxTotalRow As Long
    Dim i As Long
    Dim j As Integer
    Dim k As Integer

    'Initialize volumeTotal equal to 0. It wil be reset again in each loop.
    volumeTotal = 0

    'Loop through all active worksheets within the workbook
    For Each currentWS In Worksheets

            'Set the variable named maxRow equal to the count of rows with data in them.
            maxDataRow = currentWS.Cells(Rows.Count, 1).End(xlUp).Row

            'Set the variable named totalTableRow for the total table equal to row 2 for each sheet. This resets the totalTableRow counter on each sheet.
            totalTableRow = 2
        
            'Set the variable named openingPrice equal to the second row of the third column (C2) for each sheet. The variable will increment later in the routine. This resets the openingPrice counter on each sheet.
            openingPrice = currentWS.Cells(2, 3).Value

            'Set the column headings for the easy solution to their appropriate names.
            currentWS.Range("I1").Value = "Ticker"
            currentWS.Range("L1").Value = "Total Stock Volume"

            'Set the new column headings for the moderate solution to their appropriate names.
            currentWS.Range("J1").Value = "Yearly Change"
            currentWS.Range("K1").Value = "Percent Change"

            'Set the new column / cell headings for the hard solution to their appropriate names.
            currentWS.Range("P1").Value = "Ticker"
            currentWS.Range("Q1").Value = "Value"
            currentWS.Range("O2").Value = "Greatest % Increase"
            currentWS.Range("O3").Value = "Greatest % Decrease"
            currentWS.Range("O4").Value = "Greatest Total Volume"

            'Format the Greatest % increase and decrease cells to the percentage number format
            currentWS.Range("Q3").NumberFormat = "0.00%"
            currentWS.Range("Q2").NumberFormat = "0.00%"

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
                
                'Store the last instance of i in this loop as the closingPrice from column F.
                closingPrice = currentWS.Cells(i, 6).Value
                
                'Now that we have the closingPrice and openingPrice, do the math to find the yearly change.
                yearlyChange = closingPrice - openingPrice

                'Check if closingPrice or openingPrice is equal to 0. If either are, then set percentChange = 0. Else if openingPrice equals 0 and closing price does not equal 0, the percent change is 100%. Else if neither openingPrice or closingPrice equal 0, do the math to find the percent change.
                If openingPrice = 0 and closingPrice = 0 Then

                    percentChange = 0
                
                Elseif openingPrice = 0 and closingPrice <> 0 then

                    percentChange = 1
                
                Elseif openingPrice <> 0 and closingPrice <> 0 then

                    percentChange = ((closingPrice - openingPrice) / openingPrice)

                End If

                'Set volumeTotal equal to volumeTotal (which starts at 0) plus the value of the current cell in the loop from column G. This creates a rolling sum by ticker.
                volumeTotal = volumeTotal + currentWS.Cells(i, 7).Value

                'Set the values in I through L equal to all the variables defined in the loop.
                currentWS.Range("I" & totalTableRow).Value = tickerSymbol
                currentWS.Range("J" & totalTableRow).Value = yearlyChange
                currentWS.Range("K" & totalTableRow).Value = percentChange
                currentWS.Range("L" & totalTableRow).Value = volumeTotal

                'Reset volumeTotal equal to 0 for the next ticker name. Go to the next row in the totalTable and set the opening price to the current row plus 1 (which is the first opening price for the next ticker sybmol.)
                volumeTotal = 0
                totalTableRow = totalTableRow + 1
                openingPrice = currentWS.Cells(i + 1, 3).Value               

            Else

                volumeTotal = volumeTotal + currentWS.Cells(i, 7).Value

            End If

        Next i

        'Find the last row count for the total table for the current worksheet.
        maxTotalRow = currentWS.Cells(Rows.Count, 9).End(xlUp).Row

        'Format column K as a percentage for cells K2 through the last row in the total table.
        currentWS.Range("K2:K" & maxTotalRow).NumberFormat = "0.00%"
        
            'Loop through the total table to find if the cell is greater than or equal to 0 and color it green. If the value is less than 0, color it red.
            For j = 2 To maxTotalRow

                If currentWS.Cells(j, 10).Value >= 0 Then

                    currentWS.Cells(j, 10).Interior.ColorIndex = 10

                ElseIf currentWS.Cells(j, 10).Value < 0 Then

                    currentWS.Cells(j, 10).Interior.ColorIndex = 3
                
                End If

            Next j

            'Loop through the total table gain to find the min and max percent change in K and the max in L and populate the greatest table in O,P and Q.
            For k = 2 To maxTotalRow

                If currentWS.Cells(k, 11).Value = Application.WorksheetFunction.Max(currentWS.Range("K2:K" & maxTotalRow)) Then

                    currentWS.Range("P2").Value = currentWS.Cells(k, 9).Value
                    currentWS.Range("Q2").Value = currentWS.Cells(k, 11).Value

                ElseIf currentWS.Cells(k, 11).Value = Application.WorksheetFunction.Min(currentWS.Range("K2:K" & maxTotalRow)) Then
                    
                    currentWS.Range("P3").Value = currentWS.Cells(k, 9).Value
                    currentWS.Range("Q3").Value = currentWS.Cells(k, 11).Value

                ElseIf currentWS.Cells(k, 12).Value = Application.WorksheetFunction.Max(currentWS.Range("L2:L" & maxTotalRow)) Then

                    currentWS.Range("P4").Value = currentWS.Cells(k, 9).Value
                    currentWS.Range("Q4").Value = currentWS.Cells(k, 12).Value

                End If

            Next k

        'Format each sheet. Make the header row bold, center the text in all cells and autosize all columns to fit the corresponding data inside them.
        currentWS.Range("A1:Q1").Font.Bold = True
        currentWS.Range("A:Q").HorizontalAlignment = xlCenter
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