'Instructions

'Create a script that will loop through all the stocks for one year and output the following information.
'   The ticker symbol.
'   Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'   The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'   The total stock volume of the stock.
'You should also have conditional formatting that will highlight positive change in green and negative change in red.

'CHALLENGES

'1. Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume". 
'2. Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.

'Other Considerations

'   Use the sheet alphabetical_testing.xlsx while developing your code. This data set is smaller and will allow you to test faster. Your code should run on this file in less than 3-5 minutes.
'   Make sure that the script acts the same on each sheet. The joy of VBA is to take the tediousness out of repetitive task and run over and over again with a click of the button.
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Sub VBA_of_Wall_Street():

    'macro to loop through all pages of the workbook
    Dim ws As Worksheet
    For Each ws in Worksheets

        'Sub Primary_Analysis():

            'declare variables
            Dim i, posTracker As integer
            Dim ticker As string
            Dim priceOpen, priceClose, yearlyChange, percentChange As double
            Dim stockVolume As LongLong

            'define variables
            posTracker = 2
            stockVolume = 0
                'setting first orice on sheet to initial open price
            priceOpen = ws.Range("C2")
                'print initial opening price of first ticker
                'Cells(posTracker,13).Value = priceOpen
            priceClose = 0

            'creating table headers
            ws.Cells(1,9).Value = "Ticker"
            ws.Cells(1,10).Value = "Yearly Change"
            ws.Cells(1,11).Value = "Percent Change"
            ws.Cells(1,12).Value = "Total Stock Volume"

            'calculate the last row of the worksheet
            Dim lastRow As long
            lastRow = ws.Cells(1,1).End(xlDown).Row

            'for statement to loop through rows of data from row 2 to last row of worksheet
            for i = 2 to lastRow

                'if statement to check when the ticket symbol changes
                if ws.Cells(i,1).Value <> ws.Cells(i+1,1).Value Then

                    'defining table variables
                    ticker = ws.Cells(i,1).Value
                    'order of variable declaration matters - have to calculate yearly change before the priceOpen changes
                    priceClose = ws.Cells(i,6).Value
                    yearlyChange = priceClose - priceOpen
                    'if statement to work around divide by 0 error when calculating percentChange and priceClose = 0
                    if priceClose = 0 and priceOpen = 0 Then
                        percentChange = 0                    
                    Else
                        percentChange = ((priceClose - priceOpen)/ ABS(priceClose))
                    End if
                    priceOpen = ws.Cells(i+1,3).Value

                        'fill in table based on ticker symbol
                        'display ticker code on column I
                        ws.Cells(posTracker,9).Value = ticker
                        'calculate and display yearly change on column J
                        ws.Cells(posTracker,10).Value = yearlyChange
                        'Cells(posTracker+1,13).Value = priceOpen
                        'Cells(posTracker,14).Value = priceClose
                        'display percent change on column K
                        ws.Cells(posTracker,11).Value = percentChange
                        'display stock volume on column L
                        ws.Cells(posTracker,12).Value = stockVolume + ws.Cells(i,7).Value
                    
                    'moving down 1 row on table when ticker changes
                    posTracker = posTracker + 1

                    'resetting table trackers
                    yearlyChange = 0
                    percentChange = 0
                    stockVolume = 0
                                
                'else statement to keep track of totals and changes per ticker
                Else

                    'running total for stock volume
                    stockVolume = stockVolume + ws.Cells(i,7).Value

                'end if statement to check when ticker symbol changes
                End if

            'end for statement to loop through rows of data
            Next i

            'formatting of cells
            'change percent change column to 2 decimal places
            ws.Range("K1:K" & Cells(1,11).End(xlDown).Row).NumberFormat = "0.00%"

            'bold table headers
            'outline table cell boxes

            'autofit columns of worksheet table
            ws.Columns("J:L").Autofit
            'change colour of yearly change cells based on +ive or -ive result 
            lastRow = ws.Cells(1,10).End(xlDown).Row
            for i = 2 to lastRow
                'if statement to loop throuch cells and check value of cells to determine color
                if ws.Cells(i,10).Value > 0 Then
                    ws.Cells(i,10).Interior.ColorIndex = 4
                Elseif ws.Cells(i,10).Value < 0 Then
                    ws.Cells(i,10).Interior.ColorIndex = 3
                else
                    ws.Cells(i,10).Interior.ColorIndex = 15
                End if
            Next
            
        'end primary analysis
        'End Sub
        
        'Sub Secondary_Analysis():

            'declare variables
            Dim maxPercentIncrease, MaxPercentDecrease As double
            Dim maxStockVolume As LongLong

            'define variables
            maxPercentIncrease = 0
            MaxPercentDecrease = 0
            maxStockVolume = 0
            
            'creating table headers
            ws.Cells(1,16).Value = "Ticker"
            ws.Cells(1,17).Value = "Value"
            ws.Cells(2,15).Value = "Greatest % Increase"
            ws.Cells(3,15).Value = "Greatest % Decrease"
            ws.Cells(4,15).Value = "Greatest Total Volume"

            'calculate the last row of primary analysis table
            lastRow = ws.Cells(9,1).End(xlDown).Row

            'for statement to loop through primary analysis table
            for i = 2 to lastRow
               
                'if statement to calculate the greatest % increase
                if ws.Cells(i,11).Value > maxPercentIncrease Then
                    
                    maxPercentIncrease = ws.Cells(i,11).Value
                    'print ticker values to table
                    ws.Cells(2,16).Value = ws.Cells(i,9).Value
                    'print greatest % increase value to table
                    ws.Cells(2,17).Value = maxPercentIncrease

                'ending if statement for greatest % increase
                End if

                'if statement to calculate the greatest % decrease
                if ws.Cells(i,11).Value < maxPercentDecrease Then
                    
                    maxPercentDecrease = ws.Cells(i,11).Value
                    'print ticker values to table
                    ws.Cells(3,16).Value = ws.Cells(i,9).Value
                    'print greatest % decrease value to table
                    ws.Cells(3,17).Value = maxPercentDecrease

                'ending if statement for greatest % decrease
                End if
                
                'if statement to calculate the greatest volume traded
                if ws.Cells(i,12).Value > maxStockVolume Then
                    
                    maxStockVolume = ws.Cells(i,12).Value
                    'print ticker values to table
                    ws.Cells(4,16).Value = ws.Cells(i,9).Value
                    'print greatest volume traded to table
                    ws.Cells(4,17).Value = maxStockVolume

                'ending if statement for greatest volume traded
                End if

            'end for statement
            Next i

            'formatting secondary analysis table
            ws.Range("Q2:Q3").NumberFormat = "0.00%"
            'auto fit colums of secondary analysis table
            ws.Columns("O").Autofit
            ws.Columns("Q").Autofit
        
        'end secondary analysis
        'End Sub

    'end worksheet loop
    Next ws

End Sub
