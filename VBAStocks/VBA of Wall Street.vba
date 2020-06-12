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

    'Sub Primary_Analysis():

        'declare variables
        Dim i, posTracker As integer
        Dim ticker As string
        Dim priceOpen, priceClose, yearlyChange, percentChange, sotckVolume As double

        'define variables
        posTracker = 2
        sotckVolume = 0
            'setting first orice on sheet to initial open price
        priceOpen = Range("C2")
            'print initial opening price of first ticker
            'Cells(posTracker,13).Value = priceOpen
        priceClose = 0


        'creating table headers
        Cells(1,9).Value = "Ticker"
        Cells(1,10).Value = "Yearly Change"
        Cells(1,11).Value = "Percent Change"
        Cells(1,12).Value = "Total Stock Volume"


        'calculate the last row of the worksheet
        Dim lastRow As long
        lastRow = Cells(1,1).End(xlDown).Row

        'for statement to loop through rows of data from row 2 to last row of worksheet
        for i = 2 to lastRow

            'if statement to check when the ticket symbol changes
            if Cells(i,1).Value <> Cells(i+1,1).Value Then

                'defining table variables
                ticker = Cells(i,1).Value
                'order of variable declaration matters - have to calculate yearly change before the priceOpen changes
                priceClose = Cells(i,6).Value
                yearlyChange = priceClose - priceOpen
                percentChange = ((priceClose - priceOpen)/ ABS(priceClose))
                priceOpen = Cells(i+1,3).Value

                    'fill in table based on ticker symbol
                    'display ticker code on column I
                    Cells(posTracker,9).Value = ticker
                    'calculate and display yearly change on column J
                    Cells(posTracker,10).Value = yearlyChange
                    'Cells(posTracker+1,13).Value = priceOpen
                    'Cells(posTracker,14).Value = priceClose
                    'display percent change on column K
                    Cells(posTracker,11).Value = percentChange
                    'display stock volume on column L
                    Cells(posTracker,12).Value = sotckVolume + Cells(i,7).Value
                
                'moving down 1 row on table when ticker changes
                posTracker = posTracker + 1

                'resetting table trackers
                yearlyChange = 0
                percentChange = 0
                sotckVolume = 0
                            
            'else statement to keep track of totals and changes per ticker
            Else

                'running total for stock volume
                sotckVolume = sotckVolume + Cells(i,7).Value

            'end if statement to check when ticker symbol changes
            End if

        'end for statement to loop through rows of data
        Next i

        'formatting of cells
        'change percent change column to 2 decimal places
        Range("K1:K" & Cells(1,11).End(xlDown).Row).NumberFormat = "0.00%"

        'bold table headers
        'outline table cell boxes

        'autofit columns of worksheet table
        Columns("J:L").Autofit
        'change colour of yearly change cells based on +ive or -ive result 
        lastRow = Cells(1,10).End(xlDown).Row
        for i = 2 to lastRow
            'if statement to loop throuch cells and check value of cells to determine color
            if Cells(i,10).Value > 0 Then
                Cells(i,10).Interior.ColorIndex = 4
            Elseif Cells(i,10).Value < 0 Then
                Cells(i,10).Interior.ColorIndex = 3
            else
                Cells(i,10).Interior.ColorIndex = 15
            End if
        Next
        
    'end primary analysis
    'End Sub
    
    'Sub Secondary_Analysis():

        'declare variables
        Dim maxPercentIncrease, MaxPercentDecrease As double
        Dim maxStockVolume as long

        'define variables
        maxPercentIncrease = 0
        MaxPercentDecrease = 0
        maxStockVolume = 0
        
        'creating table headers
        Cells(1,16).Value = "Ticker"
        Cells(1,17).Value = "Value"
        Cells(2,15).Value = "Greatest % Increase"
        Cells(3,15).Value = "Greatest % Decrease"
        Cells(4,15).Value = "Greatest Total Volume"

        'calculate the last row of primary analysis table
        lastRow = Cells(9,1).End(xlDown).Row

        'for statement to loop through primary analysis table
        for i = 2 to lastRow
            
            'if statement to calculate the greatest % increase
            if Cells(i,11).Value > maxPercentIncrease Then
                
                maxPercentIncrease = Cells(i,11).Value
                'print ticker values to table
                Cells(2,16).Value = Cells(i,9).Value
                'print greatest % increase value to table
                Cells(2,17).Value = maxPercentIncrease

            'ending if statement for greatest % increase
            End if

            'if statement to calculate the greatest % decrease
            if Cells(i,11).Value < maxPercentDecrease Then
                
                maxPercentDecrease = Cells(i,11).Value
                'print ticker values to table
                Cells(3,16).Value = Cells(i,9).Value
                'print greatest % decrease value to table
                Cells(3,17).Value = maxPercentDecrease

            'ending if statement for greatest % decrease
            End if

            'if statement to calculate the greatest volume traded
            if Cells(i,12).Value > maxStockVolume Then
                
                maxStockVolume = CLng(Cells(i,12).Value)
                'print ticker values to table
                Cells(4,16).Value = Cells(i,9).Value
                'print greatest volume traded to table
                Cells(4,17).Value = maxStockVolume

            'ending if statement for greatest volume traded
            End if

        'end for statement
        Next i

        'formatting secondary analysis table
        Range("Q2:Q3").NumberFormat = "0.00%"
        'auyto fit colums of secondary analtsys table
        Columns("O").Autofit
        Columns("Q").Autofit
    
    'end secondary analysis
    'End Sub

End Sub
