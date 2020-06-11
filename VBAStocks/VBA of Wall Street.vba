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

'declare variables
Dim i, posTracker As integer
Dim ticker As string
Dim priceOpen, priceClose, yearlyChange, sotckVolume As double

'define variables
posTracker = 2
sotckVolume = 0
priceOpen = 0
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
        priceOpen = Cells(i,3).Value
        priceClose = Cells(i,6).Value
        yearlyChange = priceClose - priceOpen

            'fill in table based on ticker symbol
            'display ticker code on column I
            Cells(posTracker,9).Value = ticker
            'display yearly change on column J
            Cells(posTracker,10).Value = yearlyChange
            Cells(posTracker,13).Value = priceOpen
            Cells(posTracker,14).Value = priceClose
            'display percent change on column K

            'display stock volume on column L
            Cells(posTracker,12).Value = sotckVolume + Cells(i,7).Value
        
        'moving down 1 row on table when ticker changes
        posTracker = posTracker + 1

        'resetting table trackers
        yearlyChange = 0
        sotckVolume = 0
            
    'else statement to keep track of totals and changes per ticker
    Else

        'running total for stock volume
        sotckVolume = sotckVolume + Cells(i,7).Value

    'end if statement to check when ticker symbol changes
    End if

'end for statement to loop through rows of data
Next i
        
'autofit columns of worksheet table
Columns("J:L").Autofit        

End Sub
