Sub stockdata():

    'run code for each worksheet
    For Each ws In Worksheets
    
        'declaring variables
        Dim lastrow As Long
        Dim yearlychange As Double
        Dim percentchange As Double
        Dim totalvolume As Double
        Dim summaryrow As Integer
        Dim openprice As Double

        'need to manually set the first open price data to calculate percent change, the rest is set in the loop
        openprice = ws.Cells(2, 3).Value
    

        totalvolume = 0
        'need our summary table to start at second row, first row for headers
        summaryrow = 2
    
        'summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    
        'find last row
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To lastrow

            'add up the total volume for each ticker
            totalvolume = totalvolume + ws.Cells(i, 7)
        
            'go down the row to until different ticker name
            'refer back to creditcard activity
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'adding ticker name to summary table
                ws.Cells(summaryrow, 9).Value = ws.Cells(i, 1).Value
            
                'figure out the yearly change
                closeprice = ws.Cells(i, 6).Value
                yearlychange = closeprice - openprice
                ws.Cells(summaryrow, 10).Value = yearlychange
            
                'percent change
                'we need this if statement to get rid of the divide by zero error
                If openprice = 0 Then
                    percentchange = 0
                Else
                    percentchange = yearlychange / openprice
                End If
                
                'add percent change value to the summary table
                'use Format() to format the cell values as a percentage
                ws.Cells(summaryrow, 11).Value = Format(percentchange, "Percent")
            
                'add up total volume
                ws.Cells(summaryrow, 12).Value = totalvolume
            
                'reset total volume for next ticker
                totalvolume = 0
            
                'go to the next row in the summary table for the next ticker
                summaryrow = summaryrow + 1
            
                'get the open price of the next ticker
                openprice = ws.Cells(i + 1, 3).Value
            
            End If
        Next i
        
        'need to count the rows in the summary table instead of the original data
        lastrowpercent = ws.Cells(Rows.Count, 10).End(xlUp).Row


        'highlight percent change
        'color index: Red = 3, Green = 4
        
        For i = 2 To lastrowpercent
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
            
        Next i
        
    'summary table headers'
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"


        'loop to find the max, min, etc. values to construct summary table
        'remember to use lastrowpercent instead of lastrow because we need to go to the last row of the proccessed data
        For i = 2 To lastrowpercent
        
        'declaring variables for each value
        maxincrease = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrowpercent))
        maxdecrease = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrowpercent))
        maxvolume = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrowpercent))
            
            'max percent increase
            'use the if statement so we can also find the corresponding ticker for max increase
            If ws.Cells(i, 11).Value = maxincrease Then
                ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
                'need to format this as a percent as well
                ws.Cells(2, 16).Value = Format(maxincrease, "Percent")
                
            'min percent increase
            ElseIf ws.Cells(i, 11).Value = maxdecrease Then
                ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 16).Value = Format(maxdecrease, "Percent")
            
            'max total volume
            'also format this to curreny as well
            ElseIf ws.Cells(i, 12).Value = maxvolume Then
                ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 16).Value = maxvolume
            
            End If
        Next i
        
    Next ws
    
End Sub