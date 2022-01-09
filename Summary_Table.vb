Attribute VB_Name = "Summary_Table"
Sub StockMarketAnalysis()
    
    'In order for this to run through multiple worksheets 'ws' had to be added for this to work
    For Each ws In Worksheets
    
        ws.Cells(1, "I").Value = "Ticker"
        ws.Cells(1, "J").Value = "Yearly Change"
        ws.Cells(1, "K").Value = "Percent Change"
        ws.Cells(1, "L").Value = "Total Stock Volume"
        ws.Cells(2, "O").Value = "Greatest % Increase"
        ws.Cells(3, "O").Value = "Greatest % Decrease"
        ws.Cells(4, "O").Value = "Greatest Total Volume"
        ws.Cells(1, "P").Value = "Ticker"
        ws.Cells(1, "Q").Value = "Value"
    
        'Originally wanted to use 'open' and 'close' as variables but for whatever reason vba didnt like it, so here are my variables
        'These are the variables I defined to use in my script
        Dim tick As String
        Dim lastrow As Long
        Dim openprice As Double
        Dim closeprice As Double
        Dim yearchange As Double
        Dim percentchange As Double
        Dim totalstockvol As Double
        totalstockvol = 0
        Dim totalrow As Long
        totalrow = 2
        Dim prevamt As Long
        prevamt = 2
        'Bonus ----------
        Dim GreatestIncrease As Double
        GreatestIncrease = 0
        Dim GreatestDecrease As Double
        GreatestDecrease = 0
        Dim GreatestTotalVolume As Double
        GreatestTotalVolume = 0
        
        'This is essentially how vba goes to the last row and grabs the data from the last cell (googled for this one)
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'This is a for loop starting at index 2 and it loops through to the last row
        For i = 2 To lastrow
            
            totalstockvol = totalstockvol + ws.Cells(i, 7).Value
            
            'This is how we check if the ticker name has changed, whenever the ticker cell above doesnt match ticker cell below
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                'This code will organize all of our ticker names
                tick = ws.Cells(i, 1).Value
                
                'This grabs the ticker name and its respective summed amount for our table
                ws.Range("I" & totalrow).Value = tick
                ws.Range("L" & totalrow).Value = totalstockvol
                
                'This line is important to be able to move from ticker to ticker cleanly, it resets the value from the previous ticker
                totalstockvol = 0
                
                'Pretty self explanitory with the names but this gets us our opening and closing price in order to get our yearly change
                openprice = ws.Range("C" & prevamt)
                closeprice = ws.Range("F" & i)
                yearchange = closeprice - openprice
                
                'This will be how we grab the respective yearly change per ticker for our new table
                ws.Range("J" & totalrow).Value = yearchange
                
                'Since the alphabet testing document did not have zeros this was added specific to this data set because if not we get errors when trying to get some percentages (cant #DIV/0)
                If openprice = 0 Then

                    percentchange = 0
                Else
                    openprice = ws.Range("C" & prevamt)
                    percentchange = yearchange / openprice
                End If
                
                'This is the grabbing the respective ticker percent change for our table and formatting it to be a percent
                ws.Range("K" & totalrow).NumberFormat = "0.00%"
                ws.Range("K" & totalrow).Value = percentchange

                'For the color index I found the regular green and red used in the example too blinding so I went with some lighter colors
                'It looks and sees if the specific cell in column J is greater than or = to 0 to determine if it will be green or red
                If ws.Range("J" & totalrow).Value >= 0 Then
                    ws.Range("J" & totalrow).Interior.ColorIndex = 35
                Else
                    ws.Range("J" & totalrow).Interior.ColorIndex = 38
                End If
                
                'This adds one to get this ready to loop through again
                totalrow = totalrow + 1
                prevamt = i + 1
                End If
            
            Next i

            lastrow = ws.Cells(Rows.Count, 11).End(xlUp).Row
            
            'Our last for loop to give us our values for the bonus as they look through column K and L to find the the largest numbers there to grab with their respective ticker name
            For i = 2 To lastrow
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
            'This makes the values in Q2 and Q3 have the percent format
            Next i
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"

    Next ws

End Sub
