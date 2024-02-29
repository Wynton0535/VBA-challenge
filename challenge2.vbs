Sub stocks():
            'DECLARATION OF VARIABLES
    Dim i As Long
    Dim lastrow As Long
    Dim tickervolume As Double
    Dim printcount As Integer
    Dim totalvolume As Double
    Dim tickeropen As Double
    Dim tickerclose As Double
    Dim percentchange As Double
    Dim ticker As String
    Dim ws As Worksheet
    Dim newincrease As Double
    Dim newdecrease As Double
    Dim newtotal As Double
    
            'FOR EACH WORKSHEET
    For Each ws In Worksheets
    
    

            'INITIALIZE VARIABLES
        tickervolume = 0
        printcount = 1
        totalvolume = 0
        tickeropen = 0
        tickerclose = 0
        percentchange = 0
        newincrease = 0
        newdecrease = 0
        newtotal = 0
        
                'ASSIGNING THE LAST ROW OF THE WORKSHEET TO VARIABLE LASTROW
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        
                'PRINT FORMAT FOR TABLE
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        
        For i = 2 To lastrow
            ticker = ws.Cells(i, 1).Value
            

                    'IF THE VALUE OF THE TICKER IS DIFFERENT FROM THE PREVIOUS ONE THEN...
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                tickeropen = ws.Cells(i, 3).Value
                totalvolume = totalvolume + ws.Cells(i, 7).Value

                
                    'IF THE VALUE OF THE TICKER IS DIFFERENT FROM THE NEXT ONE THEN...
            ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                printcount = printcount + 1
                tickerclose = tickerclose + ws.Cells(i, 6).Value
                
                yearchange = tickerclose - tickeropen
                totalvolume = totalvolume + ws.Cells(i, 7).Value
                
                percentchange = (tickerclose - tickeropen) / tickeropen
                
                        'CONDITIONAL FORMATTING TO PERCENT CHANGE COLUMN
                ws.Range("K:K").NumberFormat = "0.00%"
                
                
                        'CONDITIONAL FORMATTING TO YEARLY CHANGE COLUMN
                If yearchange > 0 Then
                    ws.Range("J" & printcount).Interior.ColorIndex = 4
                
                ElseIf yearchange < 0 Then
                    ws.Range("J" & printcount).Interior.ColorIndex = 3
                End If
                             
                        'PRINTS THE VALUES OF TICKER, YEARLY CHANGE, PERCENT CHANGE, AND TOTAL STOCK VOLUME
                ws.Cells(printcount, 9).Value = ticker
                ws.Cells(printcount, 10).Value = yearchange
                ws.Cells(printcount, 11).Value = percentchange
                ws.Cells(printcount, 12).Value = totalvolume
                
                        'SEARCH FOR THE HIGHEST AND LOWEST NUMBER IN PERCENT CHANGE COLUMN
                If percentchange > newincrease Then
                    newincrease = percentchange
                    ws.Range("P2").Value = ticker
                ElseIf percentchange < newdecrease Then
                    newdecrease = percentchange
                    ws.Range("P3").Value = ticker
                End If
                
                        'SEARCH FOR THE HIGHEST TOTAL STOCK VOLUME
                If totalvolume > newtotal Then
                    newtotal = totalvolume
                    ws.Range("P4").Value = ticker
                End If
                
                ws.Range("Q2").NumberFormat = "0.00%"
                ws.Range("Q3").NumberFormat = "0.00%"
                
                
                tickerclose = 0
                totalvolume = 0
                
            Else
                totalvolume = totalvolume + Cells(i, 7).Value
            End If
    
        Next i

                'PRINTS THE GREATEST INCREASE, GREATEST DECRESE, AND GREATEST TOTAL VOLUME
        ws.Range("Q2").Value = newincrease
        ws.Range("Q3").Value = newdecrease
        ws.Range("Q4").Value = newtotal

    Next ws

End Sub


