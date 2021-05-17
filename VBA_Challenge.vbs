Sub StockMarketData()

    'Initializing variables that will hold the intial price of the year and closing price of the year
    'Initializing variables that will hold the yearly price change of a stock ticker, the total
    '% change of a stock % over the course of the year, and the total volume of trade for that stock
    'ticker over the course of the year
    
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TotalStockVolume As LongLong
    
          
    'Create Variable counter to help build analysis table
    
        Dim StockCounter As Integer
    
    'Bonus Variables
    'Need a lastrow counter for the analysis table to create a range
    'Need a variable to hold the ticker of the greatest % increase
    'Need a variable to hold the value of the greatest % increase
    'Need a variable to hold the ticker of the greatest % decrease
    'Need a variable to hold the value of the greatest % decrease
    'Need a variable to hold the ticker of the greatest total volume
    'Need a variable to hold the value of the greatest total volume
       
        Dim tablelastrow As Integer
        Dim TickerPercentIncrease As String
        Dim ValuePercentIncrease As Double
        Dim TickerPercentDecrease As String
        Dim ValuePercentDecrease As Double
        Dim TickerVolume As String
        Dim ValueVolume As LongLong
        
    ' --------------------------------------------
    ' Creating a For loop to perform analysis from sheet to sheet
           
    For Each ws In Worksheets
             
        'Setting analysis table counting variable to 2 so that values will begin to populate on row 2
        
        StockCounter = 2
        
        'Creating table analysis headers in worksheet
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Perchent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'Identifying the last row of the current worksheet
                
        lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        
        'Assign the 0 value to the TotalStockVolume variable
        'Assign the first row of data's open price to OpenPrice variable for the first stock of the sheet
        
        TotalStockVolume = 0
        OpenPrice = ws.Cells(2, 3).Value
        
        'Creating a For loop to iterate from i = 2 to lastrow times
                
        For i = 2 To lastrow
        
            'Conditional to check if the ticker symbol is the same in the current row and the next row
            
            If ws.Cells(i, 1) = ws.Cells(i + 1, 1) Then
            
            'if the ticker symbol stays the same then we need to update the volume calcualation
                
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                        
            'If the ticker symbol changes in the next row then we need to do a number of actions
            
            Else
                
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                'Number of Actions to fill in the analysis table
                ws.Cells(StockCounter, 9).Value = ws.Cells(i, 1).Value
                'Align the closing price of this row's data to the ClosePrice variable
                ClosePrice = ws.Cells(i, 6).Value
                'Calculate the yearly change of the stock price by subtracting the opening price from the closing price
                YearlyChange = ClosePrice - OpenPrice
                'Insert YearlyChange variable into the analysis table
                ws.Cells(StockCounter, 10).Value = YearlyChange
                
                'Conditional statement to assign formating to the cell with YearlyChange
                    If YearlyChange >= 0 Then
                
                    'Assigning the color green for positive or 0 change over the course of a year
                        ws.Cells(StockCounter, 10).Interior.ColorIndex = 4
                    'Assigning the color red for negative change over the course of a year
                    Else
                        
                        ws.Cells(StockCounter, 10).Interior.ColorIndex = 3
                    
                    End If
                'Calculate Value for % change by dividing the yearly change by the opening price
                If YearlyChange = 0 Then
                    PercentChange = 0
                ElseIf OpenPrice = 0 Then
                    PercentChange = 0
                Else
                    PercentChange = YearlyChange / OpenPrice
                End If
                'Insert PercentChange into the analysis table
                ws.Cells(StockCounter, 11).Value = PercentChange
                'Change formatting of the PercentChange cell to %
                ws.Cells(StockCounter, 11).Style = "Percent"
                ws.Cells(StockCounter, 11).NumberFormat = "0.00%"
                'Insert TotalStockVolume into the analysis table
                ws.Cells(StockCounter, 12).Value = TotalStockVolume
                
                'Iterate and reset variables for the next ticker symbol
                TotalStockVolume = 0
                OpenPrice = ws.Cells(i + 1, 3)
                StockCounter = StockCounter + 1
                
            End If
        
        Next i
        
        'Bonus Work ---------------------------------------------
        'Identifying the last row of the analysis table in the current worksheet
        tablelastrow = ws.Cells(Rows.Count, "I").End(xlUp).Row
        'Assigning values using the Min and Max function to Bonus variables of interest
        ValuePercentIncrease = WorksheetFunction.Max(ws.Range("K2:K" & tablelastrow))
        ValuePercentDecrease = WorksheetFunction.Min(ws.Range("K2:K" & tablelastrow))
        ValueVolume = WorksheetFunction.Max(ws.Range("L2:L" & tablelastrow))
        'Need an iterative loop to find the corresponding stock tickers with the values that were just found
        For j = 2 To tablelastrow
        'Conditional checks to identify if this row contains any of the found Value variables
            If ws.Cells(j, 11).Value = ValuePercentIncrease Then
                TickerPercentIncrease = ws.Cells(j, 9).Value
            End If
        
            If ws.Cells(j, 11).Value = ValuePercentDecrease Then
                TickerPercentDecrease = ws.Cells(j, 9).Value
            End If
        
            If ws.Cells(j, 12).Value = ValueVolume Then
                TickerVolume = ws.Cells(j, 9).Value
            End If
        Next j
        
        'Create bonus analysis table
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(2, 16).Value = TickerPercentIncrease
        ws.Cells(3, 16).Value = TickerPercentDecrease
        ws.Cells(4, 16).Value = TickerVolume
        ws.Cells(2, 17).Value = ValuePercentIncrease
        ws.Cells(3, 17).Value = ValuePercentDecrease
        ws.Cells(4, 17).Value = ValueVolume
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Columns("A:Q").AutoFit
        
    Next ws

   
End Sub
