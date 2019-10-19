Sub VBAStocks()

For Each ws In Worksheets
    
    'create summary table column headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Totoal Stock Volume"
    
    'create greatest summary table
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'set initial variables
    Dim Ticker_Name As String
    Dim Yearly_Chg As Double
    Dim Percent_Chg As Double
    Dim Tot_Stock_Vol As LongLong
    Tot_Stock_Vol = 0
    
    'find the last row for each worksheet and create a variable for it
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'keep track of row location in summary table
    Dim Table_Row As Integer
    Table_Row = 2
    
        'loop through all tickers
        For i = 2 To lastrow
    
            'if ticker is not the same as the one below it then...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                Ticker_Name = ws.Cells(i, 1).Value
            
                Tot_Stock_Vol = Tot_Stock_Vol + ws.Cells(i, 7).Value
            
                Year_Close = ws.Cells(i, 6).Value
                Data_Row = Data_Row + 1
                year_open = ws.Cells(((i + 1) - Data_Row), 3).Value
            
                Yearly_Chg = Round((Year_Close - year_open), 9)
            
                'division by zero check
                If year_open = 0 Then
                    Percent_Chg = 0
                Else
                    Percent_Chg = Yearly_Chg / year_open
                End If
                
                'fill summary table with values
                ws.Range("I" & Table_Row).Value = Ticker_Name
                ws.Range("J" & Table_Row).Value = Yearly_Chg
                ws.Range("K" & Table_Row).Value = Percent_Chg
                ws.Range("L" & Table_Row).Value = Tot_Stock_Vol
            
                'color conditional format for yearly change
                If Yearly_Chg < 0 Then
                    ws.Range("J" & Table_Row).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & Table_Row).Interior.ColorIndex = 4
                End If
            
                'move to next row in summary table and set stock vol and data row to zero
                Table_Row = Table_Row + 1
                Tot_Stock_Vol = 0
                Data_Row = 0
        
            Else
                'continue to count tickers row spot and sum its stock volume
                Tot_Stock_Vol = Tot_Stock_Vol + ws.Cells(i, 7).Value
                Data_Row = Data_Row + 1
            End If
    
        Next i

    'format the summary table columns
    ws.Columns("I:L").AutoFit
    tablelastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    
    'number format for percent change
    Dim Percent_Chg_Format As Range
    Set Percent_Chg_Format = ws.Range("K2:K" & tablelastrow)
    Percent_Chg_Format.NumberFormat = "#0.00%"

        'greatest values summary table challenge
        Dim max_vol As LongLong
        max_vol = 0
        max_percent_chg = 0
        min_percent_chg = 0
        
        For i = 2 To tablelastrow
            
            If ws.Cells(i, 12).Value > max_vol Then
            max_vol = ws.Cells(i, 12).Value
            max_vol_ticker = ws.Cells(i, 9).Value
            End If
            
            If ws.Cells(i, 11).Value > max_percent_chg Then
            max_percent_chg = ws.Cells(i, 11).Value
            max_percent_chg_ticker = ws.Cells(i, 9).Value
            End If
            
            If ws.Cells(i, 11).Value < min_percent_chg Then
            min_percent_chg = ws.Cells(i, 11).Value
            min_percent_chg_ticker = ws.Cells(i, 9).Value
            End If
            
        Next i
        
    'put greatest values into summary table and format
    ws.Range("p2").Value = max_percent_chg_ticker
    ws.Range("q2").Value = max_percent_chg
    ws.Range("p3").Value = min_percent_chg_ticker
    ws.Range("q3").Value = min_percent_chg
    ws.Range("p4").Value = max_vol_ticker
    ws.Range("q4").Value = max_vol
    ws.Range("q2:q3").NumberFormat = "#0.00%"
    ws.Columns("O:Q").AutoFit
    
Next ws

End Sub
