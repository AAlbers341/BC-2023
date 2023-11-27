Attribute VB_Name = "Module1"
Sub multipleyearstockdata()


'Looping logic through each worksheet
For Each ws In Worksheets
    
    'Finding last row of each sheet in Workbook
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Creating sorthand worksheet logic
    WorksheetName = ws.Name
    
    'Create 'Ticker', 'Yearly Change', 'Percent Change' & 'Total Stock Volume' columns on each ws
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'Greatest '% Increase', '% Decrease', 'Total Stock Volume', 'Ticker' & 'Volume' Headers
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
'Declare ticker name
Dim ticker_name As String

'Declare ticker location in analysis table
Dim ticker_analysis_table As Integer
ticker_analysis_table = 2

'Delcare opening price, closing price, percent change, and ticker volume for each ticker
Dim opening_price As Double
Dim closing_price As Double
Dim percent_change As Double
Dim ticker_volume As Variant
Dim greatest_percent_ticker As String
Dim lowest_percent_ticker As String
Dim greatest_volume As String
greatest_volume = 0
lowest_percent_ticker = 0
percent_increase_ticker = 0
ticker_volume = 0
percent_change = 0
opening_price = 0
closing_price = 0


    'Loop through to identify individual attributes for each ticker
    For i = 2 To LastRow
    
        'Formating the 'Date' column using IsNumeric to validate if column is a 'numeric' value. Then using 'left', 'mid' and 'right' index functions to split the data and Format the Datevalue by concatentation "/"
        If IsNumeric(ws.Cells(i, 2).Value) Then
        ws.Cells(i, 2).Value = Format(DateValue(Left(ws.Cells(i, 2).Value, 4) & "/" & Mid(ws.Cells(i, 2).Value, 5, 2) & "/" & Right(ws.Cells(i, 2).Value, 2)), "mm/dd/yyyy")
        End If
        
        'Conditional to see obtain opening price for each ticker
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        opening_price = opening_price + ws.Cells(i, 3).Value
        End If
        
        'Volume to add all the stocks associated with that ticker
        If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
        ticker_volume = ticker_volume + ws.Cells(i, 7).Value
        ws.Range("L" & ticker_analysis_table).Value = ticker_volume
        End If
        
        
        'Conditional to see if we are still within the name ticker name
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        'Set ticker name
        ticker_name = ws.Cells(i, 1).Value
        
        'Set closing price
        closing_price = ws.Cells(i, 6).Value
    
        
        ws.Range("I" & ticker_analysis_table).Value = ticker_name
        ws.Range("J" & ticker_analysis_table).Value = closing_price - opening_price
        
        'Calculate % Change and Round to the nearest 100th
        percent_change = (opening_price - closing_price) / opening_price
        ws.Range("K" & ticker_analysis_table).Value = percent_change
        
         'Format color index for negative & positive opening - closing prices
        If closing_price - opening_price < 0 Then
        ws.Range("J" & ticker_analysis_table).Interior.ColorIndex = 3
        ws.Range("J" & ticker_analysis_table).Font.ColorIndex = 0
        Else: ws.Range("J" & ticker_analysis_table).Interior.ColorIndex = 4
        ws.Range("J" & ticker_analysis_table).Font.ColorIndex = 0
        End If
        
        'Format % change to %
        ws.Range("K" & ticker_analysis_table).NumberFormat = "0.00%"
        
        
        'Next row on ticker analysis table
        ticker_analysis_table = ticker_analysis_table + 1
        
        
        
        'Clear closing and opening prices for next ticker
        closing_price = 0
        opening_price = 0
        ticker_volume = 0
        
        End If
    Next i
    
    'Greatest % increase
    ws.Cells(2, 17).Value = Application.WorksheetFunction.Max(ws.Range("k:k"))
    ws.Cells(2, 17).NumberFormat = "0.00%"
    

    'Greatest % decrease
    ws.Cells(3, 17).Value = Application.WorksheetFunction.Min(ws.Range("k:k"))
    ws.Cells(3, 17).NumberFormat = "0.00%"

    'Greatest Volume Total
    ws.Cells(4, 17).Value = Application.WorksheetFunction.Max(ws.Range("l:l"))
    
    'Assign ticker to 'Greatest % increase', 'Greatest % decrease" & 'Greatest Volume'
    For j = 2 To LastRow
        If ws.Cells(2, 17).Value = ws.Cells(j, 11).Value Then
         greatest_percent_ticker = ws.Cells(j, 11 - 2).Value
         ws.Range("P" & 2).Value = greatest_percent_ticker
        End If
        
        If ws.Cells(3, 17).Value = ws.Cells(j, 11).Value Then
        lowest_percent_ticker = ws.Cells(j, 11 - 2).Value
        ws.Range("P" & 3).Value = lowest_percent_ticker
        End If
        
        If ws.Cells(4, 17).Value = ws.Cells(j, 12).Value Then
        greatest_volume = ws.Cells(j, 12 - 3).Value
        ws.Range("P" & 4).Value = greatest_volume
        End If
    Next j
    
        
    
Next ws


End Sub
