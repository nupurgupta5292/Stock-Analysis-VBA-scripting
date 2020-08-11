Sub VBAChallenge()

    'Declaring all the variable used throughout the VBA Script
    Dim Ticker_Symbol, MIN_TICKER_NAME, MAX_TICKER_NAME, MAX_VOLUME_TICKER As String
    Dim Yearly_Change, Percent_Change As Double
    Dim Tot_Stock_Volume, MAX_VOLUME As Double
    Dim Summary_Table_Row, Summary_Table_column As Integer
    Dim first_opening_price, last_closing_price As Double
    Dim MAX_PERCENT As Double
    Dim ws As Worksheet
    Set ws = Worksheets("2016")
    
    'For loop to run the VBA script across worksheets
    For Each ws In Worksheets
    
        'Initializing variables that need to be reset for every new worksheet
        Summary_Table_Row = 2
        Summary_Table_column = 9
        Tot_Stock_Volume = 0
        first_opening_price = ws.Cells(2, 3).Value
        
        'Printing headers for the summary table
        ws.Cells(1, Summary_Table_column).Value = "Ticker"
        ws.Cells(1, Summary_Table_column + 1).Value = "Yearly Change"
        ws.Cells(1, Summary_Table_column + 2).Value = "Percent Change"
        ws.Cells(1, Summary_Table_column + 3).Value = "Total Stock Volume"
     
        'Determining the last row of the worksheet
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'For loop to compare values for each row
        For i = 2 To last_row
        
            'If condition on the Ticker Symbol to identify the point where information for one Ticker Symbol ends and another starts
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'Assigning value of Ticker Symbol to the variable
                Ticker_Symbol = ws.Cells(i, 1).Value
                
                'Calculating yearly change from opening price at the beginning of a given year to the closing price at the end of that year
                last_closing_price = ws.Cells(i, 6).Value
                Change = last_closing_price - first_opening_price
                               
                'Calculating percentage of yearly change
                If first_opening_price <> 0 Then
                    Percent_Change = Change / first_opening_price
                Else
                    ' Unlikely, but it needs to be checked to avoid program crushing
                    MsgBox ("For " & Ticker_Symbol & ", Row " & CStr(i) & ": Open Price = " & first_opening_price & ". Fix <open> field manually and save the spreadsheet.")
                End If
                
                
                'Calculating the total stock volume for that ticker symbol
                Tot_Stock_Volume = Tot_Stock_Volume + ws.Cells(i, 7).Value
                
                'Putting values and formatting in summary table for different metrics
                'Printing the Ticker Symbol in point
                ws.Cells(Summary_Table_Row, Summary_Table_column).Value = Ticker_Symbol
                
                'Printing the yearly change
                ws.Cells(Summary_Table_Row, Summary_Table_column + 1).Value = Change
                
                'Conditional formatting for highlighting negative values in red and positive value of yearly change in Green
                If ws.Cells(Summary_Table_Row, Summary_Table_column + 1).Value < 0 Then
                    ws.Cells(Summary_Table_Row, Summary_Table_column + 1).Interior.ColorIndex = 3
                    
                Else
                    'Calculating sum of the stock volume for a particular ticker symbol till a ticker symbol mismatch is incurred
                    ws.Cells(Summary_Table_Row, Summary_Table_column + 1).Interior.ColorIndex = 4
                    
                End If
                
                'Printing the percentage of yearly change
                ws.Cells(Summary_Table_Row, Summary_Table_column + 2).Value = Percent_Change
                
                'Converting the format of percent change to percentage
                ws.Cells(Summary_Table_Row, Summary_Table_column + 2).NumberFormat = "0.00%"
                
                'Printing the total stock volume for the particular Ticker Symbol
                ws.Cells(Summary_Table_Row, Summary_Table_column + 3).Value = Tot_Stock_Volume
                
		' To check greaest increase and decrease % and maximum volume
                If (Percent_Change > MAX_PERCENT) Then
                    MAX_PERCENT = Percent_Change
                    MAX_TICKER_NAME = Ticker_Symbol
                ElseIf (Percent_Change < MIN_PERCENT) Then
                    MIN_PERCENT = Percent_Change
                    MIN_TICKER_NAME = Ticker_Symbol
                
                End If
                       
                If (Tot_Stock_Volume > MAX_VOLUME) Then
                    MAX_VOLUME = Tot_Stock_Volume
                    MAX_VOLUME_TICKER = Ticker_Symbol
                End If
                
                
                'Resetting total stock volume to 0 before new iteration for the next Ticker Symbol starts
                Tot_Stock_Volume = 0
                
                'Reseting opening price for next Ticker symbol
                first_opening_price = ws.Cells(i + 1, 3).Value
                last_closing_price = 0
  
                'Incrementing summary table row by 1 for the next ticker symbol to print in next row
                Summary_Table_Row = Summary_Table_Row + 1
            
            Else
                'Adding to the sum of total stock volume if the row is for same ticker symbol
                Tot_Stock_Volume = Tot_Stock_Volume + ws.Cells(i, 7).Value
            
            End If
        
        Next i
        
        
        'Printing the headers for the the summary table with values or greatest
        ws.Cells(1, Summary_Table_column + 7).Value = "Ticker"
        ws.Cells(1, Summary_Table_column + 8).Value = "Value"
        ws.Cells(Summary_Table_Row, Summary_Table_column + 6).Value = "Greatest % Increase"
        ws.Cells(Summary_Table_Row + 1, Summary_Table_column + 6).Value = "Greatest % Decrease"
        ws.Cells(Summary_Table_Row + 2, Summary_Table_column + 6).Value = "Greatest Total Value"
        
            
        'Printing the values of greatest percent increase, decrease and total stock value
        ws.Cells(Summary_Table_Row, Summary_Table_column + 7).Value = MAX_TICKER_NAME
        ws.Cells(Summary_Table_Row, Summary_Table_column + 8).Value = MAX_PERCENT
        ws.Cells(Summary_Table_Row, Summary_Table_column + 8).NumberFormat = "0.00%"
        
        ws.Cells(Summary_Table_Row + 1, Summary_Table_column + 7).Value = MIN_TICKER_NAME
        ws.Cells(Summary_Table_Row + 1, Summary_Table_column + 8).Value = MIN_PERCENT
        ws.Cells(Summary_Table_Row + 1, Summary_Table_column + 8).NumberFormat = "0.00%"
        
        ws.Cells(Summary_Table_Row + 2, Summary_Table_column + 7).Value = MAX_VOLUME_TICKER
        ws.Cells(Summary_Table_Row + 2, Summary_Table_column + 8).Value = MAX_VOLUME
           
             
    Next ws

End Sub
