Attribute VB_Name = "Module1"
Sub ticker_summarizer()
    'This script is to detect the tickers in a worksheet and print
    'it to the summary table.
    
    
    'Declare variables
    Dim ticker As String
    Dim summary_row As Integer
    Dim end_row As Long
    Dim total_volume As LongLong
    Dim year_open As Single
    Dim year_close As Single
    Dim percent_change As Single
    Dim year_change As Single
    Dim summary_end_row As Integer
    Dim last_sum_cell As String
            
    'Make the columns fit the width of the data contained
    Columns("A:G").AutoFit
            
    'Detect the last row of the spreadsheet, taken from example 5 of lecture 2.3
    end_row = Cells(Rows.Count, 1).End(xlUp).Row
    'MsgBox ("The last row is row " & end_row)
    
    'Create headers for summary table
    Cells(1, 9).Value = "Summary Table"
    Cells(2, 9).Value = "Ticker"
    Cells(2, 10).Value = "Yearly Change"
    Cells(2, 11).Value = "Percent Change"
    Cells(2, 12).Value = "Total Stock Volume"
    
    'Start summary table writing at row 3, print initial values
    summary_row = 3
    year_open = Cells(2, 3).Value
    ticker = Cells(2, 1).Value
    Cells(summary_row, 9).Value = ticker
    
    'Start total stock volume at 0
    total_volume = 0
        
    'Read ticker value, write unique tickers to table in column I
    For r = 2 To end_row
        
        'Add up total volume over all rows of the ticker
        total_volume = total_volume + Cells(r, 7).Value
        
        'Compare row r to row r+1 to see if they are the same (assuming tickers are in alphabetical order)
        If ticker <> Cells(r + 1, 1) Then
            'MsgBox ("Ticker changed from " & ticker & " to " & Cells(r + 1, 1).Value & " at row " & r)
           
            'Print total stock volume to table
            Cells(summary_row, 12).Value = total_volume
           
            'Reset the total stock volume
            total_volume = 0
                   
            'Get next ticker
            ticker = Cells(r + 1, 1).Value
            'MsgBox ("New ticker is " & ticker)
            'MsgBox (ticker & "starts at row " & year_start_row)
            
            'Printing close value for a given ticker (assuming all are in alphabetical and chonological order)
            year_close = Cells(r, 6).Value
                
            'Calculate and print yearly change
            year_change = year_close - year_open
            Cells(summary_row, 10) = year_change
           
            'Calculate percent change (I'm leaving the conversion to a percent to the loop where I format the summary table
            percent_change = ((year_close - year_open) / year_open)
            Cells(summary_row, 11).Value = percent_change
           
            summary_row = summary_row + 1
            Cells(summary_row, 9).Value = ticker
        
            'Pull in year_open value for next ticker
            year_open = Cells(r + 1, 3).Value
            num_of_tickers = num_of_tickers + 1
            
        End If
        
    Next r
        
    'Format summary table
    
    
    summary_end_row = num_of_tickers + 2
    
    For cell = 3 To summary_end_row
       
        Cells(cell, 10).NumberFormat = "0.00" 'Found instructions on StackOverflow
        Cells(cell, 11).NumberFormat = "0.00%"
        
    Next cell
    
    'Apply green/red conditional formatting (I added that cells with 0% change should be yellow, since they neither increased nor decreased)
    For s_row = 3 To summary_end_row
    
        If Cells(s_row, 10).Value > 0 Then
            Cells(s_row, 10).Interior.ColorIndex = 4
        ElseIf Cells(s_row, 10).Value < 0 Then
            Cells(s_row, 10).Interior.ColorIndex = 3
        Else
            Cells(s_row, 10).Interior.ColorIndex = 6
            
        End If
    Next s_row
        
    'Make summary table columns fit data within
    Columns("I:L").AutoFit
        
    'Create bonus table
    Cells(1, 15).Value = "Bonus Table"
    Cells(3, 15).Value = "Greatest % Increase"
    Cells(4, 15).Value = "Greatest % Decrease"
    Cells(5, 15).Value = "Greatest Total Volume"
    Cells(2, 16).Value = "Ticker"
    Cells(2, 17).Value = "Value"
    
    '-----------------------------------------------
    
    'Find maximum percent increase, print to bonus table
    Dim i As Single
    
    i = 0
    ticker = ""
    
    For r = 3 To summary_end_row
        If Cells(r, 11).Value > i Then 'borrowed from educba.com
            i = Cells(r, 11).Value
            ticker = Cells(r, 9).Value
        End If
    Next r
    'MsgBox ("Max value is " & i)
    'MsgBox ("Max value is " & i & " for ticker " & ticker)
    Cells(3, 16).Value = ticker
    Cells(3, 17).Value = i
    Cells(3, 17).NumberFormat = "0.00%"
    
    'Find minimum percent increase
    
    i = 0
    ticker = ""
    
    For r = 3 To summary_end_row
        If Cells(r, 11).Value < i Then
            i = Cells(r, 11).Value
            ticker = Cells(r, 9).Value
        End If
    Next r
    'MsgBox ("Min value is " & i & " for ticker " & ticker)
    Cells(4, 16).Value = ticker
    Cells(4, 17).Value = i
    Cells(4, 17).NumberFormat = "0.00%"
    
    'Find maximum stock volume
    
    i = 0
    ticker = ""
    
    For r = 3 To summary_end_row
        If Cells(r, 12).Value > i Then
            i = Cells(r, 12).Value
            ticker = Cells(r, 9).Value
        End If
    Next r
    'MsgBox ("Max total volume is " & i & " for ticker " & ticker)
    Cells(5, 16).Value = ticker
    Cells(5, 17).Value = i
    
    'Make bonus table columns fit data within
    Columns("O:Q").AutoFit
    
End Sub
