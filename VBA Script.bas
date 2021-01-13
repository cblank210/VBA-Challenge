Attribute VB_Name = "Module1"
Sub Stock_Analysis()

'-------VBA Script for Stock Analysis-------
    
   Dim ws As Worksheet
    
    'Create for each Loop that will go through each worksheet
    For Each ws In Worksheets
    
             'Add labels for the column header stock summary information
             ws.Cells(1, 10).Value = "Ticker"
             ws.Cells(1, 11).Value = "Yearly Change"
             ws.Cells(1, 12).Value = "Percent Change"
             ws.Cells(1, 13).Value = "Total Stock Volume"
             
             'Auto format column widths for data
             Columns("A:M").AutoFit
    
            'Add variable to hold ticker symbol
            Dim ticker_symbol As String
            
            'Add variable to hold yearly change in stock price
            Dim yearly_change As Double
            yearly_change = 0
            
            'Add variable to hold percent change of stock during the  year
            Dim percent_change As Double
            percent_change = 0
            
            'Add variable to hold total volume of stock
            Dim total_volume As Double
            total_volume = 0
            
            'Add variable to track opening price of each stock yearly
            Dim open_year As Double
            open_year = 0
            
            'Add variable to track closing price of each stock yearly
            Dim close_year As Double
            close_year = 0
            
            'Add variable to track location of each ticker in summary table
            Dim row_count As Long
            row_count = 2
            
            'Make last row variable to identify end of each data set for loop
            Dim last_row As Long
            last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
                'Create loop to search through all ticker symbols
                Dim i As Long
                For i = 2 To last_row
                
                   'Get total volume for each ticker symbol before entering if arguments
                    total_volume = total_volume + ws.Cells(i, 7).Value
                
                        'Find open and end year values and place in table
                        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                        open_year = ws.Cells(i, 3).Value
                            
                        End If
                    
                
                     'If then conditional to find each change in ticker value
                      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                        
                        'Place ticker symbol in table
                        ws.Cells(row_count, 10).Value = ws.Cells(i, 1).Value
            
                        ws.Cells(row_count, 13).Value = total_volume
                        
                        ws.Cells(row_count, 13).NumberFormat = "#,000"
                           
                        'Get year end stock price
                        end_year = ws.Cells(i, 6).Value
                        
                        'Calculate yearly change
                        yearly_change = end_year - open_year
                        
                        ws.Cells(row_count, 11).Value = yearly_change
                        
                        ws.Cells(row_count, 11).NumberFormat = "0.00"
                        
                                'New if to change formatting for positive and negative yearly changes
                                If yearly_change >= 0 Then
                                
                                    ws.Cells(row_count, 11).Interior.ColorIndex = 4
                                Else
                                
                                    ws.Cells(row_count, 11).Interior.ColorIndex = 3
                                End If
                        
                                'New if and calculate percentage change in each stock and place in table, and account for "0" or infinite calculations
                                If open_year = 0 And close_year = 0 Then
                                
                                    percent_change = 0
                                
                                ElseIf open_year = 0 Then
                                    percent_change = "N/A"
                                
                                Else
                                    percent_change = yearly_change / open_year
                                    
                                    ws.Cells(row_count, 12).Value = percent_change
                                    
                                    ws.Cells(row_count, 12).NumberFormat = "0.00%"
                                    
                                End If
                      
                        'Move down row in table for each change in ticker
                        row_count = row_count + 1
                        
                        'reset values for each loop
                        total_volume = 0
                        open_year = 0
                        close_year = 0
                        yearly_change = 0
                        percent_change = 0
                        
                    End If
                        
                           
            Next i
                
Next ws

End Sub

