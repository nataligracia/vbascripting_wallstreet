Sub vbascriptingwallstreet():

    For Each ws In Worksheets
    
            'set total start
            total_volume = 0
            analysis_row = 2
            openprice_row = 2
            greatestincrease_row = 2
            greatestdecrease_row = 3
            greatestvolume_row = 4
            
            'name analysis columns
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Volume"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
            ws.Range("I1:Q1").Font.Bold = True
            ws.Range("I1:Q1").Columns.AutoFit
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            ws.Range("O2:O4").Font.Bold = True
            
            'loop through sheet and grab values of stock
            stock_end = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
            For Row = 2 To stock_end
                
                'calculate total stock volume
                total_volume = total_volume + ws.Cells(Row, 7).Value
                
                    'to do list per stock
                    If ws.Cells(Row, 1).Value <> ws.Cells(Row + 1, 1).Value Then
                    
                        'grab and place stock
                        ticker_stock = ws.Cells(Row, 1).Value
                        ws.Cells(analysis_row, 9).Value = ticker_stock
                        
                        'calculate and place yearly change
                        open_price = ws.Cells(openprice_row, 3).Value
                        close_price = ws.Cells(Row, 6).Value
                        yearly_change = close_price - open_price
                        ws.Cells(analysis_row, 10).Value = yearly_change
                        
                        'formatting yearly_change, Red 255/0/0 for negative & Green0/255/0 for positive
                        If yearly_change <= 0 Then
                            ws.Cells(analysis_row, 10).Interior.Color = RGB(255, 0, 0)
                        Else
                            ws.Cells(analysis_row, 10).Interior.Color = RGB(0, 255, 0)
                        End If
                        
                        'calculate percent change, place it, and fix overflow error
                        If open_price = 0 Then
                            ws.Cells(analysis_row, 11).Value = 0
                        Else
                            percent_change = yearly_change / open_price
                            ws.Cells(analysis_row, 11).Value = percent_change
                        End If
                        
                        'formatting percent change and stock summary table results
                         ws.Cells(analysis_row, 11).NumberFormat = "0.00%"
                         ws.Cells(greatestincrease_row, 17).NumberFormat = "0.00%"
                         ws.Cells(greatestdecrease_row, 17).NumberFormat = "0.00%"
                        
                        'placing total stock volume
                        ws.Cells(analysis_row, 12).Value = total_volume
                        
                        'formatting total stock volume and stock summary table results
                        ws.Cells(analysis_row, 12).NumberFormat = "#,###"
                        ws.Cells(analysis_row, 17).NumberFormat = "#,###"
                        
                        'reset total stock volume
                        total_volume = 0
                        analysis_row = analysis_row + 1
                        
                        'reset open price row
                        openprice_row = Row + 1
                        
                      End If
            
                Next Row
                
            'loop through sheet and grab values for summary of stock table
            sum_end = ws.Cells(analysis_row, 9).End(xlUp).Row
            
                For SumRow = 2 To sum_end
            
                    'grab and place greast percentage increase amount and ticker
                    If ws.Cells(SumRow, 11).Value >= GreatestIncrease Then
                            GreatestIncreaseTicker = ws.Cells(SumRow, 9)
                            GreatestIncrease = ws.Cells(SumRow, 11)
                            ws.Cells(greatestincrease_row, 16).Value = GreatestIncreaseTicker
                            ws.Cells(greatestincrease_row, 17).Value = GreatestIncrease
                    End If
                    
                    'grab and place greast percentage decrease amount and ticker
                    If ws.Cells(SumRow, 11).Value <= GreatestDecrease Then
                            GreatestDecreaseTicker = ws.Cells(SumRow, 9)
                            GreatestDecrease = ws.Cells(SumRow, 11)
                            ws.Cells(greatestdecrease_row, 16).Value = GreatestDecreaseTicker
                            ws.Cells(greatestdecrease_row, 17).Value = GreatestDecrease
                    End If
            
                    'grab and place greast volume amount and ticker
                    If ws.Cells(SumRow, 12).Value >= GreatestVolume Then
                            GreatestVolumeTicker = ws.Cells(SumRow, 9)
                            GreatestVolume = ws.Cells(SumRow, 12)
                            ws.Cells(greatestvolume_row, 16).Value = GreatestVolumeTicker
                            ws.Cells(greatestvolume_row, 17).Value = GreatestVolume
                    End If
            
                Next SumRow
                
                    'reset summary table
                    GreatestIncrease = 0
                    GreatestDecrease = 0
                    GreatestVolume = 0
                    
                    'formatting summary stack table
                    ws.Range("O1:Q4").Columns.AutoFit

    Next ws

End Sub