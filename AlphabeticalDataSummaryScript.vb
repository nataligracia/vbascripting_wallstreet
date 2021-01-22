    Function WorksheetExists(sName As String) As Boolean
    WorksheetExists = Evaluate("ISREF('" & sName & "'!A1)")
    
End Function

Sub vbascriptingwallstreet():

    'create and place summary sheet
    If Not WorksheetExists("Summary") Then
    Sheets.Add.Name = "Summary"
    Sheets("Summary").Move After:=Sheets(8)
    End If
    
    Set summary_sheet = Worksheets("Summary")

    summary_row = 2

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
            
        If ws.Name <> ("Summary") Then
            
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
                        summary_sheet.Cells(summary_row, 9).Value = ticker_stock
                        
                        'calculate and place yearly change
                        open_price = ws.Cells(openprice_row, 3).Value
                        close_price = ws.Cells(Row, 6).Value
                        yearly_change = close_price - open_price
                        ws.Cells(analysis_row, 10).Value = yearly_change
                        summary_sheet.Cells(summary_row, 10).Value = yearly_change
                        
                        'formatting yearly_change, Red 255/0/0 for negative & Green0/255/0 for positive
                        If yearly_change <= 0 Then
                            ws.Cells(analysis_row, 10).Interior.Color = RGB(255, 0, 0)
                            summary_sheet.Cells(summary_row, 10).Interior.Color = RGB(255, 0, 0)
                        Else
                            ws.Cells(analysis_row, 10).Interior.Color = RGB(0, 255, 0)
                            summary_sheet.Cells(summary_row, 10).Interior.Color = RGB(0, 255, 0)
                        End If
                        
                        'calculate percent change, place it, and fix overflow error
                        If open_price = 0 Then
                            ws.Cells(analysis_row, 11).Value = 0
                            summary_sheet.Cells(summary_row, 11).Value = 0
                        Else
                            percent_change = yearly_change / open_price
                            ws.Cells(analysis_row, 11).Value = percent_change
                            summary_sheet.Cells(summary_row, 11).Value = percent_change
                        End If
                        
                        'formatting percent change and stock summary table results
                         ws.Cells(analysis_row, 11).NumberFormat = "0.00%"
                         ws.Cells(greatestincrease_row, 17).NumberFormat = "0.00%"
                         ws.Cells(greatestdecrease_row, 17).NumberFormat = "0.00%"
                        summary_sheet.Cells(summary_row, 11).NumberFormat = "0.00%"
                        
                        'place total stock volume
                        ws.Cells(analysis_row, 12).Value = total_volume
                        summary_sheet.Cells(summary_row, 12).Value = total_volume
                        
                        'formatting total stock volume and stock summary table results
                        ws.Cells(analysis_row, 12).NumberFormat = "#,###"
                        ws.Cells(analysis_row, 17).NumberFormat = "#,###"
                        summary_sheet.Cells(summary_row, 12).NumberFormat = "#,###"
                        summary_sheet.Cells(summary_row, 17).NumberFormat = "#,###"
                        
                        'reset total stock volume
                        total_volume = 0
                        analysis_row = analysis_row + 1
                        summary_row = summary_row + 1
                        
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

        End If

    Next ws
    
            'loop through sheet and grab values for summary of stock table
            sum_end = summary_sheet.Cells(summary_row, 9).End(xlUp).Row
            
                For SumRow = 2 To sum_end
            
                    'grab and place greast percentage increase amount and ticker
                    If summary_sheet.Cells(SumRow, 11).Value >= GreatestIncrease Then
                            GreatestIncreaseTicker = summary_sheet.Cells(SumRow, 9)
                            GreatestIncrease = summary_sheet.Cells(SumRow, 11)
                            summary_sheet.Cells(greatestincrease_row, 16).Value = GreatestIncreaseTicker
                            summary_sheet.Cells(greatestincrease_row, 17).Value = GreatestIncrease
                    End If
                    
                    'grab and place greast percentage decrease amount and ticker
                    If summary_sheet.Cells(SumRow, 11).Value <= GreatestDecrease Then
                            GreatestDecreaseTicker = summary_sheet.Cells(SumRow, 9)
                            GreatestDecrease = summary_sheet.Cells(SumRow, 11)
                            summary_sheet.Cells(greatestdecrease_row, 16).Value = GreatestDecreaseTicker
                            summary_sheet.Cells(greatestdecrease_row, 17).Value = GreatestDecrease
                    End If
            
                    'grab and place greast volume amount and ticker
                    If summary_sheet.Cells(SumRow, 12).Value >= GreatestVolume Then
                            GreatestVolumeTicker = summary_sheet.Cells(SumRow, 9)
                            GreatestVolume = summary_sheet.Cells(SumRow, 12)
                            summary_sheet.Cells(greatestvolume_row, 16).Value = GreatestVolumeTicker
                            summary_sheet.Cells(greatestvolume_row, 17).Value = GreatestVolume
                    End If
            
                    'formatting summary stack table
                    summary_sheet.Range("I1:Q" & summary_row).Columns.AutoFit
            
                Next SumRow
                
            'remove extra columns in summary sheet
            summary_sheet.Columns("A:H").EntireColumn.Delete

End Sub
