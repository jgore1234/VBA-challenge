Attribute VB_Name = "Module1"
Sub stockanalysis()

For Each ws In Worksheets


    total_volume = 0
    summary_table_row = 2
    open_row = 2
   ws.Range("N1").Value = "ticker"
   ws.Range("O1").Value = "yearly change"
   ws.Range("P1").Value = "Percent Change"
   ws.Range("Q1").Value = "Total Volume"
   
    ' loop through and grab values for a ticker
    For Row = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'calculate total volume
        total_volume = total_volume + ws.Cells(Row, 7).Value
        If ws.Cells(Row, 1).Value <> ws.Cells(Row + 1, 1).Value Then
        
            'grab ticker
            ticker = ws.Cells(Row, 1).Value
            ws.Cells(summary_table_row, 14).Value = ticker
            
            'calculate yearly change
            open_value = ws.Cells(open_row, 3).Value
            close_value = ws.Cells(Row, 6).Value
            yearly_change = close_value - open_value
            ws.Cells(summary_table_row, 15).Value = yearly_change
            
               'Color
                If yearly_change <= 0 Then
                    'color red
                    ws.Cells(summary_table_row, 15).Interior.Color = RGB(255, 0, 0)
               Else
            
                    'color green
                    ws.Cells(summary_table_row, 15).Interior.Color = RGB(0, 255, 0)
                End If
            
            'calculate percent change
            If open_value = 0 Then
            percent_change = 0
            Else: ws.Cells(summary_table_row, 16).Value = yearly_change / open_value
            End If
            
           ws.Cells(summary_table_row, 17).Value = total_volume
           total_volume = 0
           summary_table_row = summary_table_row + 1
           
           
                
                'formating
                ws.Cells(summary_table_row, 16).NumberFormat = "0.00%"
                
                'set open row to be open to next ticker open
            open_row = Row + 1
        End If
    Next Row
    Next ws
End Sub

