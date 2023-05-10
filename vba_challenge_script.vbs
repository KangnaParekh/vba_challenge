Sub MULTIPLE_STOCK()

Dim ticker_name As String
Dim open_Price As Double
Dim close_Price As Double
Dim yearly_Change As Double
Dim percent_Change As Double
Dim total_Stock_Volume As Long
Dim open_Price_Row As Long


total_Stock_Volume = 0


Dim summary_table_row As Integer


For Each ws In Worksheets
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly_Change"
    ws.Range("K1").Value = "Percent_Change"
    ws.Range("L1").Value = "Total_Stock_Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest_Total_Volume"

    summary_table_row = 2
    open_Price_Row = 2

    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        
        On Error Resume Next
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            ticker = ws.Cells(i, 1).Value
            
            open_Price = ws.Cells(open_Price_Row, 3).Value
            
            close_Price = ws.Cells(i, 6).Value
            
            yearly_Change = close_Price - open_Price
            
            
            If open_Price = 0 Then
            
                percent_Change = 0
            Else
                
                percent_Change = yearly_Change / open_Price
            End If
            
            
            total_Stock_Volume = total_Stock_Volume + ws.Cells(i, 7).Value
            
            ws.Range("I" & summary_table_row).Value = ticker
            
            ws.Range("J" & summary_table_row).Value = yearly_Change
            
            If ws.Range("J" & summary_table_row).Value > 0 Then
                ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
            Else
                ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
            End If
                
            ws.Range("K" & summary_table_row).Value = percent_Change
            
            ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
            
            ws.Range("L" & summary_table_row).Value = total_Stock_Volume
            
            summary_table_row = summary_table_row + 1
            
            open_Price_Row = i + 1
            
            total_Stock_Volume = 0
        Else
            total_Stock_Volume = total_Stock_Volume + ws.Cells(i, 7).Value
        End If
    Next i
    
    
     aRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
    min_Value = 0
    max_Value = 0
    max_Total_Volume = 0
    
    For i = 2 To aRow
    
        If ws.Cells(i, 11) > max_Value Then
            max_Value = ws.Cells(i, 11)
            max_Ticker = ws.Cells(i, 9)
        Else
            max_Value = max_Value
        End If
        
        If ws.Cells(i, 11) < minValue Then
            min_Value = ws.Cells(i, 11)
            min_Ticker = ws.Cells(i, 9)
        Else
            min_Value = min_Value
        End If
        
        If ws.Cells(i, 12) > max_Total_Volume Then
            max_Total_Volume = ws.Cells(i, 12)
            max_Total_Volume_Ticker = ws.Cells(i, 9)
        Else
            max_Total_Volume = max_Total_Volume
        End If
        
    Next i
    
    ws.Range("P2").Value = max_Ticker
    ws.Range("Q2").Value = max_Value
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("P3").Value = min_Ticker
    ws.Range("Q3").Value = min_Value
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("P4").Value = max_Total_Volume_Ticker
    ws.Range("Q4").Value = max_Total_Volume
    
    ws.Columns("I:Q").AutoFit
    
Next ws

End Sub
