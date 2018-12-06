Sub Alyza_VBA_HW()

'Loop through the worksheets
For Each ws In Worksheets

'Assign variables
Dim stock_name As String

Dim stock_total As Double
stock_total = 0

Dim open_value As Double
Dim close_value As Double
Dim year_change As Double
Dim percent_change As Double

Dim table_row As Long
table_row = 2

'Print headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Total Stock Volume"
ws.Cells(1, 11).Value = "Yearly Change"
ws.Cells(1, 12).Value = "Percent Change"

'Set first open value
open_value = ws.Cells(2, 3).Value

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Loop through all stocks
For i = 2 To lastrow
    
    'Check if we are still looking at same stock. If it is different
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'Set the stock name
            stock_name = ws.Cells(i, 1).Value
            
            'Calculate volume total
            stock_total = stock_total + ws.Cells(i, 7).Value
            
            'Set the close value
            close_value = ws.Cells(i, 6).Value
            
            'Calculate yearly change
            year_change = close_value - open_value
            
            'Calculate percent change (avoid dividing by zero error)
            If open_value = 0 Then
                percent_change = 0
            Else
                percent_change = year_change / open_value
            
            End If
                        
            'Print stock name in table
            ws.Range("I" & table_row).Value = stock_name
                 
            'Print volume total in table
            ws.Range("J" & table_row).Value = stock_total
            
            'Print yearly change in table
            ws.Range("K" & table_row).Value = year_change
            
            'Print yearly change in table
            ws.Range("L" & table_row).Value = percent_change
            
            'Format percent change cells as percent
            ws.Range("L" & table_row).Style = "Percent"
            ws.Range("L" & table_row).NumberFormat = "0.00%"
            
            
            'Format yearly change cells with fill
                If ws.Range("K" & table_row).Value >= 0 Then
                    ws.Range("K" & table_row).Interior.ColorIndex = 4
                
                Else
                    ws.Range("K" & table_row).Interior.ColorIndex = 3
                End If
            
            'Adjust the table row
            table_row = table_row + 1
            
            'Reset volume total
            stock_total = 0
            
            'Set new open value
            open_value = ws.Cells(i + 1, 3).Value
                      
                     
      'If stock is the same
        Else
        stock_total = stock_total + ws.Cells(i, 7).Value
        
    End If
    Next i
    
'Find max values
Dim max_increase As Double
Dim max_decrease As Double
Dim max_volume As Double

max_increase = ws.Cells(2, 12).Value
max_decrease = ws.Cells(2, 12).Value
max_volume = ws.Cells(2, 10).Value

    
last_table_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    'Loop through volume values
    For i = 2 To last_table_row
        
        'Check if next value is higher
        If ws.Cells(i + 1, 10).Value > max_volume Then
        'Replace value if higher
        max_volume = ws.Cells(i + 1, 10).Value
        
        End If
        
    Next i
    
    'Loop through percentage values
    For i = 2 To last_table_row
        
        'Check if next value is higher
        If ws.Cells(i + 1, 12).Value > max_increase Then
        'Replace value if higher
        max_increase = ws.Cells(i + 1, 12).Value
        
        End If
        
    Next i
    
    'Loop through percentage values
    For i = 2 To last_table_row
    
        'Check if next value is lower
        If ws.Cells(i + 1, 12).Value < max_decrease Then
        'Replace value if higher
        max_decrease = ws.Cells(i + 1, 12).Value
        
        
        End If
        
    Next i
        
        
'Print headers and format
ws.Range("M1").Value = "Greatest % Inc"
ws.Range("N1").Value = "Greatest % Dec"
ws.Range("O1").Value = "Greatest Vol"
ws.Range("M2").Value = max_increase
ws.Range("N2").Value = max_decrease
ws.Range("O2").Value = max_volume
ws.Range("M2:N2").Style = "Percent"

Next ws

End Sub
