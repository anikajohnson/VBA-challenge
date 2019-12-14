Attribute VB_Name = "Module1"
Sub Yearly_Stock_Overview()

    ' set worksheet variable
    Dim ws As Worksheet
    
'Start loop
For Each ws In Worksheets

    'set last row
    Dim lastRow As Long
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim Ticker_Name As String
    
    'where summary table starts
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    'summary table titles
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'best/worst table titles
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
            
    'define variable types
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Volume As Double
    
    'set counters to zero
    Opening_Price = 0
    Closing_Price = 0
    Yearly_Change = 0
    Percent_Change = 0
    Total_Volume = 0
 
    ' Loop from the beginning of the current worksheet
    For i = 2 To lastRow
    
        'Set opening price
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
          
            Opening_Price = ws.Cells(i, 3).Value

        End If
    
        ' Add to the Volume Total
        Total_Volume = Total_Volume + ws.Cells(i, 7).Value
          
       'if ticker name changes
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        
            'Set Ticker name
            Ticker_Name = ws.Cells(i, 1).Value
                
            'Print the TickerName in the Summary Table
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
        
            ' Print the Total Volume Amount to the Summary Table
            ws.Range("L" & Summary_Table_Row).Value = Total_Volume
            
            'set year end price
            Closing_Price = ws.Cells(i, 6).Value
            
          ' Calculate Yearly_Change and Percent_Change
            Yearly_Change = Closing_Price - Opening_Price
            Percent_Change = Opening_Price + Closing_Price / 2
            
          
          ' Print the Total Yearly Change Amount to the Summary Table
          ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
          
          If (Yearly_Change > 0) Then
                'Make green
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    
          Else
                'Make red
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    
          End If
                    
          ' Print the percent change in the Summary Table
          ws.Range("K" & Summary_Table_Row).Value = (CStr(Percent_Change) & "%")
             
          'Add one to the summary table row
          Summary_Table_Row = Summary_Table_Row + 1
                      
        ' Reset the Totals
        Opening_Price = 0
        Closing_Price = 0
        Yearly_Change = 0
        Percent_Change = 0
        Total_Volume = 0
                
        End If
          
Next i

    'set last row
    lastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'set variables
    Dim Greatest_Stock As String
    Dim Greatest_Value As Double
    Dim Worst_Stock As String
    Dim Worst_Value As Double
    Dim Biggest_Vol_Stock As String
    Dim Biggest_Vol_Value As Double
    
    'set best/worst values to first stock in list
    Greatest_Value = ws.Cells(2, 11).Value
    Worst_Value = ws.Cells(2, 11).Value
    Biggest_Vol_Value = ws.Cells(2, 12).Value
    
    
For j = 2 To lastRow

    'If the current value is greater than the greatest_value
    'set current value as new greatest_value and record stock name
    If ws.Cells(j, 11).Value > Greatest_Value Then
        Greatest_Value = ws.Cells(j, 11).Value
        Greatest_Stock = ws.Cells(j, 9).Value
    End If
         
     'If current value is worst that the worst_value
     'set current value as new worst_value and record stock name
    If ws.Cells(j, 11).Value > Worst_Value Then
        Worst_Value = ws.Cells(j, 11).Value
        Worst_Stock = ws.Cells(j, 9).Value
    End If
        
     'If current value is bigger that the biggest_vol_value
     'set current value as new biggest_vol_value and record stock name
    If ws.Cells(j, 11).Value > Biggest_Vol_Value Then
        Biggest_Vol_Value = ws.Cells(j, 11).Value
        Biggest_Vol_Stock = ws.Cells(j, 9).Value
    End If
                      
 Next j
 
        'place values in table
        ws.Cells(2, 16).Value = Greatest_Stock
        ws.Cells(2, 17).Value = Greatest_Value
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 16).Value = Worst_Stock
        ws.Cells(3, 17).Value = Worst_Value
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 16).Value = Biggest_Vol_Stock
        ws.Cells(4, 17).Value = Biggest_Vol_Value

    Next ws
    
End Sub
