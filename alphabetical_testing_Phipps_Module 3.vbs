Attribute VB_Name = "Module3"
Sub ws_ticker()

Dim w As Worksheet
For Each w In Worksheets

    Dim column As Integer
    column = 1
    
    Dim ticker_name As String
    
    'Keep track of opening price
    Dim Opening_Price As Double
    Opening_Price = 0
    
    'Opening Price counter
    Dim op_counter As Double
    op_counter = 0
    
    'Keep track of closing price
    Dim Closing_Price As Double
    Closing_Price = 0
    
    'Keep track of percentage change
    Dim Percent_Change As String
    Percent_Change = 0
    
    'Total stock volume
    Dim Total_Stock_Volume As LongLong
    Total_Stock_Volume = 0
    
    'Keep track of what row we're on in the output table
    Dim Summary_Row_Table As Integer
    Summary_Row_Table = 2
    
    lastrow = w.Cells(Rows.Count, 1).End(xlUp).Row
    
    w.Range("I1").Value = "Ticker"
    w.Range("J1").Value = "Yearly Change"
    w.Range("K1").Value = "Percentage Change"
    w.Range("L1").Value = "Total Stock Volume"
    
    For i = 2 To lastrow
    
        If w.Cells(i + 1, column).Value <> w.Cells(i, column).Value Then
        
        'Set ticker name
        ticker_name = w.Cells(i, column).Value
        
        'Set closing price
        Closing_Price = w.Cells(i, 3).Value
        
        'Calculate Yearly change
        Yearly_Change = Closing_Price - Opening_Price
        
        'Calculate Percentage Change
        Percent_Change = (Round((Yearly_Change / Opening_Price * 100), 2)) & "%"
        
        'Print ticker name
        w.Cells(Summary_Row_Table, 9).Value = ticker_name
        
        'Print yearly change
        w.Cells(Summary_Row_Table, 10).Value = Yearly_Change
        If Yearly_Change < 0 Then
        w.Cells(Summary_Row_Table, 10).Interior.ColorIndex = 3
        Else
        w.Cells(Summary_Row_Table, 10).Interior.ColorIndex = 4
        End If
        
        'Print percentage change
        w.Cells(Summary_Row_Table, 11).Value = Percent_Change
        
        'Print total stock volume
        w.Cells(Summary_Row_Table, 12).Value = Total_Stock_Volume
        
        'Move to the next summary row
        Summary_Row_Table = Summary_Row_Table + 1
        
        'reset Yearly change to zero
        Yearly_Change = 0
        
        'Reset Opening price to zero
        Opening_Price = 0
        
        'Reset Closing_Price to zero
        Closing_Price = 0
        
        'Reset Percent_Change to zero
        Percent_Change = 0
        
        'Reset Total_Stock_Volume to zero
        Total_Stock_Volume = 0
        
        'Reset op_counter
        op_counter = 0
        
            'if the tickers are the same
            Else
            
            'Add total stock volume
            Total_Stock_Volume = Total_Stock_Volume + w.Cells(i, 7).Value
            
            op_counter = op_counter + 1
            
            If op_counter = 1 Then
            
            'Set open price
            Opening_Price = w.Cells(i, 3).Value
        
            End If
    
        End If
    
    Next i
    
    Dim percent_inc As String
    Dim percent_dec As String
    Dim percent_range As Range
    Dim greatest_total_volume As LongLong
    Dim greatest_total_range As Range
    
    Set percent_range = w.Range("K2:K" & lastrow)
    Set greatest_total_range = w.Range("L1:L" & lastrow)
        
    w.Range("P1").Value = "Ticker"
    w.Range("Q1").Value = "Value"
    w.Cells(2, 15).Value = "Greatest % Increase"
    w.Cells(3, 15).Value = "Greatest % Decrease"
    w.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Find Max of Percentage change
        percent_inc = WorksheetFunction.Max(percent_range) * 100 & "%"
        w.Cells(2, 17).Value = percent_inc
        
        'Find Min of Percentage Change
        percent_dec = WorksheetFunction.Min(percent_range) * 100 & "%"
        w.Cells(3, 17).Value = percent_dec
        
        'Find Greatest total volume of total stock volume
        greatest_total_volume = WorksheetFunction.Max(greatest_total_range)
        w.Cells(4, 17).Value = greatest_total_volume
         
    For i = 2 To lastrow
    
    If w.Cells(2, 17).Value = w.Cells(i, 11).Value Then
    ticker_name = w.Cells(i, 9).Value
    w.Cells(2, 16).Value = ticker_name
    
    End If
    
    Next i
    
    For i = 2 To lastrow
    
    If w.Cells(3, 17).Value = w.Cells(i, 11).Value Then
    ticker_name = w.Cells(i, 9).Value
    w.Cells(3, 16).Value = ticker_name
    
    End If
    
    Next i
    
    For i = 2 To lastrow
    
    If w.Cells(4, 17).Value = w.Cells(i, 12).Value Then
    ticker_name = w.Cells(i, 9).Value
    w.Cells(4, 16).Value = ticker_name
    
    End If
    
    Next i

Next w

End Sub


