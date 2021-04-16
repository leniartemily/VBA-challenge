Attribute VB_Name = "Module1"
Sub Stock_Analysis()
'Start loop
    For Each WS In Worksheets
'Create a loop to cycle through the worksheets in the workbook

'Set variable to hold total volume of stock traded
    TotalStockVolume = 0
    
'Set Parameters
        Dim rowcount As Long
        rowcount = 2
        Dim year_open As Double
        year_open = 0
        Dim year_close As Double
        year_close = 0
        Dim year_change As Double
        year_change = 0
        Dim percent_change As Double
        percent_change = 0
        
'Create header labels for the summary table
        WS.Cells(1, 9).Value = "Ticker"
        WS.Cells(1, 10).Value = "Yearly Change"
        WS.Cells(1, 11).Value = "Percent Change"
        WS.Cells(1, 12).Value = "Total Stock Volume"

'Determine The last row
    LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
       
       'Rename row by looping through and renaming
        For i = 2 To LastRow

            'Conditional to grab year open value
            If WS.Cells(i, 1).Value <> WS.Cells(i - 1, 1).Value Then
                year_open = WS.Cells(i, 3).Value
            End If
            'Total up the volume for each row to determine the total stock volume for the year
            total_vol = total_vol + WS.Cells(i, 7)
            'Conditional to determine if the ticker symbol is changing
            If WS.Cells(i, 1).Value <> WS.Cells(i + 1, 1).Value Then
                'Move ticker symbol to summary table "Ticker"
                WS.Cells(rowcount, 9).Value = WS.Cells(i, 1).Value
                 '"Total Stock Volume" value
                WS.Cells(rowcount, 12).Value = total_vol
                
                'Grab year end price
                year_close = WS.Cells(i, 6).Value
                '"Yearly Change" price change
                year_change = year_close - year_open
                WS.Cells(rowcount, 10).Value = year_change
                'Conditional to format negative (red=3) and positive (green=4)
                If year_change >= 0 Then
                    WS.Cells(rowcount, 10).Interior.ColorIndex = 4
                Else
                    WS.Cells(rowcount, 10).Interior.ColorIndex = 3
                End If
                
                'Conditional for calculating "percent change" (not divisible by zero because result will be zero)
                If year_open = 0 And year_close = 0 Then
                    percent_change = 0
                    WS.Cells(rowcount, 11).Value = percent_change
                    WS.Cells(rowcount, 11).NumberFormat = "0.00%"
                ElseIf year_open = 0 Then
                    Dim percent_change_NA As String
                    percent_change_NA = "New Stock"
                    WS.Cells(rowcount, 11).Value = percent_change
                Else
                    percent_change = year_change / year_open
                    WS.Cells(rowcount, 11).Value = percent_change
                    WS.Cells(rowcount, 11).NumberFormat = "0.00%"
                End If
                'Add 1 to rowcount to move to empty cell
                rowcount = rowcount + 1
                'Reset parameters
                total_vol = 0
                year_open = 0
                year_close = 0
                year_change = 0
                percent_change = 0
            End If
        Next i
        
'BONUS
        'Create a best/worst performance table
        'Axis Titles
        WS.Cells(2, 15).Value = "Greatest % Increase"
        WS.Cells(3, 15).Value = "Greatest % Decrease"
        WS.Cells(4, 15).Value = "Greatest Total Volume"
        WS.Cells(1, 16).Value = "Ticker"
        WS.Cells(1, 17).Value = "Value"
        
        'Assign lastrow to count the number of rows in the summary table
        LastRow = WS.Cells(Rows.Count, 9).End(xlUp).Row
        'Set Parameters
        Dim best_stock As String
        Dim best_value As Double
        Dim worst_stock As String
        Dim worst_value As Double
        Dim most_vol_stock As String
        Dim most_vol_value As Double
        
        'Set best and worst performer equal to the first stock
        best_value = WS.Cells(2, 11).Value
        worst_value = WS.Cells(2, 11).Value

        'Set most volume equal to the first stock
        most_vol_value = WS.Cells(2, 12).Value
        
        'Loop to search through summary table
        For j = 2 To LastRow
            'Conditional for best performer
            If WS.Cells(j, 11).Value > best_value Then
                best_value = WS.Cells(j, 11).Value
                best_stock = WS.Cells(j, 9).Value
            End If
            'Conditional for worst performer
            If WS.Cells(j, 11).Value < worst_value Then
                worst_value = WS.Cells(j, 11).Value
                worst_stock = WS.Cells(j, 9).Value
            End If
            'Conditional for stock with the greatest volume traded
            If WS.Cells(j, 12).Value > most_vol_value Then
                most_vol_value = WS.Cells(j, 12).Value
                most_vol_stock = WS.Cells(j, 9).Value
            End If
        Next j
        'Move best performer, worst performer, and stock with the most volume items to the performance table
        WS.Cells(2, 16).Value = best_stock
        WS.Cells(2, 17).Value = best_value
        WS.Cells(2, 17).NumberFormat = "0.00%"
        WS.Cells(3, 16).Value = worst_stock
        WS.Cells(3, 17).Value = worst_value
        WS.Cells(3, 17).NumberFormat = "0.00%"
        WS.Cells(4, 16).Value = most_vol_stock
        WS.Cells(4, 17).Value = most_vol_value
        'Autofit table columns
        WS.Columns("I:L").EntireColumn.AutoFit
        WS.Columns("O:Q").EntireColumn.AutoFit
    Next WS
End Sub
