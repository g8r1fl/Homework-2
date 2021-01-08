Sub Stock_lookup()


'Set loop to run code on each worksheet
Dim ws as worksheet
Dim starting_ws as worksheet
Set starting_ws = ActiveSheet

For Each ws in ThisWorkbook.Worksheets
    ws.Activate

        'Set initial variables
        Dim lRow As Long
        Dim lCol As Long
        Dim Stock_symbol As String
        Dim Open_price As Double
        Dim Close_price As Double
        Dim Yearly_change As Double
        Dim Percentage_change As Double
        ' Changed Stock volume data type as LongLong due to "runtime error 6 overflow" error
        Dim Stock_volume As LongLong
        
        'Start stock volumes at zero
        Stock_volume = 0
        
        'set first open price of data to print in summary table
        Open_price = Cells(2, 3).Value
        

        'Keep track of the location for each stock symbol in the summary
        Dim Summary_row As Integer
        Summary_row = 2

        'Find last row and column
        lRow = Cells(Rows.Count, 1).End(xlUp).Row
        lCol = Cells(1, Columns.Count).End(xlToLeft).Column

        'Create summary table
        Range("I1") = "Ticker"
        Range("J1") = "Yearly Change"
        Range("K1") = "Percentage Change"
        Range("L1") = "Total Stock Volume"
        Range("M1") = "Opening Price"
        Range("N1") = "Closing Price"
        Range("P1") = "Ticker"
        Range("Q1") = "Value"
        Range("O2") = "Greatest % increase"
        Range("O3") = "Greatest % decrease"
        Range("O4") = "Greatest total volume"

        'print first open price
        Range("M2") = Open_price



        'Loop through all rows and columns
        For I = 2 To lRow

            'Check if we are still in the same stock symbol, if not
                If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

                    'Set the stock symbol
                    Stock_symbol = Cells(I, 1).Value

                    'print stock symbol in summary table
                    Range("I" & Summary_row).Value = Stock_symbol
                    
                    'calculate yearly change
                    Yearly_change = Cells(I, 6).Value - Open_price
                    
                    'Print yearly change in summary table
                    Range("J" & Summary_row).Value = Yearly_change
                    
                    'calculate percentage change if denominator not zero
                    If Open_price = 0 Then
                        Range("K" & Summary_row) = ""
                    Else    
                        Range("K" & Summary_row) = Format((Yearly_change / Open_price), "0.00%")
                    End If    
                    
                    'highlight percentages based on
                    If Range("K" & Summary_row).Value < 0 Then
                        Range("K" & Summary_row).Interior.ColorIndex = 3
                    ElseIf Range("K" & Summary_row).Value > 0 Then
                        Range("K" & Summary_row).Interior.ColorIndex = 4
                    Else 
                        Range("K" & Summary_row).Interior.ColorIndex = 2
                    End If
                        
                    'Set closing price
                    Close_price = Cells(I, 6).Value
                            
                    'Add to the stock volume
                    Stock_volume = Stock_volume + Cells(I, 7).Value
                    
                    'print the total volume in summmary table
                    Range("L" & Summary_row).Value = Stock_volume
                    
                    'print opening price
                    Range("M" & Summary_row) = Open_price
                    
                    'print closing price
                    Range("N" & Summary_row) = Close_price

                    'set opening price
                    Open_price = Cells(i +1, 3).Value
                    
                    'Add one row to the summary table
                    Summary_row = Summary_row + 1

                    'reset the stock volume
                    Stock_volume = 0

                ' if the cell immediately following a row is the same symbol
                Else

                    'add to the stock volume
                    Stock_volume = Stock_volume + Cells(I, 7).Value
                    
                End If
                
                
        Next I

        'Bonus
        'set variables for bonus
        Dim Summary_lrow As Long
        Dim Greatest_increase As Double
        Dim Greatest_decrease As Double
        Dim Greatest_volume As Double

        'Find last row of summary table
        Summary_lrow = Cells(1, 11).End(xlDown).Row

        'Set greatest % increase
        Greatest_increase = Application.WorksheetFunction.Max(Range(Cells(2, 11), Cells(Summary_lrow, 11)))

        'Print values
        Range("Q2") = Format(Greatest_increase, "0.00%")

        'Set greatest % decrease
        Greatest_decrease = Application.WorksheetFunction.Min(Range(Cells(2, 11), Cells(Summary_lrow, 11)))

        'Print values
        Range("Q3") = Format(Greatest_decrease, "0.00%")

        'Set greatest total volume
        Greatest_volume = Application.WorksheetFunction.Max(Range(Cells(2, 12), Cells(Summary_lrow, 12)))

        'Print values
        Range("Q4") = Format(Greatest_volume, "Scientific")

        'Print ticker symbols to corresponding metric
        For I = 2 To Summary_lrow

            If Cells(I, 11) = Greatest_increase Then
                Range("P2") = Cells(I, 9)
            ElseIf Cells(I, 11) = Greatest_decrease Then
                Range("P3") = Cells(I, 9)
            ElseIf Cells(I, 12) = Greatest_volume Then
                Range("P4") = Cells(I, 9)
            End If
            
        Next I

            

        Columns("O").EntireColumn.AutoFit

Next

starting_ws.Activate

End Sub