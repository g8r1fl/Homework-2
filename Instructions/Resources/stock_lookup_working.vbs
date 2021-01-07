Sub Stock_lookup()

'Set initial variables
Dim lRow As Long
Dim lCol As Long
Dim Stock_symbol As String
' Changed Stock volume data type as LongLong due to "runtime error 6 overflow" error
Dim Stock_volume As LongLong
Stock_volume = 0
Dim Open_price As Currency
'set first open price of data to print in summary table
Open_price = Cells(2, 3).Value
Dim Close_price As Currency
Dim Yearly_change As Double
Dim Percentage_change As Double

'Keep track of the location for each stock symbol in the summary
Dim Summary_row As Integer
Summary_row = 2

'Find last row and column
lRow = Cells(Rows.Count, 1).End(xlUp).Row
lCol = Cells(1, Columns.Count).End(xlToLeft).Column

'MsgBox "Last row is " & lRow & vbNewLine & _
        "Last column is " & lCol

'Create summary table
Range("I1") = "Ticker"
Range("J1") = "Yearly Change"
Range("K1") = "Percentage Change"
Range("L1") = "Total Stock Volume"
Range("M1") = "Opening Price"
Range("N1") = "Closing Price"
Range("M2") = Open_price
Range("Q1") = "Ticker"
Range("R1") = "Value"
Range("P2") = "Greatest % increase"
Range("P3") = "Greatest % decrease"
Range("P4") = "Greatest total volume"



'Loop through all rows and columns
For I = 2 To lRow

    'Check if we are still in the same stock symbol, if not
        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

            'Set the stock symbol
            Stock_symbol = Cells(I, 1).Value
            
            'calculate yearly change
            Yearly_change = Cells(I, 6).Value - Open_price
            
            'Print yearly change in summary table
            Range("J" & Summary_row).Value = Yearly_change
            
            'calculate percentage change
            Percentage_change = (Yearly_change / Open_price)
            
            'Print percent change to summary table
            Range("K" & Summary_row).Value = Format(Percentage_change, "0.00%")
            
            If Range("K" & Summary_row).Value < 0 Then
                Range("K" & Summary_row).Interior.ColorIndex = 3
            Else
                Range("K" & Summary_row).Interior.ColorIndex = 4
            End If
                
            'Set closing price
            Close_price = Cells(I, 6).Value
                    
            
                    
            'Add to the stock volume
            Stock_volume = Stock_volume + Cells(I, 7).Value
            
            'print stock symbol in summary table
            Range("I" & Summary_row).Value = Stock_symbol
            
            'print the total volume in summmary table
            Range("L" & Summary_row).Value = Stock_volume
            
            'print opening price
            Range("M" & Summary_row) = Open_price
            
            'Print closing; Price
            Range("N" & Summary_row) = Close_price
            
            'set opening price
            Open_price = Cells(I + 1, 3).Value
            
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

'MsgBox "Last row is " & Summary_lrow

'Set greatest % increase
Greatest_increase = Application.WorksheetFunction.Max(Range(Cells(2, 11), Cells(Summary_lrow, 11)))

'Print values
Range("R2") = Format(Greatest_increase, "0.00%")

'Set greatest % decrease
Greatest_decrease = Application.WorksheetFunction.Min(Range(Cells(2, 11), Cells(Summary_lrow, 11)))

'Print values
Range("R3") = Format(Greatest_decrease, "0.00%")

'Set greatest total volume
Greatest_volume = Application.WorksheetFunction.Max(Range(Cells(2, 12), Cells(Summary_lrow, 12)))

'Print values
Range("R4") = Format(Greatest_volume, "Scientific")

For I = 2 To Summary_lrow

    If Cells(I, 11) = Greatest_increase Then
        Range("Q2") = Cells(I, 9)
    ElseIf Cells(I, 11) = Greatest_decrease Then
        Range("Q3") = Cells(I, 9)
    ElseIf Cells(I, 12) = Greatest_volume Then
        Range("Q4") = Cells(I, 9)
    End If
    
Next I

    

Columns("P").EntireColumn.AutoFit

End Sub




