Sub StockAnalysis()

'##############################################################################################################
'Create a script that will loop through all the stocks for one year and output
'The ticker symbol.
'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'The total stock volume of the stock.
'##############################################################################################################

'set headers
Cells(1, 10) = "Ticker"
Cells(1, 11).Value = "Yearly Change"
Cells(1, 12).Value = "Percent Change"
Cells(1, 13).Value = "Total Stock Volume"

'Declaring Variables
Dim Ticker As String
Dim Yearly_Change, year_open, Percent_Change, Stock_Vol As Double
Dim i As Long
Dim current_row As Long
Stock_Vol = 0
Yearly_Change = 0
current_row = 2
'Count Number of rows in Worksheet
NumRows = Range("A2", Range("A2").End(xlDown)).Rows.Count

For i = 2 To NumRows
year_open = Cells(current_row, 3).Value
    'Checking if within the same stock
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker = Cells(i, 1).Value
        Range("J" & current_row).Value = Ticker
        
        Yearly_Change = Yearly_Change + Cells(i, 6).Value - year_open
        Range("K" & current_row).Value = Yearly_Change
        
        Percent_Change = (Yearly_Change / year_open)
        Range("L" & current_row).Value = Percent_Change
        
        'Vol = Cells(i, 7).Value
        'Total_Stock_Vol = Vol + Stock_Vol
        Stock_Vol = Stock_Vol + Cells(i, 7).Value
        Range("M" & current_row).Value = Stock_Vol
        
        Stock_Vol = Total_Stock_Vol
        current_row = current_row + 1
        Yearly_Change = 0
        Stock_Vol = 0
        year_open = Cells(current_row, 3).Value
    Else: Stock_Vol = Stock_Vol + Cells(i, 7).Value
    End If
Next i

End Sub




