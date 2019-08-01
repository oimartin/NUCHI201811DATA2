Sub stockdata_moderate()

'loop through each year of stock data
'grab total amount of volume each tock had over the year
'grab yearly change from what the stock opened the year
' at to what the closing price was
'grab the percent change from the what it opened the year
' at to what it closed.

Dim ticker_name As String

'Define total stock vol; set to 0
Dim total_stock_vol As Double
total_stock_vol = 0

'Define yearly change, open and close prices
Dim yearly_change As Double
Dim open_price As Double
open_price = Range("C2").Value
Dim close_price As Double

'The percent change from the what it
' opened the year at to what it closed.
Dim percent_change As Double

'Define row in ouptut table for new information
Dim summary_table_row As Integer
summary_table_row = 2

'Define what the last row number is in the sheet
Dim lastrow As Long
lastrow = Range("A:A").SpecialCells(xlLastCell).Row

'Loop to go through all rows in table
For i = 2 To lastrow

    'if next cell doesn't match current ticker_name then
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        'Record the ticker name and put in summary output table
        ticker_name = Cells(i, 1).Value
        Range("H" & summary_table_row).Value = ticker_name

        'add stock volume from 7th column to total_stock_vol
        total_stock_vol = total_stock_vol + Cells(i, 7).Value
        Range("I" & summary_table_row).Value = total_stock_vol

        'Set close price to be the value in the last common row
        close_price = Cells(i, 6).Value
        'MsgBox (close_price & " " & open_price)

        'Set yearly stock price change and put in summary table
        yearly_change = close_price - open_price
        Range("J" & summary_table_row).Value = yearly_change

        'percent yearly change
        If open_price = 0 Then

        'when open price is 0, percent change is the close price * 100 as a percent
        Range("K" & summary_table_row).Value = Round((close_price * 100), 2) & "%"
        
        'when yearly change is 0, percent change is 0% 
        ElseIf yearly_change = 0 Then
        Range("K" & summary_table_row).Value = 0
        
        'otherwise, use the normal formula for percent change 
        ElseIf open_price <> 0 Then
            percent_change = Round(((yearly_change / open_price) * 100), 2)
            Range("K" & summary_table_row).Value = percent_change & "%"
        End If

        'Now, increase summary table by 1 to go to next row
        summary_table_row = summary_table_row + 1

        'reset total stock vol to 0 bc done summing ticker_name vol
        total_stock_vol = 0
        
        'Set the next open price to be
        'the next row in col 3
        open_price = Cells(i + 1, 3)
    Else

        'if ticker_name is the same as next cell
        'add next value of total stock vol to current count
        total_stock_vol = total_stock_vol + Cells(i, 7).Value
    
    End If

Next i

'Loop to go through all rows in table..
For i = 2 To lastrow

    'If yearly percentage change is greater than 0...
    If Cells(i, 11).Value > 0 Then
        
        'Color cells > 0 green
        Cells(i, 11).Interior.ColorIndex = 4

    'if yearly percentage change is less than 0...
    ElseIf Cells(i, 11).Value < 0 Then

        'Color cell < 0 red, otherwise, no color change
        Cells(i, 11).Interior.ColorIndex = 3
    End If
Next i


End Sub

