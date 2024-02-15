Attribute VB_Name = "Module1"

Sub Stock_Analytics():
    For Each ws In Worksheets

    'Set Variables
    Dim ticker_name As String
    Dim row As Long
    Dim row_end As Long
    Dim summary_table_row As Long
    Dim first_opening_price, last_closing_price As Double
    Dim opening_price, closing_price, yearly_change, percent_change, Total_stock_vol As Double

    'Generate headers for new data
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    summary_table_row = 2
    'Define row and row_end
    row = 2
    row_end = ws.Cells(Rows.Count, 1).End(xlUp).row

    'Set original ticker vars
    ticker_name = ws.Cells(row, 1).Value
    Total_stock_vol = 0
    first_opening_price = ws.Cells(row, 3).Value

    'Create loop for whole data set
    For row = 2 To row_end
        'Have total vol increase with each row
        Total_stock_vol = ws.Cells(row, 7).Value + Total_stock_vol
       
        'Change when ticker changes
        If ws.Cells(row, 1).Value <> ws.Cells(row + 1, 1).Value Then
            'capture last closing price
             last_closing_price = ws.Cells(row, 6).Value

           'Calculate yearly, percent, and total vol

            closing_price = ws.Cells(row, 6).Value
            yearly_change = last_closing_price - first_opening_price
            If first_opening_price <> 0 Then
                percent_change = (yearly_change / first_opening_price)
            Else: percent_change = 0
            End If

            'Display values
'            ws.Cells(row, 9).Value = ticker_name
'            ws.Cells(row, 10).Value = yearly_change
'            ws.Cells(row, 11).Value = percent_change
'            ws.Cells(row, 12).Value = Total_stock_vol
                
            'Print to summary table
            ws.Range("I" & summary_table_row).Value = ticker_name
            ws.Range("J" & summary_table_row).Value = yearly_change
            ws.Range("K" & summary_table_row).Value = percent_change
           'apply percent format
            ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
            ws.Range("L" & summary_table_row).Value = Total_stock_vol
    
            'Add one to the summary table row
            summary_table_row = summary_table_row + 1
        
            'Reset variables
            ticker_name = ws.Cells(row + 1, 1).Value
            Total_stock_vol = 0
            first_opening_price = ws.Cells(row + 1, 3).Value
        End If

        Next row
    Next ws
End Sub

Sub Conditionals():
For Each ws In Worksheets
    'Create parameters for summary table
Dim row_sub As Long
Dim row_subend As Long
    row_sub = 1
    row_subend = ws.Cells(Rows.Count, 11).End(xlUp).row
    
    For row_sub = 2 To row_subend
    
    'Create loop for color change
    If ws.Cells(row_sub, 11).Value <= 0 Then
        ws.Cells(row_sub, 11).Interior.Color = vbRed
    ElseIf ws.Cells(row_sub, 11).Value > 0 Then
        ws.Cells(row_sub, 11).Interior.Color = vbGreen
End If
Next row_sub

Next ws
End Sub

Sub Create_Columns():
For Each ws In Worksheets
    'Create labels
ws.Cells(2, 15).Value = "Greatest Percent Increase"
ws.Cells(3, 15).Value = "Greatest Percent Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
    
    'Dim variables
Dim tick_inc, tick_dec, tick_vol As String
Dim great_increase As Double
Dim least_increase As Double
Dim Greatest_vol As Double
Dim i As Double
Dim end_i As Double

    'Define loop and set variables to 0
great_increase = 0
least_increase = 0
Greatest_vol = 0
end_i = ws.Cells(Rows.Count, 1).End(xlUp).row
i = 2

For i = 2 To end_i
    'Find greatest increase
If ws.Cells(i, 11).Value > great_increase Then
    great_increase = ws.Cells(i, 11).Value
   tick_inc = ws.Cells(i, 9).Value
End If
   
   'Find greatest decrease
If ws.Cells(i, 11).Value < least_increase Then
    least_increase = ws.Cells(i, 11).Value
    tick_dec = ws.Cells(i, 9).Value
End If
    
    'Find Greatest vol
If ws.Cells(i, 12).Value > Greatest_vol Then
    Greatest_vol = ws.Cells(i, 12).Value
    tick_vol = ws.Cells(i, 9).Value
End If
Next i

'Show results
ws.Cells(2, 16).Value = tick_inc
ws.Cells(2, 17).Value = great_increase
    'use % format
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 16).Value = tick_dec
ws.Cells(3, 17).Value = least_increase
    'Use percent format
ws.Cells(3, 17).NumberFormat = "0.00%"
ws.Cells(4, 16).Value = tick_vol
ws.Cells(4, 17).Value = Greatest_vol

Next ws
End Sub

Sub Run_Analytics():

Stock_Analytics
Conditionals
Create_Columns

End Sub




