Sub stocks()

'Loop through all sheets
For Each ws In ActiveWorkbook.Worksheets

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Create headers for summary table
ws.Cells(1, 10).Value = "Ticker"
ws.Cells(1, 11).Value = "Total Stock Volume"

' Set an initial variable for holding the ticker
Dim ticker As String

' Set an initial variable for holding the volume per stock
Dim stock_volume As Double
stock_volume = 0

'Keep track of the location for each ticker in the summary table
Dim summary_table_row As Long
summary_table_row = 2

'Loop through all stock volumes
For i = 2 To LastRow

'Check if we are still within the same ticker, if it is not...
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    'Set the Ticker
    ticker = ws.Cells(i, 1).Value
    
    'Add to the Stock Volume
    stock_volume = stock_volume + ws.Cells(i, 7).Value
    
    'Print the Ticker in the Summary Table
    ws.Range("J" & summary_table_row).Value = ticker
    
    'Print the Stock Volume in the Summary Table
    ws.Range("K" & summary_table_row).Value = stock_volume
    
    'Add one to the Summary Table Row
    summary_table_row = summary_table_row + 1
    
    'Reset the Stock Volume
    stock_volume = 0
    
    'If the cell immediately following a row is the same ticker...
    Else
    
    'Add to the Stock Volume
    stock_volume = stock_volume + ws.Cells(i, 7).Value
    
    End If
    
  Next i
  
    For i = 2 To LastRow
        ws.Cells(i, 11).NumberFormat = "#,##0"
        
    Next i
    
  Next ws

End Sub




