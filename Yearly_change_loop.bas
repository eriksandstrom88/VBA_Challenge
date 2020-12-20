Attribute VB_Name = "Module3"


'THIS IS A WORK IN PROGRESS.
'I RAN OUT OF TIME TO COMPLETE THIS ASSIGNMENT DUE TO YEAR-END
'OBLIGATIONS FOR WORK.

'Conceptually, I know I need to develop a script that counts the number
'of rows for each ticker, then subtract the first "Open" from the last
'"Close" to get the yearly change.  Then, to get the percent change,
'I need to write a script that executes:
'(yearly change-first open)/first open * 100

'Finally, I need a script that sums the volume column for each ticker


'output yearly change in price from opeing to close (column H)

Sub yearly_change():
    Dim column_index As Integer
    Dim row_index As Long
    Dim ticker_counter As Integer
    Dim ticker_index As Integer
    
    
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    last_unique_ticker_row = Cells(Rows.Count, 9).End(xlUp).Row
    ticker_counter = 0
    For ticker_index = 2 To 5 'last_unique_ticker_row
        For row_index = 2 To 5 'last_row
            If Cells(ticker_index, 9) = Cells(row_index, 1) Then
                ticker_counter = ticker_counter + 1
                Cells(ticker_index, 10).Value = ticker_counter
            End If
        ticker_counter = 0
        Next row_index
        
        
    Next ticker_index
    
        
    
End Sub
    'output percent change from opening to closing
    'output total stock volume of stock

