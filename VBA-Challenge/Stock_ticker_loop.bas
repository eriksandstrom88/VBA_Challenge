Attribute VB_Name = "Module1"
Sub stock_loop():
    'name columns
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Stock Volume"
    
    'declare variables
        Dim Ticker As String
        Dim Open_Rate As Double
        Dim Close_Rate As Double
        Dim Volume As Long
        Dim row_index As Integer
        Dim column_index As Integer
        Dim i As Long
        Dim last_row As Long
        
        i_summary = 2
        
        last_row = Cells(Rows.Count, 1).End(xlUp).Row
        
    ' MsgBox (last_row)
        For i = 2 To last_row:
            Ticker = Cells(i, 1).Value
            If Ticker <> Cells(i + 1, 1).Value Then
                Cells(i_summary, 9) = Ticker
                i_summary = i_summary + 1
        End If
    Next i
    
End Sub

