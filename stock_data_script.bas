Attribute VB_Name = "Module1"
Sub stock_volume()

For Each ws In Worksheets

Dim Ticker_name As String
Dim Ticker_total As Double
    Ticker_total = 0

Dim Summary_table_row As Integer
    Summary_table_row = 2
    
For I = 2 To 797711
    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
        Ticker_name = Cells(I, 1).Value
        Ticker_total = Ticker_total + Cells(I, 7).Value
        Range("H" & Summary_table_row).Value = Ticker_name
        Range("I" & Summary_table_row).Value = Ticker_total
        Summary_table_row = Summary_table_row + 1
        Ticker_total = 0
        
    Else
        Ticker_total = Ticker_total + Cells(I, 7).Value
        
        End If
    Next I
    
Next ws

End Sub


