Attribute VB_Name = "Module1"
Sub Not_Easy()

' Set an initial variable for holding the ticker name
Dim ticker_name As String

' Set an initial variable for holding the total ticker count
Dim ticker_total As Double
ticker_name = 0

' Keep track of the location for each unique ticker name in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

' Loop through all tickers
For i = 2 To 797711

' Check if we are still within the same ticker name, if it is not...
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

' Set the Ticker name
ticker_name = Cells(i, 1).Value

' Add to the Ticker Total
ticker_total = ticker_total + Cells(i, 7).Value

' Print the Ticket Total in the Summary Table

Range("J" & Summary_Table_Row).Value = ticker_name

' Print the Ticker Amount to the Summary Table
Range("K" & Summary_Table_Row).Value = ticker_total

' Add one to the summary table row
Summary_Table_Row = Summary_Table_Row + 1

' Reset the Ticker Total
ticker_total = 0

' If the cell immediately following a row is the same ticker...
Else

' Add to the Ticker Total
ticker_total = ticker_total + Cells(i, 7).Value


End If

Next i

End Sub
