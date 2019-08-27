Attribute VB_Name = "Module1"
Sub Stocks()
Dim Ticker As String
Dim Vol_Total As Double

Vol_Total = 0

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

For i = 2 To 70926
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
Ticker = Cells(i, 1).Value
Vol_Total = Vol_Total + Cells(i, 7).Value
Range("I" & Summary_Table_Row).Value = Ticker
Range("J" & Summary_Table_Row).Value = Vol_Total


Summary_Table_Row = Summary_Table_Row + 1
Vol_Total = 0
Else
Vol_Total = Vol_Total + Cells(i, 7).Value
End If
Next i

End Sub

