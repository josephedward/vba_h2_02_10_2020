Sub GetTickers()
Range("A:A").Copy Range("J:J")
Dim MyRange As Range

Dim LastRow As Long
LastRow = ActiveSheet.Range("A" & Rows.count).End(xlUp).Row

Set MyRange = ActiveSheet.Range("J1:J" & LastRow)
MyRange.RemoveDuplicates Columns:=1, Header:=xlYes

Dim LastTickerRow As Long
LastTickerRow = ActiveSheet.Range("J" & Rows.count).End(xlUp).Row
MsgBox (LastTickerRow)
'Dim Stocks(LastTickerRow)

Dim Cell As Range
For Each Cell In MyRange
    Dim Stock As Object
    Set Stock = CreateStockObjs(Cell)
Next Cell

End Sub




