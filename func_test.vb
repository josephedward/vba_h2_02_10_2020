Sub FuncTest()
        
Dim LastRow As Long
LastRow = ActiveSheet.Range("A" & Rows.count).End(xlUp).Row

Dim count As Long
Dim currentStock As String
Dim openPrice As Double
Dim closePrice As Double
Dim priceDiff As Double
Dim percentChange As Double
Dim totalVolume As Variant

Dim i As Long
For i = 2 To CLng(LastRow)

    If Cells(i, 1) <> currentStock Then
    totalVolume = 0
'    Debug.Print ("new")
    count = count + 1
    currentStock = Cells(i, 1)
    openPrice = Cells(i, 3)
    'Debug.Print (openPrice)
    closePrice = Cells(i + 261, 6)
    'Debug.Print (closePrice)
    priceDiff = closePrice - openPrice
    percentChange = (priceDiff / openPrice) * 100
    Cells(count + 1, 11) = priceDiff
    Cells(count + 1, 12) = percentChange & "%"
'    totalVolume = totalVolume + Cells(i, 7)
    
    End If
    
    totalVolume = totalVolume + Cells(i, 7)
    'Debug.Print (totalVolume)
    Cells(CLng(count + 1), 13) = totalVolume
    
Next i

'MsgBox (count)

End Sub
