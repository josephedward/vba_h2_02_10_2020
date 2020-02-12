
Sub Stock_Market()

Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning
For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    'do whatever you need
    
Application.DisplayAlerts = False

' remove duplicate ticker strings
Range("A:A").Copy Range("J:J")
Dim MyRange As Range
Dim LastRow As Long
LastRow = Cells(Rows.count, 1).End(xlUp).Row
Set MyRange = ActiveSheet.Range("J1:J" & LastRow)
MyRange.RemoveDuplicates Columns:=1, Header:=xlYes
    
    
'set headers
Range("J1") = "<ticker>"
Range("K1") = "<price difference>"
Range("L1") = "<percentage change>"
Range("M1") = "<total volume>"
    
Dim count As Long
Dim currentStock As String
Dim openPrice As Double
Dim closePrice As Double
Dim priceDiff As Double
Dim percentChange As Double
Dim totalVolume As Variant

count = 0

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
    Cells(count + 1, 13) = totalVolume
Next i

'this sets cell A1 of each sheet to "1"
Next
starting_ws.Activate 'activate the worksheet that was originally active

End Sub
