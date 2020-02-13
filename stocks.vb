Sub Stock_Market()

'loop through worksheets
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
    ws.Activate

    'testing
    'Application.DisplayAlerts = False

    'grab last row
    Dim LastRow As Long
    LastRow = Cells(Rows.count, 1).End(xlUp).Row
    
    'allocate memory for variables
    Dim currentStock As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim priceDiff As Double
    Dim percentChange As Double
    Dim totalVolume As Variant

    'for tracking where to print results
    Dim printCount As Long

    'initialize
    printCount = 0
    openPrice = 0
    closePrice = 0
    priceDiff = 0
    percentChange = 0

    'loop through rows
    Dim i As Long
    For i = 2 To CLng(LastRow)
        ' conditional for new stock
        If Cells(i, 1) <> currentStock Then
            currentStock = Cells(i, 1)
            If i = 2 Then
                openPrice = Cells(i, 3)
            End If
            If VarType(Cells(i - 1, 6)) <> 8 Then
                closePrice = Cells(i - 1, 6)
                priceDiff = closePrice - openPrice
                openPrice = Cells(i, 3)
                'calculate as percentage if not dividing by zero
                If openPrice <> 0 Then
                    percentChange = (priceDiff / openPrice) * 100
                End If
            End If
            'iterate printCounter
            printCount = printCount + 1
            'print name, difference and percentage
            Cells(printCount + 1, 10) = currentStock
            Cells(printCount, 11) = priceDiff
            Cells(printCount, 12) = percentChange & "%"
            'zero out total volume because its a new stock
            totalVolume = 0
        End If
        'add to total volume
        totalVolume = totalVolume + Cells(i, 7)
        'print value
        Cells(printCount + 1, 13) = totalVolume
    Next i
    
    'set headers
    Range("J1") = "<ticker>"
    Range("K1") = "<price difference>"
    Range("L1") = "<percentage change>"
    Range("M1") = "<total volume>"
        
Next
End Sub
