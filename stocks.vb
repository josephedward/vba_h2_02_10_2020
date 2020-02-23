Sub Stock_Market()

'loop through worksheets
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
    ws.Activate

    'testing
    'Application.DisplayAlerts = False

    'grab last row
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
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
            'special condition for first iteration of loop
            If i = 2 Then
                openPrice = Cells(i, 3)
            End If
            'if you dont have a string
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
            If priceDiff > 0 Then
                Cells(printCount, 11).Interior.Color = vbGreen
            ElseIf priceDiff < 0 Then
                Cells(printCount, 11).Interior.Color = vbRed
            End If
            
            Cells(printCount, 12) = percentChange & "%"
            'zero out total volume because its a new stock
            totalVolume = 0
        End If
        'add to total volume
        totalVolume = totalVolume + Cells(i, 7)
        'print value
        Cells(printCount + 1, 13) = totalVolume
        '** ADD HANDLING FOR LAST ROW **
        
        
    Next i
    
    'set headers
    Range("J1") = "<ticker>"
    Range("K1") = "<price difference>"
    Range("L1") = "<percentage change>"
    Range("M1") = "<total volume>"
    
    
    Range("O2") = "<greatest increase>"
    Range("O3") = "<greatest decrease>"
    Range("O4") = "<greatest total volume>"
    
    Dim finalRow_totals As Long
    finalRow_totals = Range("L800000").End(xlUp).Row
    Dim rowString As String
    rowString = "L" + CStr(finalRow_totals)
    
    Dim dblMin As Double
    Dim dblMax As Double
    Dim sMin As String
    Dim sMax As String
    Dim varMax As Variant
    dblMin = Application.WorksheetFunction.Min(Range("L2", rowString))
    sMin = dblMin * 100
    dblMax = Application.WorksheetFunction.Max(Range("L2", rowString))
    sMax = dblMax * 100
    varMax = Application.WorksheetFunction.Max(Range("M2", rowString))
        
    Range("P1") = "<Ticker>"
    Range("Q1") = "<Value>"
    
    Range("Q2") = sMax
    Range("Q3") = sMin
    Range("Q4") = varMax
    
    
    ' 'vars for finding ticker
    ' Dim rng As Range
    ' Dim cell As Range
    ' Dim search As Double
    ' Set rng = Range("L:L")
    ' Dim ticker As String
    
    ' 'find maximum on sheet for ticker
    ' search = sMax
    ' 'Set cell = rng(CLng(search))
    ' Set cell = rng.Find(What:=CLng(search), LookIn:=xlValues, MatchCase:=False, After:=ActiveCell)
    ' ticker = Cells(cell.Row, cell.Column - 2)
    ' 'prints error if not found
    ' 'Debug.Print cell.Address
    ' Range("P2") = ticker
    
    ' 'find minimum on sheet for ticker
    ' search = sMin
    ' Set cell = rng.Find(What:=CLng(search), LookIn:=xlValues, MatchCase:=False, After:=ActiveCell)
    ' ticker = Cells(cell.Row, cell.Column - 2)
    ' 'prints error if not found
    ' 'Debug.Print cell.Address
    ' Range("P3") = ticker
    
'   **OVERFLOW**
'    search = varMax
'    Set cell = rng.Find(What:=CLng(search), LookIn:=xlValues, MatchCase:=False, After:=ActiveCell)
'    ticker = CStr(Cells(cell.Row, cell.Column - 2))
'    Range("P4") = ticker
    
    
Next 'next sheet
End Sub



