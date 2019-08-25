Option Explicit

Sub stockVol()

'declare vars
Dim ticker As String
Dim openPrice As Single
Dim endClose As Single
Dim stockVol As Single
Dim yrChange As Single
Dim pctChange As Double
Dim greatestChg(3) As Single
Dim tickerCtr As Single



'set counters
Dim i As Long
Dim LastRow As Long
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
Dim jTable As Long
stockVol = 0
jTable = 2
tickerCtr = 1

'produce summary to jTable section
Range("J1").Value = "<TICKER SYM.>"
Range("K1").Value = "<Yrly. CHANGE>"
Range("L1").Value = "<Pct. CHANGE>"
Range("M1").Value = "<Ttl. VOL.>"
Range("J1:M1").Font.Bold = True
Range("J1:M1").Font.Underline = True
Range("J1:M1").Font.ColorIndex = 5
    
'iterate thru worksheet
For i = 2 To LastRow

    'catch open price
    If tickerCtr < 2 Then
    openPrice = Cells(i, 3).Value
    End If
        
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ticker = Cells(i, 1).Value
        stockVol = stockVol + Cells(i, 7).Value
        'set the ticker symbol
        Range("J" & jTable).Value = ticker
        
        'catch year end close price
        endClose = Cells(i, 6).Value
        yrChange = endClose - openPrice
        Range("K" & jTable).Value = Format(yrChange, "0.000")
        If Range("K" & jTable).Value >= 0 Then
            Range("K" & jTable).Interior.ColorIndex = 4
        ElseIf Range("K" & jTable).Value < 0 Then
            Range("K" & jTable).Interior.ColorIndex = 3
        End If
        
        'set percentage change
        If openPrice = 0# And yrChange = 0# Then
        pctChange = 0
        Else
        pctChange = yrChange / openPrice
        End If
        Range("L" & jTable).NumberFormat = "0.00%"
        Range("L" & jTable).Value = pctChange
        
        'summation of stock volume
        Range("M" & jTable).Value = stockVol
        jTable = jTable + 1
        
        'reset stockVol, tickerCtr for next ticker symbol
        stockVol = 0
        tickerCtr = 1
       
    'summation of stockVol per ticker symbol
    Else
        stockVol = stockVol + Cells(i, 7).Value
        tickerCtr = tickerCtr + 1
    
    End If
    
Next i
    

End Sub
