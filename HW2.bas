Attribute VB_Name = "Module1"
Sub VBHW2_StockExercise():

For Each ws In Worksheets

'Set summary table headers and auto fit column width
    ws.Range("J1").Value = "Ticker"
        ws.Columns("J:J").AutoFit
    ws.Range("K1").Value = "Yearly Change"
        ws.Columns("K:K").AutoFit
    ws.Range("L1").Value = "Precent Change"
        ws.Columns("L:L").AutoFit
    ws.Range("M1").Value = "Total Volume Change"
        ws.Columns("M:M").AutoFit
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
        ws.Columns("O:O").AutoFit
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"

    'Declare Variables for the initial summary
    Dim TickerRowCounter As Double, TradeVolumeTotal As Double
    Dim OpenPrice As Double, ClosePrice As Double
    Dim TickerTotal As Long
    'Initialize the variables. Stock Row is 2 as that is the first ticker data
    TickerRowCounter = 2
    TradeVolumeTotal = 0
    OpenPrice = 0
    ClosePrice = 0
    'Find the Last Row
    TickerTotal = ws.Cells(Rows.Count, 1).End(xlUp).Row
    OpenPrice = ws.Cells(2, 3).Value
    'Determine where the Ticker changes symbol and calc price change and total volume traded. Also identify any null values
    Dim i As Long
    For i = 2 To TickerTotal
        
        'Begin to tally up the total stock volume
        TradeVolumeTotal = TradeVolumeTotal + ws.Cells(i, 7).Value
        
        
        'Identify where the stock ticker changes
         If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'Record ticker
            ws.Cells(TickerRowCounter, 10).Value = ws.Cells(i, 1).Value
            'Record the TradeVolumeTotal
            ws.Cells(TickerRowCounter, 13).Value = TradeVolumeTotal
            'calculate and record price change
            ClosePrice = ws.Cells(i, 6)
            ws.Cells(TickerRowCounter, 11).Value = ClosePrice - OpenPrice
            'Remove null values from calculation
            If OpenPrice = 0 Then
                ws.Cells(TickerRowCounter, 12).Value = 0
                Else
            'Percent Change
            ws.Cells(TickerRowCounter, 12).Value = FormatPercent((ClosePrice - OpenPrice) / OpenPrice)
            End If
            'Set Open Price for next ticker
            OpenPrice = ws.Cells(i + 1, 3).Value
            'Set row counter so new Ticker is added below previous stock
            TickerRowCounter = TickerRowCounter + 1
            'Set TradeVolumeTotal and OpenPrice and ClosePrice back to 0 for next Stock
            'Reset Trade volume and close price
            TradeVolumeTotal = 0
            ClosePrice = 0
        End If
    Next i
    
'Set formating to show positive/negative percent change
    Dim ChangeTickerTotal As Long
    ChangeTickerTotal = ws.Cells(Rows.Count, 11).End(xlUp).Row

    Dim j As Integer
    For j = 2 To ChangeTickerTotal
            If ws.Cells(j, 11).Value >= 0 Then
                ws.Cells(j, 11).Interior.ColorIndex = 4
            ElseIf ws.Cells(j, 11).Value < 0 Then
                ws.Cells(j, 11).Interior.ColorIndex = 3
          End If
    Next j
    
    'Find Max and Min of the columns
    Dim PercentRange As Range, TradeRange As Range
    Dim MaxPercent As Double, MinPercent As Double, MaxVol As Double
        
    'Calc Max of %range
    Set PercentRange = ws.Range("L:L")
    MaxPercent = Application.WorksheetFunction.Max(PercentRange)
    ws.Range("Q2").Value = FormatPercent(MaxPercent)
    'Calc Min of %range
    MinPercent = Application.WorksheetFunction.Min(PercentRange)
    ws.Range("Q3").Value = FormatPercent(MinPercent)
    'Calc max trade volume
    Set TradeRange = ws.Range("M:M")
    MaxVol = Application.WorksheetFunction.Max(TradeRange)
    ws.Range("Q4").Value = MaxVol
        
    'Find the ticker values
    Dim t As Integer
        For t = 2 To ChangeTickerTotal
            If ws.Cells(t, 12).Value = ws.Range("Q2").Value Then
                ws.Range("P2").Value = ws.Cells(t, 10).Value
            ElseIf ws.Cells(t, 12).Value = ws.Range("Q3").Value Then
                ws.Range("P3").Value = ws.Cells(t, 10).Value
            ElseIf ws.Cells(t, 13).Value = ws.Range("Q4").Value Then
                ws.Range("P4").Value = ws.Cells(t, 10).Value
            End If
        Next t
    ws.Columns("Q:Q").AutoFit
    ws.Columns("P:P").AutoFit
Next ws
MsgBox ("Summary Tables created for all Stocks and Year Worksheets")
End Sub



