Sub HomeWork2()
' Create script that will loop through all the stocks for one year and output Ticker symbol,
'   Yearly change from opening price to closing price, % change from opening to closing price (with
'   conditional formatting for + and - changes),Total stock volume

    ' LOOP THROUGH ALL SHEETS
    For Each ws In Worksheets
        ' GET THE WORKSHEET NAME SO LATER CAN AUTOFIT COLUMNS
        Dim WorkSheetName As String
        WorkSheetName = ws.Name
                
        ' DETERMINE LAST ROW
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'INSERT NEW COLUMN HEADERS
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "% Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1,16).Value = "Ticker"
        ws.Cells(1,17).Value = "Value"
        ws.Cells(2,15).Value = "Greatest % Increase"
        ws.Cells(3,15).Value = "Greatest % Decrease"
        ws.Cells(4,15).Value = "Greatest Total Volume"
        
        ' FORMAT THE NEW COLUMNS
        Columns(9).NumberFormat = "@"
        Columns(10).NumberFormat = "0.00"
        Columns(11).NumberFormat = "0.00%"
        Columns(12).NumberFormat = "0"
        Columns(15).NumberFormat = "@"
        Columns(16).NumberFormat = "@"
        
        ' DECLARE CURRENT TICKER VARIABLE
        Dim CurrentTicker As String
                        
        ' DECLARE TICKER OPENING AND CLOSING VALUE, TOTAL VOLUME, YEARLY CHANGE and TICKER ROW VARIABLES
        Dim TickOpen As Double
        Dim TickClose As Double
        Dim TotalVolume As Double
        Dim TickerRow As Integer
        Dim YearlyChange As Double
        Dim PercentChange As Double
        
        ' INITIALIZE VARIABLES
        TickOpen = ws.Cells(2, 3).Value
        TickClose = 0#
        TotalVolume = 0
        TickerRow = 2
        CurrentTicker = ws.Cells(2, 1).Value
        PercentChange = 0#
        
        ' FOR EACH ROW FROM 2 THROUGH LastRow
        For i = 2 To LastRow
        
            ' PROCESS THE DATA IN THE ROW
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
    
            ' IF AT END OF TICKER (CurrentTicker NOT EQUAL NEXT ROW TICKER)
            If i = LastRow Or ws.Cells(i + 1, 1) <> CurrentTicker Then
                ' PUT CurrentTicker in CELL(TickerRow,9)
                ws.Cells(TickerRow, 9).Value = CurrentTicker
                ' CALCULATE YEARLY CHANGE (TickClose - TickOpen) AND PUT IN CELL(TICKER ROW,10)
                TickClose = ws.Cells(i, 6).Value
                YearlyChange = TickClose - TickOpen
                ws.Cells(TickerRow, 10).Value = YearlyChange
                If (YearlyChange < 0) Then
                    ' IF NEGATIVE CHANGE MAKE CELL BACKGROUND RED
                    ws.Cells(TickerRow, 10).Interior.ColorIndex = 3
                    ' ELSE MAKE GREEN (Could make zero no-color but decided not to test for that)
                    Else: ws.Cells(TickerRow, 10).Interior.ColorIndex = 4
                End If
                ' CALCULATE % CHANGE (Yearly Change / TickOpen AND PUT IN CELL(TICKER ROW,11)
                ' MAKE SURE TickOpen <> 0
                If TickOpen <> 0 Then
                    PercentChange = YearlyChange / TickOpen
                    ws.Cells(TickerRow, 11).Value = PercentChange
                End If
                ' PUT TotalVolume IN CELL (TickerRow, 12)
                ws.Cells(TickerRow, 12).Value = TotalVolume
                
                ' SET UP FOR NEXT ROW -- INCREMENT THE TickerRow AND SET NEXT TICKER IF NOT LAST ROW
                TickerRow = TickerRow + 1
                ' SET TotalVolume TO ZERO
                TotalVolume = 0
                'IF NOT LAST ROW, SET CurrentTicker TO NEXT ROW TICKER VALUE and SET TickOpen TO NEXT ROW OPEN VALUE
                If i <> LastRow Then
                    CurrentTicker = ws.Cells(i + 1, 1).Value
                    TickOpen = ws.Cells(i + 1, 3).Value
                End If
                
            ' ENDIF
            End If
            
        ' NEXT ROW
        Next i
        
        ' DECLARE AND INITIALIZE VARIABLES FOR GREATEST INCREASE, GREATEST DECREASE and GREATEST TOTAL VOLUME
        Dim maxIncrease As Double
        maxIncrease = 0.00
        Dim maxDecrease As Double
        maxDecrease = 0.00
        Dim maxVolume As Double
        maxVolume = 0
        Dim rowSoFar as Double

        ' DETERMINE GREATEST % INCREASE AND PUT TICKER AND VALUE IN ROW 2, COLUMNS P AND Q RESPECTIVELY
        ws.Cells(2,15) = "Greatest % Increase"
        For j = 2 to LastRow
            if ws.Cells(j,11) > maxIncrease Then
                maxIncrease = ws.Cells(j,11)
                rowSoFar = j
            End If
        next j
        ws.Cells(2,17) = maxIncrease
        ws.Cells(2,17).NumberFormat = "0.00%"
        ws.Cells(2,16) = ws.Cells(rowSoFar,9)
        ' DETERMINE GREATEST % DECREASE AND PUT TICKER AND VALUE IN ROW 3, COLUMNS P AND Q RESPECTIVELY
        ws.Cells(3,15) = "Greatest % Decrease"
        For j = 2 to LastRow
            if ws.Cells(j,11) < maxDecrease Then
                maxDecrease = ws.Cells(j,11)
                rowSoFar = j
            End If
        next j
        ws.Cells(3,17) = maxDecrease
        ws.Cells(3,17).NumberFormat = "0.00%"
        ws.Cells(3,16) = ws.Cells(rowSoFar,9)
        ' DETERMINE GREATEST TOTAL VOLUME AND PUT TICKER AND VALUE IN ROW 4, COLUMNS P AND Q RESPECTIVELY
        ws.Cells(4,15) = "Greatest Total Volume"
        For j = 2 to LastRow
            if ws.Cells(j,12) > maxVolume Then
                maxVolume = ws.Cells(j,12)
                rowSoFar = j
            End If
        next j
        ws.Cells(4,17) = maxVolume
       '  ws.Cells(4,17).NumberFormat = "0.00"
        ws.Cells(4,16) = ws.Cells(rowSoFar,9)
        'AUTOFIT THE NEW COLUMNS (I through L) IN THIS WORKSHEET
        Columns("I:Q").EntireColumn.AutoFit
    ' GO TO NEXT WORKSHEET
    Next ws
End Sub
       