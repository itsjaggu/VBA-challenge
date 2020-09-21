Option Explicit

Sub StockReport()

    Dim ws As Worksheet
    Dim StartRow As Long
    Dim EndRow As Long
    Dim DataEndRow As Long
    
    'Set ws = Sheet7
    For Each ws In ActiveWorkbook.Worksheets
    
        'Getting end row of the data in the sheet
        DataEndRow = ws.Range("A1").End(xlDown).Row
        
        ws.Range("A1:A" & DataEndRow).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ws.Range("I1"), Unique:=True
        
        'Getting row numbers for unique list of tickers copied in column I
        StartRow = 2
        EndRow = ws.Range("I1").End(xlDown).Row
        
        'Getting first date of the year for the ticker
        ws.Range("J1").Value = "Open Date"
        ws.Range("J2:J" & EndRow).FormulaR1C1 = "=MINIFS(R2C2:R" & DataEndRow & "C2,R2C1:R" & DataEndRow & "C1,RC9)"
        'Getting last date of the year for the ticker
        ws.Range("K1").Value = "Close Date"
        ws.Range("K2:K" & EndRow).FormulaR1C1 = "=MAXIFS(R2C2:R" & DataEndRow & "C2,R2C1:R" & DataEndRow & "C1,RC9)"
        'Getting Open price on first date of the year for the ticker
        ws.Range("L1").Value = "First Open Price"
        ws.Range("L2:L" & EndRow).FormulaR1C1 = "=SUMPRODUCT((R2C1:R" & DataEndRow & "C1=RC9)*(R2C2:R" & DataEndRow & "C2=RC10),R2C3:R" & DataEndRow & "C3)"
        'Getting Close price on last date of the year for the ticker
        ws.Range("M1").Value = "Last Close Price"
        ws.Range("M2:M" & EndRow).FormulaR1C1 = "=SUMPRODUCT((R2C1:R" & DataEndRow & "C1=RC9)*(R2C2:R" & DataEndRow & "C2=RC11),R2C6:R" & DataEndRow & "C6)"
        'Calculating Yearly change by subtracking Last Close Price - First Open Price
        ws.Range("N1").Value = "Yearly Change"
        ws.Range("N2:N" & EndRow).FormulaR1C1 = "=RC[-1]-RC[-2]"
        'Calculating % change in price based on Open price
        ws.Range("O1").Value = "Percent Change"
        ws.Range("O2:O" & EndRow).FormulaR1C1 = "=IF(RC[-3]=0,0,RC[-1]/RC[-3])"
        'ws.Range("O2:O" & EndRow).Style = "Percent"
        ws.Range("O2:O" & EndRow).NumberFormat = "0.00%"
        'Calculating Sum of volume of the whole year
        ws.Range("P1").Value = "Total Stock Volume"
        ws.Range("P2:P" & EndRow).FormulaR1C1 = "=SUMIF(R2C1:R" & DataEndRow & "C1,RC9,R2C7:R" & DataEndRow & "C7)"
        'Formatting Yearly Change for positive and negative values
        With ws.Range("N2:N" & EndRow)
            'Format for Negative value
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
            .FormatConditions(.FormatConditions.Count).SetFirstPriority
            With .FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 255
                .TintAndShade = 0
            End With
            .FormatConditions(1).StopIfTrue = False
            'Format for Positive value
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
            .FormatConditions(.FormatConditions.Count).SetFirstPriority
            With .FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 5296274
                .TintAndShade = 0
            End With
            .FormatConditions(1).StopIfTrue = False
        End With
        
        '''Challenge'''
        Dim i As Long
        Dim CurPerVal As Double
        Dim CurVolVal As Double
        Dim CurTickerVal As String
        Dim IncTickerVal As String
        Dim GreatestIncrease As Double
        Dim DecTickerVal As String
        Dim GreatestDecrease As Double
        Dim VolTickerVal As String
        Dim GreatestTotalVolume As Double
        
        i = 0
        CurPerVal = 0
        CurVolVal = 0
        CurTickerVal = ""
        IncTickerVal = ""
        GreatestIncrease = 0
        DecTickerVal = ""
        GreatestDecrease = 0
        VolTickerVal = ""
        GreatestTotalVolume = 0
        
        ws.Range("R2").Value = "Greatest % Increase"
        ws.Range("R3").Value = "Greatest % Decrease"
        ws.Range("R4").Value = "Greatest Total Volume"
        ws.Range("S1").Value = "Ticker"
        ws.Range("T1").Value = "Value"
        For i = StartRow To EndRow
            CurPerVal = ws.Range("O" & i).Value
            CurVolVal = ws.Range("P" & i).Value
            CurTickerVal = ws.Range("I" & i).Value
            If GreatestIncrease = Null Then
                GreatestIncrease = CurPerVal
                IncTickerVal = CurTickerVal
                GreatestDecrease = CurPerVal
                DecTickerVal = CurTickerVal
            ElseIf CurPerVal > GreatestIncrease Then
                GreatestIncrease = CurPerVal
                IncTickerVal = CurTickerVal
            ElseIf CurPerVal < GreatestDecrease Then
                GreatestDecrease = CurPerVal
                DecTickerVal = CurTickerVal
            End If
            If GreatestTotalVolume = Null Then
                GreatestTotalVolume = CurVolVal
                VolTickerVal = CurTickerVal
            ElseIf CurVolVal > GreatestTotalVolume Then
                GreatestTotalVolume = CurVolVal
                VolTickerVal = CurTickerVal
            End If
        Next i
            
        ws.Range("S2").Value = IncTickerVal
        ws.Range("T2").Value = GreatestIncrease
        ws.Range("T2").NumberFormat = "0.00%"
        ws.Range("S3").Value = DecTickerVal
        ws.Range("T3").Value = GreatestDecrease
        ws.Range("T3").NumberFormat = "0.00%"
        ws.Range("S4").Value = VolTickerVal
        ws.Range("T4").Value = GreatestTotalVolume
        
        'Autofit columns to show complete headers
        ws.Columns("I:T").EntireColumn.AutoFit
    Next
        
End Sub




