Attribute VB_Name = "Module1"
Sub getStock()
    Dim ws As Worksheet
    Dim lngRow As Long, lngCount As Long, lngCount2 As Long
    Dim dbDiv1 As Double, dbDiv2 As Double
    Dim rngFound As Range
    Dim StartTime As Double, SecondsElapsed As Double

    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    StartTime = Timer
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        lngRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        lngCount = 2

        'Ticker
        ws.Range("I1").Value = "Ticker"
        ws.Range("I2:I" & lngRow).Value = ws.Range("A2:A" & lngRow).Value
        ws.Range("I2:I" & lngRow).RemoveDuplicates Columns:=(1)

        'Rest
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        For i = 2 To ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
            lngCount2 = Application.WorksheetFunction.CountIf(Range("A2:A" & lngRow), ws.Range("I" & i).Value)
            ws.Range("J" & i).Value = ws.Range("F" & lngCount2 + lngCount - 1).Value - ws.Range("C" & lngCount).Value
            dbDiv1 = ws.Range("J" & i).Value
            dbDiv2 = ws.Range("C" & lngCount).Value
            If dbDiv2 = 0 Then
                ws.Range("K" & i).Value = 0
            Else
                ws.Range("K" & i).Value = dbDiv1 / dbDiv2
            End If
            ws.Range("L" & i).Formula = "=sumifs(G2:G" & lngRow & ",A2:A" & lngRow & "," & Chr(34) & ws.Range("I" & i).Value & Chr(34) & ")"
            ws.Range("L" & i).Value = ws.Range("L" & i)
            lngCount = lngCount2 + lngCount
        Next i
        
        ws.Range("J2:J" & i - 1).Select
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 255
            .TintAndShade = 0
        End With
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
        With Selection.FormatConditions(2).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 5296274
            .TintAndShade = 0
        End With
        
        ws.Range("O2").Value = "Greatest % increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest total volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ws.Range("Q2").Formula = "=MAX(K2:K" & i - 1 & ")"
        ws.Range("Q3").Formula = "=MIN(K2:K" & i - 1 & ")"
        ws.Range("Q4").Formula = "=MAX(L2:L" & i - 1 & ")"
        ws.Columns("A:Q").AutoFit
        
        Set rngFound = Nothing
        Set rngFound = ws.Range("K2:K" & i - 1).Find(what:=ws.Range("Q2").Value, LookIn:=xlFormulas)
        ws.Range("P2").Value = ws.Range("I" & rngFound.Row).Value
        Set rngFound = Nothing
        Set rngFound = ws.Range("K2:K" & i - 1).Find(what:=ws.Range("Q3").Value)
        ws.Range("P3").Value = ws.Range("I" & rngFound.Row).Value
        Set rngFound = Nothing
        Set rngFound = ws.Range("L2:L" & i - 1).Find(what:=ws.Range("Q4").Value, LookIn:=xlValues)
        ws.Range("P4").Value = ws.Range("I" & rngFound.Row).Value
        
        ws.Range("K2:K" & i - 1).NumberFormat = "0.00%"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
    Next ws

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    SecondsElapsed = Round(Timer - StartTime, 2)
    MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation
End Sub


