Attribute VB_Name = "Module1"


Sub solved():

Dim rowcount As Long
rowcount = 0 'total rows per sheet

    rowcount = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
    'MsgBox rowcount

Dim i As Long
Dim j As Long
Dim sumstockvol As LongLong
Dim boyopen As Double
Dim eoyclose As Double
Dim outputrow As Integer
Dim outputcol As Integer

outputrow = 2
sumstockvol = 0

boyopen = Cells(2, 3).Value


For i = 2 To rowcount


If Cells(i, 1).Value = Cells(i + 1, 1).Value Then

    sumstockvol = sumstockvol + CLng(Cells(i, 7).Value)

Else

    sumstockvol = sumstockvol + CLng(Cells(i, 7).Value)

    eoyclose = Cells(i, 6)

    Cells(outputrow, 8).Value = Cells(i, 1).Value
    Cells(outputrow, 9).Value = eoyclose - boyopen

    'color yearly change
    If Sgn(Cells(outputrow, 9).Value) = -1 Then
        ActiveSheet.Cells(outputrow, 9).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 192
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With

    ElseIf Sgn(Cells(outputrow, 9).Value) = 1 Then
        ActiveSheet.Cells(outputrow, 9).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = -0.249946592608417
            .PatternTintAndShade = 0
        End With
    
    End If


    If boyopen = 0 Then
        Cells(outputrow, 10).Value = 0
    Else

        Cells(outputrow, 10).Value = (eoyclose - boyopen) / boyopen
    End If

    Cells(outputrow, 11).Value = sumstockvol

    outputrow = outputrow + 1 'moves to next output row
    'reset back to original
    sumstockvol = 0
    eoyclose = 0
    boyopen = Cells(i + 1, 3).Value

End If



Next i


End Sub


