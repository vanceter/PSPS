Sub PSPS_Format()
'
' PSPS_Format Macro
' Color columns for Gennie, vlookups
'

'
    Range("W1").Select
    ActiveCell.FormulaR1C1 = "PGE LIST"
    Range("X1").Select
    ActiveCell.FormulaR1C1 = "NG LIST"
    Range("Y1").Select
    ActiveCell.FormulaR1C1 = "BLOCK CODE"
    Range("Z1").Select
    ActiveCell.FormulaR1C1 = "3RD PARTY"
    Range("W2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC1,PGEMASTERLIST!C[-22]:C[-21],1,FALSE)"
    Range("W2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC1,PGEMASTERLIST!C[-22]:C[-21],1,FALSE)"
    Range("X2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC5,NGMASTERLIST!C[-23]:C[-22],1,FALSE)"
    Range("Y2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC1,PGEOUTAGEBLOCKS!C[-24]:C[-23],2,FALSE)"
    Range("Z2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC5,3RDPARTY!C[-25]:C[-24],2,FALSE)"
    Range("W2:Z2").Select
    Selection.Copy
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 61
    ActiveWindow.ScrollRow = 1236
    ActiveWindow.ScrollRow = 1432
    ActiveWindow.ScrollRow = 3825
    ActiveWindow.ScrollRow = 5231
    ActiveWindow.ScrollRow = 6014
    ActiveWindow.ScrollRow = 7087
    ActiveWindow.ScrollRow = 7206
    ActiveWindow.ScrollRow = 8637
    ActiveWindow.ScrollRow = 8833
    ActiveWindow.ScrollRow = 8842
    ActiveWindow.ScrollRow = 9072
    Range("W2:W9140").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollRow = 9063
    ActiveWindow.ScrollRow = 8969
    ActiveWindow.ScrollRow = 8935
    ActiveWindow.ScrollRow = 5043
    ActiveWindow.ScrollRow = 2318
    ActiveWindow.ScrollRow = 2216
    ActiveWindow.ScrollRow = 2
    Range("K:M,Q:S").Select
    Range("Q1").Activate
    Selection.FormatConditions.Add Type:=xlTextString, String:="NO", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:="YES", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Cells.Select
    Application.CutCopyMode = False
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
 Columns("Y:Y").Select
Selection.FormatConditions.Add Type:=xlTextString, String:="50", _
    TextOperator:=xlContains
Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
With Selection.FormatConditions(1).Font
    .Color = -16752384
    .TintAndShade = 0
End With
With Selection.FormatConditions(1).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13561798
    .TintAndShade = 0
End With
    Columns("O:O").Select
    Selection.FormatConditions.Add Type:=xlTextString, String:="Propane", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = -0.249946592608417
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399945066682943
    Selection.FormatConditions(1).StopIfTrue = False
    End With
End Sub
