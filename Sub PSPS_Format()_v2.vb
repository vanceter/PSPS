Sub PSPS_Format()
'
' PSPS_Format Macro
' Color columns for Gennie, vlookups
'updated to match additional column added for MGD ID
'

'
    Range("Y1").Select
    ActiveCell.FormulaR1C1 = "PGE LIST"
    Range("Z1").Select
    ActiveCell.FormulaR1C1 = "NG LIST"
    Range("AA1").Select
    ActiveCell.FormulaR1C1 = "BLOCK CODE"
    Range("AB1").Select
    ActiveCell.FormulaR1C1 = "3RD PARTY"
    Range("Y2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC1,PGEMASTERLIST!C[-24]:C[-23],1,FALSE)"
    'Range("Y2").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC1,PGEMASTERLIST!C[-22]:C[-21],1,FALSE)"
    Range("Z2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC7,NGMASTERLIST!C[-25]:C[-24],1,FALSE)"
    Range("AA2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC1,PGEOUTAGEBLOCKS!C[-26]:C[-25],2,FALSE)"
    Range("AB2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC7,3RDPARTY!C[-27]:C[-26],2,FALSE)"
    Range("Y2:AB2").Select
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
    Range("Y2:Y9140").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollRow = 9063
    ActiveWindow.ScrollRow = 8969
    ActiveWindow.ScrollRow = 8935
    ActiveWindow.ScrollRow = 5043
    ActiveWindow.ScrollRow = 2318
    ActiveWindow.ScrollRow = 2216
    ActiveWindow.ScrollRow = 2
    Range("L:M,R:T").Select
    Range("T1").Activate
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
 Columns("AA:AA").Select
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
    Columns("P:P").Select
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
