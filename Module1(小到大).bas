Attribute VB_Name = "Module1"
Sub ()
Attribute .VB_Description = "皊弘计秖パ"
Attribute .VB_ProcData.VB_Invoke_Func = "a\n14"
'
'  エ栋
' 皊弘计秖パ
'
' е硉龄: Ctrl+a
'
    Columns("B:B").Select
    ActiveWorkbook.Worksheets("1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("1").Sort.SortFields.Add Key:=Range("B2:B553"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("1").Sort
        .SetRange Range("A1:B553")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("1").Sort.SortFields.Add Key:=Range("B2:B553"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("1").Sort
        .SetRange Range("A1:B553")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("F6").Select
End Sub
Sub 羆()
Attribute 羆.VB_Description = "皊弘计秖羆"
Attribute 羆.VB_ProcData.VB_Invoke_Func = "b\n14"
'
' 羆 エ栋
' 皊弘计秖羆
'
' е硉龄: Ctrl+b
'
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-3]:R[49]C[-3])"
    Range("E2").Select
End Sub
Sub 程()
Attribute 程.VB_Description = "皊弘计秖程"
Attribute 程.VB_ProcData.VB_Invoke_Func = "c\n14"
'
' 程 エ栋
' 皊弘计秖程
'
' е硉龄: Ctrl+c
'
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=MAX(RC[-5]:R[49]C[-5])"
    Range("G1").Select
    Selection.ClearContents
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=MAX(R[1]C[-5]:R[49]C[-5])"
    Range("G2").Select
End Sub
Sub 程()
Attribute 程.VB_Description = "皊弘计秖程"
Attribute 程.VB_ProcData.VB_Invoke_Func = "d\n14"
'
' 程 エ栋
' 皊弘计秖程
'
' е硉龄: Ctrl+d
'
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "=MIN(R[1]C[-7]:R[49]C[-7])"
    Range("I2").Select
End Sub
