Attribute VB_Name = "Module1"
Sub �j��p()
Attribute �j��p.VB_Description = "�s��ƶq�Ѥj��p"
Attribute �j��p.VB_ProcData.VB_Invoke_Func = "a\n14"
'
' �j��p ����
' �s��ƶq�Ѥj��p
'
' �ֳt��: Ctrl+a
'
    Columns("B:B").Select
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add Key:=Range("B2:B553"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort
        .SetRange Range("A1:B553")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("F7").Select
End Sub
Sub �`�X()
Attribute �`�X.VB_Description = "�s��ƶq�`�X"
Attribute �`�X.VB_ProcData.VB_Invoke_Func = "c\n14"
'
' �`�X ����
' �s��ƶq�`�X
'
' �ֳt��: Ctrl+c
'
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(C[-3])"
    Range("E2").Select
    ActiveWindow.SmallScroll Down:=-21
    Range("E1").Select
    Selection.ClearContents
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-3]:R[49]C[-3])"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = ""
    Range("G5").Select
End Sub
Sub �̤j��()
Attribute �̤j��.VB_Description = "�s��ƶq�̤j��"
Attribute �̤j��.VB_ProcData.VB_Invoke_Func = "d\n14"
'
' �̤j�� ����
' �s��ƶq�̤j��
'
' �ֳt��: Ctrl+d
'
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=MAX(R[1]C[-5]:R[49]C[-5])"
    Range("G2").Select
End Sub
Sub �̤p��()
Attribute �̤p��.VB_Description = "�s��ƶq�̤p��"
Attribute �̤p��.VB_ProcData.VB_Invoke_Func = "e\n14"
'
' �̤p�� ����
' �s��ƶq�̤p��
'
' �ֳt��: Ctrl+e
'
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "=MIN((R[1]C[-7]:R[49]C[-7]))"
    Range("I2").Select
    ActiveWindow.SmallScroll Down:=-9
End Sub
