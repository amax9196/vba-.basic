Attribute VB_Name = "Module1"
Option Explicit

Sub SelectCase()
Select Case Range("A2").Value
Case "德國麻疹"
Range("B2").Value = "德國"
Case "日本腦炎"
Range("B2").Value = "日本"
Case "非洲豬瘟"
Range("B2").Value = "非洲"
End Select
End Sub
Sub IfCal()
If (Range("B1").Value > 38) Then
Range("B2").Value = "有症狀"
Else
Range("B2").Value = "無症狀"
End If
End Sub
Sub For迴圈()
Dim i As Integer
For i = 2 To 100
Cells(i, 4).Value = Cells(i, 2).Value * Cells(i, 3)
Next
End Sub
