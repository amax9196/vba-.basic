Attribute VB_Name = "Module1"
Option Explicit

Sub SelectCase()
Select Case Range("A2").Value
Case "�w��¯l"
Range("B2").Value = "�w��"
Case "�饻����"
Range("B2").Value = "�饻"
Case "�D�w�޽E"
Range("B2").Value = "�D�w"
End Select
End Sub
Sub IfCal()
If (Range("B1").Value > 38) Then
Range("B2").Value = "���g��"
Else
Range("B2").Value = "�L�g��"
End If
End Sub
Sub For�j��()
Dim i As Integer
For i = 2 To 100
Cells(i, 4).Value = Cells(i, 2).Value * Cells(i, 3)
Next
End Sub
