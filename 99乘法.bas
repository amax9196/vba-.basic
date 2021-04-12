Attribute VB_Name = "Module1"
Option Explicit

Sub nineXnine()
Dim i, s As Integer
For i = 1 To 9
For s = 1 To 9
Cells(i, s).Value = i & "*" & s & "=" & i * s
Next
Next
End Sub
