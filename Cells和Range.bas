Attribute VB_Name = "Module1"
Sub јЖѕЪЅmІЯ()
Cells(1, 5).Value = Cells(1, 1).Value + Cells(1, 3).Value 'E1Дж=A1Дж+C1Дж'
Cells(1, "E").Value = Cells(1, "A").Value + Cells(1, "C").Value 'E1Дж=A1Дж+C1Дж'
Range("E1").Value = Range("A1").Value + Range("C1").Value 'E1Дж=A1Дж+C1Дж'

Cells(2, 5).Value = Cells(1, 1).Value - Cells(1, 3).Value 'E2Дж=A1Дж-C1Дж'
Cells(2, "E").Value = Cells(1, "A").Value - Cells(1, "C").Value 'E2Дж=A1Дж-C1Дж'
Range("E2").Value = Range("A1").Value - Range("C1").Value 'E2Дж=A1Дж-C1Дж'

Cells(3, 5).Value = Cells(1, 1).Value * Cells(1, 3).Value 'E3Дж=A1Дж*C1Дж'
Cells(3, "E").Value = Cells(1, "A").Value * Cells(1, "C").Value 'E3Дж=A1Дж*C1Дж'
Range("E3").Value = Range("A1").Value * Range("C1").Value 'E3Дж=A1Дж*C1Дж'

Cells(4, 5).Value = Cells(1, 1).Value / Cells(1, 3).Value 'E4Дж=A1Дж/C1Дж'
Cells(4, "E").Value = Cells(1, "A").Value / Cells(1, "C").Value 'E4Дж=A1Дж/C1Дж'
Range("E4").Value = Range("A1").Value / Range("C1").Value 'E4Дж=A1Дж/C1Дж'
End Sub
