VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "¼g­Èform.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label3_Click()

End Sub

Private Sub btna_Click()
Dim supplyName As String
supplyName = txbName.Text
Cells(2, 1).Value = supplyName

Dim supplyPhone  As String
supplyPhone = txbPhone.Text
Cells(2, 2).Value = supplyPhone

Dim price As Integer
price = txbPrice.Text
Cells(2, 3).Value = CInt(price)

Dim newPrice As Integer
newPrice = txbFinalPrice.Text
Cells(2, 4).Value = CInt(newPrice)

Dim totalDiscount As Single
totalDiscount = (price - newPrice) / price
Cells(2, 5).Value = totalDiscount
End Sub
