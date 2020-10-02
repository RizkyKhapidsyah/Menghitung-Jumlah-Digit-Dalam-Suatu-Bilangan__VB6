VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menghitung Jumlah Digit dalam Suatu Bilangan"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   1800
      TabIndex        =   0
      Top             =   960
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function GetDigitCount(inValue As Double) _
As Double
   GetDigitCount = Int(Log(inValue) / Log(10)) + 1
End Function

Private Sub Command1_Click()
   'Ganti '123456789' dengan bilangan yang Anda
   'inginkan untuk dihitung jumlah digitnya.
   MsgBox GetDigitCount(123456789)
End Sub

