VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3132
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   4236
   LinkTopic       =   "Form1"
   ScaleHeight     =   3132
   ScaleWidth      =   4236
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
  a = 2
  b = 5
  c = 2
  x = (-b + (b ^ 2 - 4 * a * c) ^ 0.5) / (2 * a)
  X1 = (-b - (b ^ 2 - 4 * a * c) ^ 0.5) / (2 * a)
  Print x; "或"; X1
End Sub

