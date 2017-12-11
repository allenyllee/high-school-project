VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   5475
   ClientTop       =   4005
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
  Dim a As Integer
  Dim b As Integer
  a = 10
  b = 8
    a = b - a
    b = b - a
    a = a + b
  Print a
  Print b
  '-------------------------------------------------------------------------------------------------
  a = 10
  b = 8
    b = b * a
    a = b / a
    b = b * a * (1 / (a * a))
  Print a
  Print b
  '--------------------------------------------------------------------------------------------------
  a = 10
  b = 8
    a = Sqr(a * b)
    b = (a * a) / b
    a = (a * a) / b
  Print a
  Print b
End Sub
