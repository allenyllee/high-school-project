VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3330
   ClientLeft      =   5295
   ClientTop       =   3840
   ClientWidth     =   4140
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   4140
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
  Dim a As Double
  a = Val(InputBox("請輸入西元年份", "輸入"))
  If a Mod 4 = 0 And a Mod 100 <> 0 Or a Mod 400 = 0 Then Print "閏年" Else Print "平年"
  End Sub
