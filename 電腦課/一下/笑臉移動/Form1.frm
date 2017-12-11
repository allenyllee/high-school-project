VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   4605
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "結束"
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  '沒有框線
      Height          =   225
      Left            =   2760
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   840
      Width           =   225
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Const a = 38
 Const b = 40
 Const c = 39
 Const d = 37
 Const e = 33
 Const f = 36
 Const g = 35
 Const h = 34
Private Sub Command1_Click()
  End
End Sub
 Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
  Cls
  Print "keycode:"; KeyCode, "shift:"; Shift
  Select Case KeyCode
    Case a
      Picture1.Top = Picture1.Top - 100
    Case b
      Picture1.Top = Picture1.Top + 100
    Case c
      Picture1.Left = Picture1.Left + 100
    Case d
      Picture1.Left = Picture1.Left - 100
    Case e
      Picture1.Top = Picture1.Top - 100
      Picture1.Left = Picture1.Left + 100
    Case f
      Picture1.Top = Picture1.Top - 100
      Picture1.Left = Picture1.Left - 100
    Case g
      Picture1.Top = Picture1.Top + 100
      Picture1.Left = Picture1.Left - 100
    Case h
      Picture1.Top = Picture1.Top + 100
      Picture1.Left = Picture1.Left + 100
  End Select
  If Picture1.Top <= -Picture1.Height Then Picture1.Top = Form1.Height
  If Picture1.Top >= Form1.Height + Picture1.Height Then Picture1.Top = -Picture1.Height
  If Picture1.Left <= -Picture1.Width Then Picture1.Left = Form1.Width
  If Picture1.Left >= Form1.Width + Picture1.Width Then Picture1.Left = -Picture1.Width
End Sub
