VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "問候程式"
   ClientHeight    =   2490
   ClientLeft      =   5640
   ClientTop       =   4530
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   ScaleHeight     =   2490
   ScaleWidth      =   3885
   Begin VB.CommandButton Command2 
      Caption         =   "結束"
      Height          =   492
      Left            =   2040
      TabIndex        =   1
      Top             =   1800
      Width           =   1452
   End
   Begin VB.CommandButton Command1 
      Caption         =   "輸入姓名"
      Height          =   492
      Left            =   360
      TabIndex        =   0
      Top             =   1800
      Width           =   1452
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Cls
  c = InputBox("請輸入您的姓名", "請問芳名")
  If c = "" Then Else a = MsgBox(c + "是你的姓名嗎?", 36)
  If a = 6 Then Print c + "你好"
End Sub

Private Sub Command2_Click()
  b = MsgBox("確定要離開了嗎", 65)
  If b = 1 Then End
End Sub
