VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�ݭԵ{��"
   ClientHeight    =   2490
   ClientLeft      =   5640
   ClientTop       =   4530
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   ScaleHeight     =   2490
   ScaleWidth      =   3885
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   492
      Left            =   2040
      TabIndex        =   1
      Top             =   1800
      Width           =   1452
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��J�m�W"
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
  c = InputBox("�п�J�z���m�W", "�аݪڦW")
  If c = "" Then Else a = MsgBox(c + "�O�A���m�W��?", 36)
  If a = 6 Then Print c + "�A�n"
End Sub

Private Sub Command2_Click()
  b = MsgBox("�T�w�n���}�F��", 65)
  If b = 1 Then End
End Sub
