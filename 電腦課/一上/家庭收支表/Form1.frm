VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3330
   ClientLeft      =   4605
   ClientTop       =   3840
   ClientWidth     =   5325
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   5325
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   372
      Left            =   1800
      TabIndex        =   0
      Top             =   2760
      Width           =   1692
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  End
End Sub

Private Sub Form_Activate()
  Print Tab(10); "<<�a�x�����>>";
  Print Spc(5); "���:"; Format(Date, "yyyy-m-d")
  Print
  Print "����"; Spc(4); "�Ƶ�"; Spc(11); "���J���B"; Spc(4); "��X���B"
  Print
  Print "�~��"; Spc(4); "�Ѫ���"; Spc(9); Format(48560, "+0,000.00")
  Print "�~��"; Spc(4); "�Ѷ���"; Spc(9); Format(62000, "+0,000.00")
  Print "�Я�"; Spc(31); Format(21000, "-0,000.00")
  Print "����"; Spc(4); "�Τ@�o������"; Spc(3); Format(2000, "+0,000.00")
  Print "�ǶO"; Spc(4); "�s�����֯Z"; Spc(17); Format(6000, "-0,000.00")
End Sub
