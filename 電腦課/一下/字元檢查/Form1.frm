VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3915
   ClientLeft      =   5355
   ClientTop       =   3900
   ClientWidth     =   4275
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   4275
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�ˬd"
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label2 
      Height          =   1575
      Left            =   600
      TabIndex        =   4
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  '�m�����
      Caption         =   "�п�J�@�Ӧr���G"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
aa = Text1.Text
Select Case aa
  Case "a" To "z"
  Label2.Caption = "�u" + aa + "�v" + "�O�@�Ӥp�g�^��r��"
  Case "A" To "Z"
  Label2.Caption = "�u" + aa + "�v" + "�O�@�Ӥj�g�^��r��"
  Case 1 To 9
  Label2.Caption = "�u" + aa + "�v" + "�O�@�Ӫ��ԧB�Ʀr"
  Case Is >= "�@"
  Label2.Caption = "�u" + aa + "�v" + "�O�@�Ӥ���r"
  Case Else
  Label2.Caption = "�u" + aa + "�v" + "�O�@�ӲŸ�"
  End Select
End Sub

Private Sub Command2_Click()
End
End Sub

