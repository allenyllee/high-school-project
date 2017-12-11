VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "尼ん放锣地ん放"
   ClientHeight    =   2490
   ClientLeft      =   5640
   ClientTop       =   4365
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   ScaleHeight     =   2490
   ScaleWidth      =   3750
   Begin VB.CommandButton Command2 
      Caption         =   "挡"
      Height          =   300
      Left            =   1920
      TabIndex        =   4
      Top             =   2040
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "睲埃"
      Enabled         =   0   'False
      Height          =   300
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   1212
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   1572
   End
   Begin VB.Label Label2 
      Height          =   612
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   2892
   End
   Begin VB.Label Label1 
      Caption         =   "尼ん放:"
      Height          =   252
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   852
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Text1.Text = ""
  Label2.Caption = ""
  Command1.Enabled = False
End Sub

Private Sub Command2_Click()
  End
End Sub

Private Sub Text1_Change()
  Label2.Caption = "地ん放" + Str(Val(Text1.Text) * (9 / 5) + 32) + ""
  Command1.Enabled = True
End Sub
