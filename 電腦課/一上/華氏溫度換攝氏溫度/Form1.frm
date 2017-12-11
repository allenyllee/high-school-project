VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "地ん放锣尼ん放"
   ClientHeight    =   2490
   ClientLeft      =   5820
   ClientTop       =   4530
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   ScaleHeight     =   2490
   ScaleWidth      =   3750
   Begin VB.CommandButton Command3 
      Caption         =   "挡"
      Height          =   372
      Left            =   2640
      TabIndex        =   4
      Top             =   1920
      Width           =   972
   End
   Begin VB.CommandButton Command2 
      Caption         =   "睲埃"
      Enabled         =   0   'False
      Height          =   372
      Left            =   1440
      TabIndex        =   3
      Top             =   1920
      Width           =   972
   End
   Begin VB.CommandButton Command1 
      Caption         =   "传衡"
      Height          =   372
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   1572
   End
   Begin VB.Label Label2 
      Height          =   492
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   3252
   End
   Begin VB.Label Label1 
      Caption         =   "地ん放:"
      Height          =   372
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1092
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Label2.Caption = "尼ん放" + Format(Str(Val(Text1.Text) - 32) * (5 / 9), "00.00") + ""
  Command2.Enabled = True
End Sub

Private Sub Command2_Click()
  Command2.Enabled = False
  Text1.Text = ""
  Label2.Caption = ""
End Sub

Private Sub Command3_Click()
  End
End Sub
