VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7920
   ClientLeft      =   5115
   ClientTop       =   3660
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   4605
   Begin VB.OptionButton Option9 
      Caption         =   "天藍色"
      Height          =   852
      Left            =   3076
      TabIndex        =   13
      Top             =   5760
      Width           =   1092
   End
   Begin VB.OptionButton Option8 
      Caption         =   "黃"
      Height          =   852
      Left            =   1756
      TabIndex        =   12
      Top             =   5760
      Width           =   1092
   End
   Begin VB.OptionButton Option7 
      Caption         =   "白"
      Height          =   852
      Left            =   437
      TabIndex        =   11
      Top             =   5760
      Width           =   1092
   End
   Begin VB.CheckBox Check1 
      Caption         =   "粗體"
      Height          =   852
      Left            =   437
      TabIndex        =   10
      Top             =   1492
      Width           =   1092
   End
   Begin VB.CheckBox Check2 
      Caption         =   "斜體"
      Height          =   852
      Left            =   1756
      TabIndex        =   9
      Top             =   1492
      Width           =   1092
   End
   Begin VB.CheckBox Check3 
      Caption         =   "底線"
      Height          =   852
      Left            =   3076
      TabIndex        =   8
      Top             =   1492
      Width           =   1092
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   28.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   466
      TabIndex        =   7
      Text            =   "李亞倫"
      Top             =   240
      Width           =   3492
   End
   Begin VB.OptionButton Option6 
      Caption         =   "藍"
      Height          =   852
      Left            =   3076
      TabIndex        =   6
      Top             =   4334
      Width           =   1092
   End
   Begin VB.OptionButton Option5 
      Caption         =   "綠"
      Height          =   852
      Left            =   1756
      TabIndex        =   5
      Top             =   4334
      Width           =   1092
   End
   Begin VB.OptionButton Option4 
      Caption         =   "紅"
      Height          =   852
      Left            =   437
      TabIndex        =   4
      Top             =   4334
      Width           =   1092
   End
   Begin VB.OptionButton Option1 
      Caption         =   "靠左"
      Height          =   852
      Left            =   437
      TabIndex        =   3
      Top             =   2913
      Value           =   -1  'True
      Width           =   1092
   End
   Begin VB.OptionButton Option2 
      Caption         =   "置中"
      Height          =   852
      Left            =   1764
      TabIndex        =   2
      Top             =   2913
      Width           =   1092
   End
   Begin VB.OptionButton Option3 
      Caption         =   "靠右"
      Height          =   852
      Left            =   3008
      TabIndex        =   1
      Top             =   2913
      Width           =   1092
   End
   Begin VB.CommandButton Command2 
      Caption         =   "關閉"
      Height          =   492
      Left            =   1560
      TabIndex        =   0
      Top             =   7200
      Width           =   1212
   End
   Begin VB.Label Label4 
      Caption         =   "變化："
      Height          =   255
      Left            =   437
      TabIndex        =   17
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "對齊："
      Height          =   255
      Left            =   437
      TabIndex        =   16
      Top             =   2501
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "文字色："
      Height          =   255
      Left            =   437
      TabIndex        =   15
      Top             =   3922
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "背景色："
      Height          =   255
      Left            =   437
      TabIndex        =   14
      Top             =   5343
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
If Check1.Value = 1 Then Text1.FontBold = True
If Check1.Value = 0 Then Text1.FontBold = False
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then Text1.FontItalic = True
If Check2.Value = 0 Then Text1.FontItalic = False
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then Text1.FontUnderline = True
If Check3.Value = 0 Then Text1.FontUnderline = False
End Sub

Private Sub Command2_Click()
  End
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then Text1.Alignment = 0
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then Text1.Alignment = 2
End Sub

Private Sub Option3_Click()
If Option3.Value = True Then Text1.Alignment = 1
End Sub

Private Sub Option4_Click()
Text1.ForeColor = vbRed
End Sub

Private Sub Option5_Click()
Text1.ForeColor = vbGreen
End Sub

Private Sub Option6_Click()
Text1.ForeColor = vbBlue
End Sub

Private Sub Option7_Click()
Text1.BackColor = RGB(255, 255, 255)
End Sub

Private Sub Option8_Click()
Text1.BackColor = vbYellow
End Sub

Private Sub Option9_Click()
Text1.BackColor = RGB(80, 200, 255)

End Sub
