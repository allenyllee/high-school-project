VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "約分"
   ClientHeight    =   3195
   ClientLeft      =   5115
   ClientTop       =   3855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   2505
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "結束"
      Height          =   495
      Left            =   2396
      TabIndex        =   1
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "請按這輸入"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      Caption         =   "="
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2190
      TabIndex        =   4
      Top             =   1200
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A As Long
Dim B As Long
Dim i As Long
Private Sub Command1_Click()
  A = Val(InputBox("請輸入分子", "請輸入分子"))
  B = Val(InputBox("請輸入分母", "請輸入分母"))
  
  If B = 0 Then
    MsgBox "分母不可為零"
    Text1.Text = ""
    Text2.Text = ""
  Else
    Text1.Text = A & "/" & B
    
    For i = 2 To A
      If (A Mod i = 0) And (B Mod i = 0) Then
        A = A / i
        B = B / i
        i = i - 1
      End If
    Next i
    
    If B = 1 Then
      Text2.Text = A
    Else
      If A = 0 Then
        Text2.Text = 0
      Else
        Text2.Text = A & "/" & B
      End If
    End If
  End If
End Sub

Private Sub Command2_Click()
  End
End Sub

