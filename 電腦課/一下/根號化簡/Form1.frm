VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2985
   ClientLeft      =   5355
   ClientTop       =   3900
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   4125
   Begin VB.CommandButton Command1 
      Caption         =   "求平方根"
      Default         =   -1  'True
      Height          =   492
      Left            =   1452
      TabIndex        =   2
      Top             =   2280
      Width           =   1212
   End
   Begin VB.TextBox Text1 
      Height          =   612
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   1332
   End
   Begin VB.Label Label2 
      Alignment       =   1  '靠右對齊
      Height          =   852
      Left            =   192
      TabIndex        =   3
      Top             =   1200
      Width           =   3732
   End
   Begin VB.Label Label1 
      Caption         =   "請輸入一個數字:"
      Height          =   612
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1932
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  a = Text1.Text
  If a = 0 Then
  Label2.Caption = 0
  Else
  If a > 0 Then
  For i = a To 1 Step -1
  If a Mod i ^ 2 = 0 Then
  Exit For
  End If
  Next
  If i = 1 Then
  Label2.Caption = "√" & a / i ^ 2
  Else
  If a / i ^ 2 = 1 Then
  Label2.Caption = i
  Else
  Label2.Caption = i & "√" & a / i ^ 2
  End If
  End If
  Else
  For i = a To 1 Step 1
  If a Mod i ^ 2 = 0 Then
  Exit For
  End If
  Next
  If i = -1 Then
  Label2.Caption = "i"
  Else
  If a / i ^ 2 = -1 Then
  Label2.Caption = -i & "i"
  Else
  Label2.Caption = -i & "√" & -(a / i ^ 2) & "i"
  End If
  End If
  End If
  End If
End Sub
