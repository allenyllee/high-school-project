VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2925
   ClientLeft      =   5115
   ClientTop       =   4650
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   4755
   Begin VB.ListBox List1 
      Height          =   1500
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":0002
      TabIndex        =   5
      Top             =   720
      Width           =   4455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "結束"
      Height          =   372
      Left            =   2280
      TabIndex        =   3
      Top             =   2400
      Width           =   1092
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Left            =   2160
      MaxLength       =   4
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "答案"
      Height          =   372
      Left            =   3480
      TabIndex        =   0
      Top             =   2400
      Width           =   1092
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      Caption         =   "請輸入不同的數字："
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Menu Game 
      Caption         =   "遊戲(&G)"
      Begin VB.Menu ResetGame 
         Caption         =   "重新玩(&R)"
      End
      Begin VB.Menu HardGame 
         Caption         =   "進階(&H)"
      End
   End
   Begin VB.Menu About 
      Caption         =   "關於(&A)"
      Begin VB.Menu version 
         Caption         =   "版本(&V)"
      End
      Begin VB.Menu Maker 
         Caption         =   "作者(&M)"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim b As Integer
Dim c As String
Dim m As Integer
Dim n As String
Private Sub Command1_Click()
  Label1.Caption = c
End Sub
Private Sub Command2_Click()
  End
End Sub
Private Sub Form_Load()
  m = 4
  Randomize
  e = "1234567890"
  For i = 1 To m
    x = Len(e)
    y = Int(Rnd * x) + 1
    c = c + Mid(e, y, 1)
    e = Left(e, y - 1) & Right(e, x - y)
  Next i
End Sub
Private Sub HardGame_Click()
  n = InputBox("你可以輸入2~9來決定所猜的字數", "請輸入")
  If n = "" Then
  Else
    If n < 2 Or n > 9 Then
      MsgBox "請輸入2~9之間的數字", 0, "警告"
    Else
      List1.Clear
      Text1.Text = ""
      Label1.Caption = ""
      c = ""
      Randomize
      e = "1234567890"
      For i = 1 To n
        x = Len(e)
        y = Int(Rnd * x) + 1
        c = c + Mid(e, y, 1)
        e = Left(e, y - 1) & Right(e, x - y)
      Next i
      Text1.MaxLength = n
      m = n
    End If
  End If
End Sub


Private Sub Maker_Click()
  MsgBox "作者：Allen     電子郵件：allenli@msn.com", 0, "作者"
End Sub

Private Sub ResetGame_Click()
  List1.Clear
  Text1.Text = ""
  Label1.Caption = ""
  c = ""
  Randomize
  e = "1234567890"
  For i = 1 To m
    x = Len(e)
    y = Int(Rnd * x) + 1
    c = c + Mid(e, y, 1)
    e = Left(e, y - 1) & Right(e, x - y)
  Next i
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    ans = Text1.Text
    x = Len(ans)
    If x = m Then
      For i = 1 To m - 1
        For j = 2 To m
          If i < j And Mid(ans, i, 1) = Mid(ans, j, 1) Then
            MsgBox "請輸入不同的數字", 0, "警告"
            Text1.Text = ""
            Exit For
          End If
        Next j
        If Text1.Text = "" Then Exit For
      Next i
      If Text1.Text <> "" Then
        For i = 1 To m
          If Mid(c, i, 1) = Mid(ans, i, 1) Then a = a + 1
          For j = 1 To m
            If i <> j And Mid(c, i, 1) = Mid(ans, j, 1) Then b = b + 1
          Next j
        Next i
        List1.AddItem ans & " ---> " & a & "A" & b & "B"
        If a = n Then List1.AddItem "你成功了"
        a = 0
        b = 0
        Text1.Text = ""
      End If
    Else
      MsgBox "請輸入" & m & "個數字", 0, "警告"
    End If
  End If
End Sub
Private Sub version_Click()
MsgBox "版本1.00beta", 0, "版本"
End Sub
