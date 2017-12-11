VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   3168
   ClientLeft      =   4308
   ClientTop       =   3072
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   ScaleHeight     =   3168
   ScaleWidth      =   3720
   Begin VB.CommandButton Command2 
      Caption         =   "結束"
      Height          =   495
      Left            =   2094
      TabIndex        =   3
      Top             =   2400
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "請輸入x值"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   414
      TabIndex        =   1
      Top             =   2400
      Width           =   1212
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      Caption         =   "求e的x次方"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   294
      TabIndex        =   2
      Top             =   600
      Width           =   3132
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   414
      TabIndex        =   0
      Top             =   1440
      Width           =   2892
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Label1.Caption = ""
b = 1
y = 1
d = Val(InputBox("請輸入X值", "輸入", "1"))
  
  For i = 1 To 18
  x = d ^ i
  y = y * i
  e = x / y
  b = b + e
  a = a + 1
  Next i
  
 Label1.Caption = Format(b, "##.0000000")
End Sub

Private Sub Command2_Click()
  End
End Sub
