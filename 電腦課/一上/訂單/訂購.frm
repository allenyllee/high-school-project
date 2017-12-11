VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "訂單"
   ClientHeight    =   3264
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   4320
   Icon            =   "訂購.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3264
   ScaleWidth      =   4320
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command3 
      Caption         =   "結束"
      Height          =   732
      Left            =   3000
      Picture         =   "訂購.frx":0442
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   2400
      Width           =   972
   End
   Begin VB.CommandButton Command2 
      Caption         =   "重填"
      Enabled         =   0   'False
      Height          =   732
      Left            =   1560
      Picture         =   "訂購.frx":0884
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   2400
      Width           =   972
   End
   Begin VB.CommandButton Command1 
      Caption         =   "傳出"
      Height          =   732
      Left            =   240
      Picture         =   "訂購.frx":0CC6
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   2400
      Width           =   972
   End
   Begin VB.TextBox Text2 
      Height          =   372
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   2412
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   2412
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   594
      TabIndex        =   4
      Top             =   1680
      Width           =   3132
   End
   Begin VB.Label Label2 
      Caption         =   "產品"
      Height          =   252
      Left            =   480
      TabIndex        =   3
      Top             =   1080
      Width           =   372
   End
   Begin VB.Label Label1 
      Caption         =   "廠牌"
      Height          =   252
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   372
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Label3.Caption = "謝謝您訂購" + Text1.Text + Text2.Text
  Label3.Visible = True
  Command2.Enabled = True
End Sub

Private Sub Command2_Click()
  Text1 = ""
  Text2 = ""
  Label3.Visible = False
  Command2.Enabled = False
End Sub

Private Sub Command3_Click()
  End
End Sub



