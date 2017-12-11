VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '單線固定
   Caption         =   "Form1"
   ClientHeight    =   2544
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   6036
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2544
   ScaleWidth      =   6036
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   2160
   End
   Begin VB.CommandButton Command1 
      Caption         =   "結束"
      Height          =   372
      Left            =   2160
      TabIndex        =   3
      Top             =   2040
      Width           =   1452
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   252
      LargeChange     =   50
      Left            =   600
      Max             =   500
      Min             =   10
      SmallChange     =   10
      TabIndex        =   0
      Top             =   1680
      Value           =   10
      Width           =   3492
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   600
      TabIndex        =   4
      Top             =   2040
      Width           =   1092
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "慢"
      Height          =   372
      Left            =   4200
      TabIndex        =   2
      Top             =   1560
      Width           =   372
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "快"
      Height          =   372
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   372
   End
   Begin VB.Image Image1 
      Height          =   1248
      Left            =   0
      Picture         =   "Form1.frx":0000
      Top             =   120
      Width           =   984
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Integer
Private Sub Command1_Click()
  End
End Sub

Private Sub Form_Activate()
  HScroll1.Value = 10
  Label3.Caption = "速度:0.01秒"
End Sub

Private Sub HScroll1_Change()
  Timer1.Interval = HScroll1.Value
  Label3.Caption = "速度:" & HScroll1.Value / 1000 & "秒"
End Sub

Private Sub Timer1_Timer()
  If x = 0 Then
    Image1.Left = Image1.Left + 100
  Else
    Image1.Left = Image1.Left - 100
  End If
  If Image1.Left >= Form1.Width - Image1.Width Then x = 1
  If Image1.Left = 0 Then x = 0
End Sub
