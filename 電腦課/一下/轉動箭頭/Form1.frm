VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2484
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   6084
   LinkTopic       =   "Form1"
   ScaleHeight     =   2484
   ScaleWidth      =   6084
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   2040
   End
   Begin VB.Image Image5 
      Height          =   972
      Left            =   0
      Top             =   480
      Width           =   1212
   End
   Begin VB.Image Image4 
      Height          =   384
      Left            =   3480
      Picture         =   "Form1.frx":0000
      Top             =   1920
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Image3 
      Height          =   384
      Left            =   2520
      Picture         =   "Form1.frx":0442
      Top             =   1920
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Image2 
      Height          =   384
      Left            =   1680
      Picture         =   "Form1.frx":0884
      Top             =   1920
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Image1 
      Height          =   384
      Left            =   960
      Picture         =   "Form1.frx":0CC6
      Top             =   1920
      Visible         =   0   'False
      Width           =   384
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim a As Integer

Private Sub Form_Load()
  Timer1.Interval = 100
  a = 0
End Sub

Private Sub Timer1_Timer()
  If a = 0 Then
    Image5.Picture = Image1
    Image5.Left = Image5.Left + 150
    Image5.Top = Image5.Top + 500
    a = 1
  Else
    If a = 1 Then
      Image5.Picture = Image2
      Image5.Left = Image5.Left + 150
      Image5.Top = Image5.Top - 500
      a = 2
    Else
      If a = 2 Then
        Image5.Picture = Image3
        Image5.Left = Image5.Left + 150
        Image5.Top = Image5.Top + 500
        a = 3
      Else
        If a = 3 Then
          Image5.Picture = Image4
          Image5.Left = Image5.Left + 150
          Image5.Top = Image5.Top - 500
          a = 0
        End If
      End If
    End If
  End If
  If Image5.Left >= Width Then
    Image5.Left = -Image5.Width
  End If
  
  
  
  
  
  
  
End Sub
