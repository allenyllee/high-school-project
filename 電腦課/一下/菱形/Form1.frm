VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   3750
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "輸入"
      Height          =   372
      Left            =   2640
      TabIndex        =   0
      Top             =   3480
      Width           =   972
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Cls
x = Val(InputBox("請輸入介於0~9的值", "請輸入介於0~9的值"))
b = x
  If x <= 9 And x > 0 Then
    For i = -x To x
      a = Abs(i)
      Print Spc(Abs(a));
      For j = b To a Step -1
        Print j;
      Next j
      Print
    Next i
  Else
    MsgBox "請重新輸入", , "警告"
  End If
End Sub

Private Sub Form_Activate()
x = Val(InputBox("請輸入介於0~9的值", "請輸入介於0~9的值"))
b = x
  If x <= 9 And x > 0 Then
    For i = -x To x
      a = Abs(i)
      Print Spc(Abs(a));
      For j = b To a Step -1
        Print j;
      Next j
      Print
    Next i
  Else
    MsgBox "請重新輸入", , "警告"
  End If
End Sub
