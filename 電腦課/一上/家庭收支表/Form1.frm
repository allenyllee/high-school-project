VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3330
   ClientLeft      =   4605
   ClientTop       =   3840
   ClientWidth     =   5325
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   5325
   Begin VB.CommandButton Command1 
      Caption         =   "結束"
      Height          =   372
      Left            =   1800
      TabIndex        =   0
      Top             =   2760
      Width           =   1692
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  End
End Sub

Private Sub Form_Activate()
  Print Tab(10); "<<家庭收支表>>";
  Print Spc(5); "日期:"; Format(Date, "yyyy-m-d")
  Print
  Print "項目"; Spc(4); "備註"; Spc(11); "收入金額"; Spc(4); "支出金額"
  Print
  Print "薪資"; Spc(4); "老爸的"; Spc(9); Format(48560, "+0,000.00")
  Print "薪資"; Spc(4); "老媽的"; Spc(9); Format(62000, "+0,000.00")
  Print "房租"; Spc(31); Format(21000, "-0,000.00")
  Print "獎金"; Spc(4); "統一發票中獎"; Spc(3); Format(2000, "+0,000.00")
  Print "學費"; Spc(4); "山葉音樂班"; Spc(17); Format(6000, "-0,000.00")
End Sub
