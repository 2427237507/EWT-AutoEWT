VERSION 5.00
Begin VB.Form AutoAnswer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "自动应答列表"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   7110
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox AnswerList 
      Height          =   5820
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      Width           =   7080
   End
End
Attribute VB_Name = "AutoAnswer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Visible = False
'AnswerList.AddItem "6"
End Sub
