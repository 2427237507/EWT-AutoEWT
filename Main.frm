VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EWT简易挂机刷分程序"
   ClientHeight    =   10335
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   16950
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10335
   ScaleWidth      =   16950
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   15855
      Top             =   525
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   360
      Left            =   690
      TabIndex        =   12
      Top             =   9195
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.CommandButton Command2 
      Caption         =   "答案获取测试"
      Height          =   465
      Left            =   585
      TabIndex        =   9
      Top             =   8460
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.CheckBox Check1 
      Caption         =   "自动应答列表"
      Height          =   360
      Left            =   15015
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   60
      Width           =   1395
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      TabIndex        =   5
      Top             =   9885
      Width           =   16950
      _ExtentX        =   29898
      _ExtentY        =   794
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2461
            MinWidth        =   2470
            Text            =   "就绪"
            TextSave        =   "就绪"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4409
            MinWidth        =   4409
            Text            =   "挂机模式:无序刷分模式"
            TextSave        =   "挂机模式:无序刷分模式"
         EndProperty
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   8.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   12480
      Top             =   1275
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   16410
      Top             =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   360
      Left            =   16560
      TabIndex        =   4
      Top             =   60
      Width           =   285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "转到"
      Height          =   360
      Left            =   14070
      TabIndex        =   3
      Top             =   60
      Width           =   885
   End
   Begin VB.TextBox URL1 
      Height          =   375
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "https://www.ewt360.com"
      Top             =   60
      Width           =   13380
   End
   Begin VB.TextBox htmlcode 
      Height          =   1470
      Left            =   7590
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   7950
      Visible         =   0   'False
      Width           =   9150
   End
   Begin SHDocVwCtl.WebBrowser AnswerGetEngine 
      Height          =   1245
      Left            =   3555
      TabIndex        =   8
      Top             =   8205
      Visible         =   0   'False
      Width           =   4530
      ExtentX         =   7990
      ExtentY         =   2196
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   9150
      Left            =   75
      TabIndex        =   0
      Top             =   690
      Width           =   16770
      ExtentX         =   29580
      ExtentY         =   16140
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label listnum 
      Caption         =   "0"
      Height          =   225
      Left            =   5895
      TabIndex        =   18
      Top             =   420
      Width           =   870
   End
   Begin VB.Label Label3 
      Caption         =   "还剩项目(包括当前)："
      Height          =   225
      Left            =   4215
      TabIndex        =   17
      Top             =   420
      Width           =   1845
   End
   Begin VB.Label maxmin 
      Caption         =   "0"
      Height          =   225
      Left            =   3285
      TabIndex        =   16
      Top             =   420
      Width           =   870
   End
   Begin VB.Label Label2 
      Caption         =   "要求达到："
      Height          =   270
      Left            =   2415
      TabIndex        =   15
      Top             =   420
      Width           =   1170
   End
   Begin VB.Label min 
      Caption         =   "0"
      Height          =   300
      Left            =   1470
      TabIndex        =   13
      Top             =   420
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "当前已挂分钟："
      Height          =   225
      Left            =   240
      TabIndex        =   14
      Top             =   420
      Width           =   1545
   End
   Begin VB.Label tiku2 
      Caption         =   """>"
      Height          =   360
      Left            =   12780
      TabIndex        =   11
      Top             =   8805
      Width           =   2010
   End
   Begin VB.Label tiku 
      Caption         =   "/TiKuNew/HomeWorkQuestion?examid="
      Height          =   435
      Left            =   12765
      TabIndex        =   10
      Top             =   8430
      Width           =   2010
   End
   Begin VB.Label lbl_addr 
      Caption         =   "地址："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   1
      Top             =   105
      Width           =   1455
   End
   Begin VB.Menu menulist 
      Caption         =   "MENU"
      NegotiatePosition=   1  'Left
      Visible         =   0   'False
      Begin VB.Menu loadcfg 
         Caption         =   "加载自动应答文件"
      End
      Begin VB.Menu runcfg 
         Caption         =   "执行自动应答文件"
         Enabled         =   0   'False
      End
      Begin VB.Menu stopcfg 
         Caption         =   "停止执行自动应答文件"
         Enabled         =   0   'False
      End
      Begin VB.Menu makecfg 
         Caption         =   "制作自动应答文件"
         Enabled         =   0   'False
      End
      Begin VB.Menu background 
         Caption         =   "后台挂机"
      End
      Begin VB.Menu endexe 
         Caption         =   "结束退出"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 0 Then
AutoAnswer.Visible = False
Else
AutoAnswer.Visible = True
End If
End Sub

Private Sub Command1_Click()
WebBrowser1.Navigate "about:blank"
Do While WebBrowser1.Busy '等待加载完成.
    DoEvents
    Loop
WebBrowser1.Navigate URL1

End Sub

Private Sub Command2_Click()
AnswerGetEngine.Navigate "https://study.ewt360.com/TiKuNew/GetQuestionRecordInfo?questionid=244156"
End Sub

Private Sub Command3_Click()
'If menu.Visible = True Then
'menu.Visible = False
'Else
'menu.Visible = True
'End If
PopupMenu menulist
End Sub

Private Sub Command4_Click()
Dim start1 As Integer
Dim end1 As Integer
Dim tmp1 As String
Dim tmp2 As String
start1 = InStr(htmlcode, tiku) '+ 32

tmp1 = Right(htmlcode, (Len(htmlcode) - start1))

end1 = InStr(tmp1, tiku2)
tmp2 = Left(tmp1, (end1 - 1))
tmp1 = Right(tmp1, (Len(tmp1) - end1))
MsgBox tmp2
MsgBox tmp1

Do
start1 = InStr(htmlcode, tiku) '+ 32


'MsgBox tmp1
end1 = InStr(tmp1, tiku2)
tmp2 = Left(tmp1, (end1 - 1))
tmp1 = Right(tmp1, (Len(tmp1) - end1))
MsgBox tmp2

Loop Until tmp1 = ""
End Sub

Private Sub Form_Load()
WebBrowser1.Navigate URL1
'AnswerList.AddItem "China"
End Sub

Private Sub min_Change()
Dim min1, min2 As Integer
min2 = maxmin.Caption
min1 = min.Caption
If min1 = min2 + 1 Then
AutoAnswer.AnswerList.RemoveItem (0)
Dim str As String
Dim url2 As String
str = AutoAnswer.AnswerList.List(0)
If str = "" Then
MsgBox "Finish！"
Timer1 = False
Call stopcfg_click
WebBrowser1.Navigate "www.ewt360.com"
URL1 = "http://www.ewt360.com"
Exit Sub
End If
'MsgBox str
Dim a As Integer
a = InStr(str, "@")

maxmin = Right(str, (Len(str) - a))
'MsgBox maxmin
url2 = Left(str, (a - 1))
URL1 = url2
WebBrowser1.Navigate "about:blank"
WebBrowser1.Navigate URL1
min.Caption = "0"
End If
Exit Sub

End Sub

Private Sub runcfg_Click()
runcfg.Enabled = False
loadcfg.Enabled = False
stopcfg.Enabled = True
Dim str As String
Dim url2 As String
str = AutoAnswer.AnswerList.List(0)
'MsgBox str
Dim a As Integer
a = InStr(str, "@")
maxmin = Right(str, (Len(str) - a))
'MsgBox maxmin
url2 = Left(str, (a - 1))
URL1 = url2
'MsgBox url2
Timer1.Enabled = True
Call Command1_Click
runcfg.Enabled = False
End Sub


Private Sub loadcfg_Click()
On Error Resume Next
AutoAnswer.AnswerList.Clear
min.Caption = "0"

Dim sFile As String

With dlgCommonDialog
.DialogTitle = "Open"
.CancelError = False
.Filter = "自动应答文件(*.cfg)|*.cfg"
.ShowOpen
Dim data As String
Dim buffer As String
MsgBox dlgCommonDialog
Open dlgCommonDialog.FileName For Input As #1
Do While Not EOF(1)
Line Input #1, buffer
AutoAnswer.AnswerList.AddItem buffer
Loop
Close #1

runcfg.Enabled = True


StatusBar1.Panels(1).Text = "自动应答文件载入成功"
AutoAnswer.Show
Check1.Value = 1

If Len(.FileName) = 0 Then
StatusBar1.Panels(1).Text = "自动应答文件还没有载入"
Exit Sub
End If
sFile = .FileName
End With


End Sub
Private Sub endexe_click()
End

End Sub

Private Sub Timer1_Timer()
Do While WebBrowser1.LocationURL = url2

Loop
min = min + 1

End Sub

Private Sub Timer2_Timer()
listnum = AutoAnswer.AnswerList.ListCount
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
htmlcode.Text = WebBrowser1.Document.documentElement.outerHTML
End Sub

Private Sub stopcfg_click()
Timer1.Enabled = False
stopcfg.Enabled = False
runcfg.Enabled = False
loadcfg.Enabled = True

End Sub
Private Sub background_click()
Me.WindowState = 1
End Sub
