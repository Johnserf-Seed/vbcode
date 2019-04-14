VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Tools"
   ClientHeight    =   4500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6075
   Icon            =   "Tools.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   4500
   ScaleWidth      =   6075
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox Check1 
      Caption         =   "酷狗"
      Height          =   495
      Index           =   2
      Left            =   1680
      TabIndex        =   12
      Top             =   2040
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "网易云"
      Height          =   495
      Index           =   1
      Left            =   840
      TabIndex        =   11
      Top             =   2040
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "QQ音乐"
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   10
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   9
      Text            =   "请输入歌手或歌名"
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "PictureView"
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "度盘快速下载"
      Enabled         =   0   'False
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "vip音乐解析"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "解压缩"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "浏览器"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "功能明天更新"
      Height          =   495
      Left            =   480
      TabIndex        =   8
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "单纯调用rar，很鸡肋的功能对吧"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "槽点很多的浏览器"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "照片查看器"
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim jiexi As String, wz As String, i%, N%
Option Explicit '强制变量
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Check1_Click(Index As Integer)
    'if check1(index).Value=1
    jiexi = "http://tool.liumingye.cn/music/?type=qq&name=" & Text1.Text
    Dim N As Integer
        If Check1(1).Value = vbChecked Then
            wz = "qq"
            jiexi = "http://tool.liumingye.cn/music/?type= & wz & qq&name=" & Text1.Text
        ElseIf Check1(2).Value = vbChecked Then
            wz = "netease"
        Else
            wz = "kugou"
        End If
        If Check1(0).Value = 1 Then
           Check1(2).Value = 0: Check1(1).Value = 0
        ElseIf Check1(1).Value = 1 Then Check1(0).Value = 0: Check1(2).Value = 0
        Else
            Check1(2).Value = 1: Check1(0).Value = 0: Check1(1).Value = 0
        End If

End Sub

Private Sub Command1_Click()

    Form3.Show
    
End Sub

Private Sub Command2_Click()

    Form4.Show

End Sub

Private Sub Command3_Click()

        ShellExecute Me.hWnd, "open", jiexi, "", "", 1
        'Form3.Text1.Text = jiexi
        'Form3.Show
        'Form3.WebBrowser1.Navigate "http://tool.liumingye.cn/music/?type=qq&name=" & Text1.Text
        'Shell "explorer.exe http://tool.liumingye.cn/music/?type=qq&name= & Text1.Text"
End Sub

Private Sub Text1_Click()

    Text1.ForeColor = vbBlack
    Text1.Text = ""
    
End Sub

Private Sub Command5_Click()

    Form6.Show
    
End Sub

Private Sub Form_Load()

    MsgBox "布局很单调，想把很多的软件的功能集成在一起，会不断关系的，在寒假结束前会出个EA版本", vbInformation, "作者的话"
    Form5.Width = 5895
    Form5.Left = 11450
    Form5.Top = 5600
    Text1.ForeColor = &HC0C0C0

End Sub

