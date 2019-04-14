VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "历史记录"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9660
   Icon            =   "musichistory.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   9660
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   4800
      Picture         =   "musichistory.frx":424A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   13
      Top             =   1320
      Width           =   480
   End
   Begin VB.CommandButton Command9 
      Caption         =   "意见反馈"
      Height          =   495
      Left            =   4320
      TabIndex        =   12
      Top             =   720
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   3600
      Left            =   5880
      Picture         =   "musichistory.frx":1C7B4
      ScaleHeight     =   3600
      ScaleWidth      =   3600
      TabIndex        =   10
      Top             =   -120
      Width           =   3600
   End
   Begin VB.CommandButton Command8 
      Caption         =   "↓"
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "↑"
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "-"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "+"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "关于作者"
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "返回播放器"
      Height          =   615
      Left            =   3480
      TabIndex        =   4
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "删除记录"
      Height          =   615
      Left            =   1800
      TabIndex        =   3
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "恢复记录"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   3480
      ItemData        =   "musichistory.frx":20455
      Left            =   120
      List            =   "musichistory.frx":20457
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   $"musichistory.frx":20459
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5880
      TabIndex        =   11
      Top             =   3480
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "播放记录："
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    Dim l1%, j%
    l1 = List1.ListIndex
    For j = 1 To l1
        Call MyLog(Time & "恢复第" & j & "记录")
    Next j

    MsgBox "暂且不支持回溯单一data文件，只能恢复软件上一次使用过的记录,以后会完善", vbInformation, "温馨提醒"
    MsgBox "现在你可以返回播放器啦", vbYesNoCancel, "温馨提醒"

'Shell "App.path:\data.dat" & List1.List(0) & "", vbNormalFocus
Dim Path As String, Item As Long, Temp As String, i As Long
    Path = IIf(Len(App.Path) = 3, App.Path, App.Path & "\") & "data.dat"
    If Dir(Path, vbDirectory) = "" Then Exit Sub
    Open Path For Input As #1
    Line Input #1, Temp
    Item = Val(Temp)
        For i = 1 To Item
            Line Input #1, Temp
            Form1.List1.AddItem Temp
        Next i
    Close #1
    'Form1.List1.AddItem Temp  可以实现同样的功能，为以后的优化保留
    
End Sub

Private Sub Command2_Click()

    Dim l1%, j%
    l1 = List1.ListIndex
    For j = 1 To l1
        Call MyLog(Time & "删除第" & j & "记录")
    Next j
    List1.RemoveItem j

End Sub

Private Sub Command3_Click()

    Call MyLog(Time & "恢复第" & j & "记录")
    Me.Hide
End Sub

Private Sub Command4_Click()

    If Form2.Width = 5895 Then
        Form2.Width = 9895
        Command4.Caption = "<<"
    Else
        Form2.Width = 5895
        Command4.Caption = "关于作者"
    End If
End Sub

Private Sub Command9_Click()
    
    Shell "explorer.exe " & Form3.Text1.Text

End Sub

Private Sub Form_Load()

    Call MyLog(Time & "打开“历史记录”")
    Form2.Width = 5895
    'Form.StartUpPosition = 8000
    Form2.Left = 17800: Form2.Top = 5400
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

Call MyLog(Time & "退出“历史记录”")

End Sub

Private Sub Picture2_Click()

'Dim RetVal
'RetVal = Shell("C:\Program Files\internet explorer\iexplore.exe" & " " & "http://wpa.qq.com/msgrd?v=3&uin=2080979987&site=qq&menu=yes", vbNormalFocus)
    Form3.Show
    Form3.Text1.Text = "http://wpa.qq.com/msgrd?v=3&uin=2080979987&site=qq&menu=yes"
    Form3.WebBrowser1.Navigate "http://wpa.qq.com/msgrd?v=3&uin=2080979987&site=qq&menu=yes"
'Shell "explorer.exe " & "http://wpa.qq.com/msgrd?v=3&uin=2080979987&site=qq&menu=yes"

End Sub

Private Sub MyLog(nStr As String)

Dim F As String, H As Long
    nStr = Format(Now, "yyyy-mm-dd hh:mm:ss") & nStr
    F = App.Path & "\log.log" '将日志保存到exe相同的文件夹
    H = FreeFile
Open F For Append As #H
Print #H, nStr
Close #H

End Sub
