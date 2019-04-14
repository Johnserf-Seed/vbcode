VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "音乐播放器"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6525
   Icon            =   "musicplayer.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   6525
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer7 
      Interval        =   2500
      Left            =   2160
      Top             =   120
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   3360
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   2640
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer5 
      Interval        =   3000
      Left            =   6000
      Top             =   3480
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   90
      Left            =   5040
      Top             =   4440
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   6000
      Top             =   3000
   End
   Begin VB.CommandButton Command5 
      Caption         =   "更多工具（测试）"
      Height          =   495
      Left            =   4440
      TabIndex        =   10
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   6000
      Top             =   2280
   End
   Begin VB.CheckBox Check1 
      Caption         =   "循环播放"
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6000
      Top             =   1440
   End
   Begin VB.CommandButton Command4 
      Caption         =   "删除"
      Height          =   495
      Left            =   5640
      TabIndex        =   6
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "历史记录"
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      Top             =   3960
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5640
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "随机播放"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "选择歌曲目录"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   3000
      ItemData        =   "musicplayer.frx":424A
      Left            =   3840
      List            =   "musicplayer.frx":424C
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "底部时间作者"
      Height          =   615
      Left            =   5640
      TabIndex        =   14
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label6 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   13
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   12
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   4560
      Width           =   4935
   End
   Begin VB.Label Label3 
      Caption         =   "控制循环播放"
      Height          =   615
      Left            =   5520
      TabIndex        =   9
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "控制随机播放"
      Height          =   735
      Left            =   5640
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "歌曲列表："
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   6165
      _cy             =   6376
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pa As String
Dim lngErrorNumber As Long, strErrorDescription As String, strProcedure As String
'Dim C as New
Private Sub Check1_Click()

If Check1.Value = 1 Then
    Timer2.Enabled = True
    Label6.Caption = "循环播放开启"
Else
    Timer2.Enabled = False
    Label6.Caption = "循环播放关闭"
End If

End Sub

Private Sub Command1_Click()
Dim a As String, k As Integer, pa As Integer, i As Integer
    MsgBox "如果添加歌曲是批量的话歌曲列表中会多一个空白选项，暂不可以删除哦！", vbInformation, "温馨提醒"
    Dim song As String
        With CommonDialog1
            .FileName = ""
            .Filter = "*.MP3|*.mp3;*.avi;*.mid"
            .Flags = 512
            .ShowOpen
            song = .FileName
        End With
    If song = "" Then Exit Sub
        a = ""
        For i = Len(song) To 1 Step -1
        k = Mid(song, i, 1)
            If k = " " Then '多首歌的分隔符
                List1.AddItem a
                a = ""
            ElseIf k = "\" Then '一首歌时歌名与路径的分隔符
                List1.AddItem a
                pa = Left(song, i)
        Exit Sub
        Else
        a = k & a
    End If
    Next i
    
End Sub

Private Sub Command2_Click()

    Timer1.Enabled = True
    'sj = Int(Rnd * Val((List1.ListCount - 1)) + Val(List1.ListCount))
    'WindowsMediaPlayer1.URL = pa & List1.List(sj)
    
End Sub

Private Sub Command3_Click()

    Form2.Show
    
End Sub

Private Sub Command4_Click() '问题：只能删除选中的第一项，选择其他项无法删除，全部删除后判断是否选择代码报错：无效属性数组索引
    'For i = List1.ListCount - 1 To 0 Step -1
    
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) = False Then '此处判断是否选中   '报错
            MsgBox "请选中你想要删除的歌曲", , "温馨提醒"
            Exit Sub
        End If
        
        If MsgBox("你确认要删除么？", 64 + vbYesNo, "提示") = vbYes Then
        Do While i <= List1.ListCount - 1
            If List1.Selected(i) = True Then
                List1.RemoveItem i
                i = i - 1
            End If
            i = i + 1
        Loop
    End If
    Next i

End Sub

Private Sub Command5_Click()

    Form5.Show
    
End Sub

Private Sub Form_DblClick()

    MsgBox "恭喜你找到了彩蛋！！", vbInformation
    Form7.Show
    
End Sub

Private Sub Form_Initialize()

    If App.PrevInstance Then
        Picture1.LinkTopic = "musicplayer|Form1"
        Picture1.LinkMode = 2
        Picture1.LinkExecute ""
        End
    End If
    
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)

    If Me.WindowState = 1 Then Me.WindowState = 0
        MsgBox "程序已经在运行", , "MusicPlayer"
        Me.SetFocus
        Cancel = 0
        
End Sub

Private Sub List1_DblClick()
    Me.Caption = "我的播放器 当前播放" & pa & List1.Text '显示在窗体标题
    WindowsMediaPlayer1.URL = pa & List1.Text
End Sub

Private Sub Form_Resize()

'If Me.WindowState <> 1 Then '判断窗体大小
'Command1.Width = Form1.Width / 10 '按钮宽度是窗体的1/5
'Command1.Height = Form1.Height / 10 '按钮高度是窗体的1/5

End Sub

Private Sub Form_Load()

    Call MyLog(Time & "启动主程序")
    
    Timer6.Enabled = True
    Timer4.Enabled = True
    Dim Path As String, Item As Long, Temp As String, i As Long
        Path = IIf(Len(App.Path) = 3, App.Path, App.Path) & "log\ " & "data.dat"
        If Dir(Path, vbDirectory) = "" Then Exit Sub
        Open Path For Input As #1
        Line Input #1, Temp
        Item = Val(Temp)
            For i = 1 To Item
                Line Input #1, Temp
                List1.AddItem Temp
            Next i
        Close #1

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call MyLog(Time & "退出主程序")
    'Cancel = (MsgBox("欢迎使用", vbOKCancel <> vbOK))
    Dim iAnswer As Integer
    iAnswer = MsgBox("欢迎使用   真的要退出吗？", vbYesNo, "诚恳(=′ω｀=)")
    If iAnswer = vbNo Then
          Cancel = True
    Else
    Dim i As Long, Temp As String
        For i = 0 To List1.ListCount - 1
            If Temp = "" Then
                Temp = List1.List(i)
            Else
                Temp = Temp & vbNewLine & List1.List(i)
            End If
        Next i
        
    Open IIf(Len(App.Path) = 3, App.Path, App.Path & "\") & "log\" & "data.dat" For Output As #1
    Print #1, List1.ListCount & vbNewLine & Temp
    Close #1
    End
    End If
  
End Sub

Private Sub Timer1_Timer()

    If WindowsMediaPlayer1.playState = wmppsStopped Or WindowsMediaPlayer1.playState = wmppsReady Then
        SongNo = IIf(SongNo > 3, 1, SongNo + 1)
    If Dir(List1.List(SongNo - 1)) = "" Then
    Exit Sub
        WindowsMediaPlayer1.URL = List1.List(SongNo - 1)
        WindowsMediaPlayer1.Controls.play
    End If
    End If
    
End Sub

Private Sub Timer2_Timer()

    If Me.WindowsMediaPlayer1.playState = 1 Then '1为停止播放
        Me.WindowsMediaPlayer1.URL = pa & List1.Text
        Me.WindowsMediaPlayer1.Controls.play
    End If
    
End Sub

Private Sub Timer3_Timer()

    Randomize
    Label4.Caption = "今天是：" & Date & "现在是：" & Format(Now, "AMPM(hh:mm:ss)")
    Label5.Caption = "作者：SG"
    Label5.ForeColor = RGB(Int(Rnd * (256)), Int(Rnd * (256)), Int(Rnd * (256)))

End Sub

Private Sub Timer4_Timer()

    Label4.Left = Label4.Left - 50 '每次标签向左运动50像素
        If Label4.Left < -Label4.Width Then   '如果标签移出窗体左边
           Label4.Left = Me.ScaleWidth - 975     '标签从窗体右边进入
        End If
        
End Sub

Private Sub Timer5_Timer()

    If Check1.Value = 1 Then xh = xh + 1
        If xh > 0 Then
        xh = xh - 1
    Else
        Label6.Caption = ""
    End If
    
End Sub

Private Sub MyLog(nStr As String)

    Dim F As String, H As Long
        nStr = Format(Now, "yyyy-mm-dd hh:mm:ss") & nStr
        F = App.Path & "\log.log" '将日志保存到exe相同的文件夹
        H = FreeFile
    Open F For Append As #H
    Print #H, nStr
    Print #H, "Error: " & lngErrorNumber
    Print #H, "Description: " & strErrorDescription
    Print #H, "Module: " & strModule
    Print #H, "Procedure: " & strProcedure
    Close #H

End Sub

Private Sub Timer6_Timer()

    MsgBox "因为是随便做着玩玩的，所以bug会非常多，希望各位大牛不要吝啬评价，可以在历史记录中的意见反馈中进行反馈，现在在考虑这破程序可以添加什么功能，以后应该会变成一个工具箱", vbInformation, "V1.0.6简陋的程序"
    MsgBox "如果需要源码的同学，我会在后续中植入在程序中，各位可以参考，指正，借鉴一下", vbInformation, "V1.0.6简陋的程序"
    
End Sub

Private Sub Timer7_Timer()

    Timer6.Enabled = False
    
End Sub
