VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���ֲ�����"
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
   StartUpPosition =   2  '��Ļ����
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
      Caption         =   "���๤�ߣ����ԣ�"
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
      Caption         =   "ѭ������"
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
      Caption         =   "ɾ��"
      Height          =   495
      Left            =   5640
      TabIndex        =   6
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��ʷ��¼"
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
      Caption         =   "�������"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ѡ�����Ŀ¼"
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
      Caption         =   "�ײ�ʱ������"
      Height          =   615
      Left            =   5640
      TabIndex        =   14
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label6 
      BeginProperty Font 
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
      Caption         =   "����ѭ������"
      Height          =   615
      Left            =   5520
      TabIndex        =   9
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "�����������"
      Height          =   735
      Left            =   5640
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "�����б�"
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
    Label6.Caption = "ѭ�����ſ���"
Else
    Timer2.Enabled = False
    Label6.Caption = "ѭ�����Źر�"
End If

End Sub

Private Sub Command1_Click()
Dim a As String, k As Integer, pa As Integer, i As Integer
    MsgBox "�����Ӹ����������Ļ������б��л��һ���հ�ѡ��ݲ�����ɾ��Ŷ��", vbInformation, "��ܰ����"
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
            If k = " " Then '���׸�ķָ���
                List1.AddItem a
                a = ""
            ElseIf k = "\" Then 'һ�׸�ʱ������·���ķָ���
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

Private Sub Command4_Click() '���⣺ֻ��ɾ��ѡ�еĵ�һ�ѡ���������޷�ɾ����ȫ��ɾ�����ж��Ƿ�ѡ����뱨����Ч������������
    'For i = List1.ListCount - 1 To 0 Step -1
    
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) = False Then '�˴��ж��Ƿ�ѡ��   '����
            MsgBox "��ѡ������Ҫɾ���ĸ���", , "��ܰ����"
            Exit Sub
        End If
        
        If MsgBox("��ȷ��Ҫɾ��ô��", 64 + vbYesNo, "��ʾ") = vbYes Then
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

    MsgBox "��ϲ���ҵ��˲ʵ�����", vbInformation
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
        MsgBox "�����Ѿ�������", , "MusicPlayer"
        Me.SetFocus
        Cancel = 0
        
End Sub

Private Sub List1_DblClick()
    Me.Caption = "�ҵĲ����� ��ǰ����" & pa & List1.Text '��ʾ�ڴ������
    WindowsMediaPlayer1.URL = pa & List1.Text
End Sub

Private Sub Form_Resize()

'If Me.WindowState <> 1 Then '�жϴ����С
'Command1.Width = Form1.Width / 10 '��ť����Ǵ����1/5
'Command1.Height = Form1.Height / 10 '��ť�߶��Ǵ����1/5

End Sub

Private Sub Form_Load()

    Call MyLog(Time & "����������")
    
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

    Call MyLog(Time & "�˳�������")
    'Cancel = (MsgBox("��ӭʹ��", vbOKCancel <> vbOK))
    Dim iAnswer As Integer
    iAnswer = MsgBox("��ӭʹ��   ���Ҫ�˳���", vbYesNo, "�Ͽ�(=��أ�=)")
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

    If Me.WindowsMediaPlayer1.playState = 1 Then '1Ϊֹͣ����
        Me.WindowsMediaPlayer1.URL = pa & List1.Text
        Me.WindowsMediaPlayer1.Controls.play
    End If
    
End Sub

Private Sub Timer3_Timer()

    Randomize
    Label4.Caption = "�����ǣ�" & Date & "�����ǣ�" & Format(Now, "AMPM(hh:mm:ss)")
    Label5.Caption = "���ߣ�SG"
    Label5.ForeColor = RGB(Int(Rnd * (256)), Int(Rnd * (256)), Int(Rnd * (256)))

End Sub

Private Sub Timer4_Timer()

    Label4.Left = Label4.Left - 50 'ÿ�α�ǩ�����˶�50����
        If Label4.Left < -Label4.Width Then   '�����ǩ�Ƴ��������
           Label4.Left = Me.ScaleWidth - 975     '��ǩ�Ӵ����ұ߽���
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
        F = App.Path & "\log.log" '����־���浽exe��ͬ���ļ���
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

    MsgBox "��Ϊ�������������ģ�����bug��ǳ��࣬ϣ����λ��ţ��Ҫ�������ۣ���������ʷ��¼�е���������н��з����������ڿ������Ƴ���������ʲô���ܣ��Ժ�Ӧ�û���һ��������", vbInformation, "V1.0.6��ª�ĳ���"
    MsgBox "�����ҪԴ���ͬѧ���һ��ں�����ֲ���ڳ����У���λ���Բο���ָ�������һ��", vbInformation, "V1.0.6��ª�ĳ���"
    
End Sub

Private Sub Timer7_Timer()

    Timer6.Enabled = False
    
End Sub
