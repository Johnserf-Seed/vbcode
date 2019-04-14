VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form3 
   Caption         =   "SGe"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20265
   Icon            =   "SGe.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   10035
   ScaleWidth      =   20265
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command5 
      Caption         =   "刷新"
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "主页"
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "前进"
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "后退"
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "跳转"
      Height          =   615
      Left            =   8160
      TabIndex        =   3
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8175
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   20295
      ExtentX         =   35798
      ExtentY         =   15901
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
      Location        =   ""
   End
   Begin VB.Label Label1 
      Height          =   1215
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   16695
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim zy As String
Private Sub Command1_Click()

    WebBrowser1.Navigate Text1.Text

End Sub

Private Sub Command2_Click()

    WebBrowser1.GoBack

End Sub

Private Sub Command3_Click()

    WebBrowser1.GoForward

End Sub

Private Sub Command4_Click()

    zy = "https://17shiyan2.cn"
    WebBrowser1.Navigate (zy)

End Sub

Private Sub Command5_Click()

    WebBrowser1.Refresh

End Sub

Private Sub Form_Resize()

    'WebBrowser1.Move 10, 10, Me.ScaleWidth - 20, Me.ScaleHeight - 20
    On Error Resume Next
    Label1.Left = Me.Width - Label1.Width - 240
    Label1.Top = Me.Height - Label1.Height - 720
    WebBrowser1.Width = Me.Width - 240
    WebBrowser1.Height = Me.Height - 1200

End Sub

Private Sub Form_Load()

    ScriptErrorSuppress = True
    Text1.Text = "http://www.baidu.com"
    WebBrowser1.Navigate Text1.Text

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And Text1 <> "" Then
        Command1_Click
    End If
    
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)

    Text1.Text = WebBrowser1.LocationURL

End Sub

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)

    Dim frm As Form1
    Set frm = New Form1
    frm.Visible = TrueSet
    Set ppDisp = Me.WebBrowser1.object
    WebBrowser1.Height = ScaleHeight
    WebBrowser1.Width = ScaleWidth
    Me.Width = WebBrowser1.Width

End Sub

Private Sub WebBrowser1_DownloadBegin()

    WebBrowser1.Silent = True

End Sub
Private Sub WebBrowser1_DownloadComplete()

    WebBrowser1.Silent = True

End Sub

