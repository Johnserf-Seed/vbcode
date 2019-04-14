VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "简易照片查看器"
   ClientHeight    =   11445
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20280
   Icon            =   "PictureView.frx":0000
   LinkTopic       =   "Form6"
   ScaleHeight     =   11445
   ScaleWidth      =   20280
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      Height          =   495
      Left            =   12720
      TabIndex        =   4
      Top             =   10200
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   495
      Left            =   11400
      TabIndex        =   3
      Top             =   10200
      Width           =   975
   End
   Begin VB.FileListBox File1 
      Height          =   6930
      Left            =   2520
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Height          =   6390
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   9735
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   240
      Width           =   14895
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mpath As String: Dim mfile As String, N%

Private Sub Command1_Click()

    mpath = File1.Path
    mfile = File1.FileName
    Image1.Picture = LoadPicture(mfile & CStr(N))
    If N <= 5 Then '如果有5张图片
        N = N + 1
    Else
        N = 1
    End If
    
End Sub

Private Sub Dir1_Change()

    File1.Path = Dir1.Path
    
End Sub

Private Sub Drive1_Change()

    Dir1.Path = Drive1.Drive
    
End Sub

Private Sub File1_Click()

    On Error GoTo errDeal '不推荐用goto
    mpath = File1.Path
    mfile = File1.FileName
    If Right(mpath, 1) = "\" Then
        mfile = mpath + mfile
    Else
        mfile = mpath + "\" + mfile
    End If
        Image1.Picture = LoadPicture(mfile)
    Exit Sub
errDeal:
        Image1.Picture = LoadPicture("")
        MsgBox "图片类型无效", vbYesNo, "提示！"
        
End Sub

Private Sub Form_Load()

    MsgBox "简易的照片查看器，还请各位老爷过目", vbInformation, "V1.0.0"
    N = 1
    
End Sub

