VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form4 
   Caption         =   "��ѹ��ѹ��"
   ClientHeight    =   2115
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4155
   ForeColor       =   &H00C0C0C0&
   Icon            =   "SGcompress.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   4155
   StartUpPosition =   1  '����������
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1800
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "΢���ź� Light"
         Size            =   12
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Text            =   "·��"
      Top             =   1080
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��ѹ�ļ�"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ѹ���ļ�"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mystr As String
Dim Source As String ' Դ�ļ�
Dim Target As String ' Ŀ���ļ�
Dim RetVal
Private Sub Command1_Click() '=========ѹ���ļ�

    mystr = "C:\Program Files\WinRAR\WinRAR.exe" 'winrar.exe�ļ�·��
    Source = App.Path & "\ico\SGmp3.ico"
    Target = App.Path & "\ico.rar" 'ѹ����ʽ������rar,Ҳ������cab....
    mystr = mystr & " a " & Target & " " & Source '�����ַ���
    RetVal = Shell(mystr)
    
End Sub
Private Sub Command2_Click() '===========��ѹ�ļ�

    mystr = "C:\Program Files\WinRAR\WinRAR.exe"
    Source = App.Path & "\ico.rar"
    Target = App.Path & "\new" '���ѹ���ļ���λ��
    mystr = mystr & "X" & Source & " " & Target
    Text1.Text = mystr
    RetVal = Shell(mystr)
    
End Sub

Private Sub Form_Load()

    Text1.ForeColor = &HC0C0C0
    
End Sub

Private Sub Text1_Click()

    Text1.ForeColor = vbBlack
    Text1.Text = ""

End Sub
