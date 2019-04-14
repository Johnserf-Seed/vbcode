VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "动画彩蛋"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   Icon            =   "MovieEgg.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "MovieEgg.frx":424A
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'动画彩蛋项目，代码借鉴贴吧
Option Explicit
Private WithEvents Timer1 As Timer
Attribute Timer1.VB_VarHelpID = -1
Dim i%, j%, x1%, y1%, blockw%, blockh%, carX%, carY%, pcolor$
Dim N%, L%, C$
Const Captions As String = "SG的小汽车"

Private Sub Form_Load()

    Me.AutoRedraw = True
    Me.DrawWidth = 2
    Me.Width = 5120
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
    Me.Cls
    carY = Me.Height \ 2: blockw = 500: blockh = 200
    x1 = 0: y1 = carY - 230
    Set Timer1 = Controls.Add("vb.timer", "timer1")
    Timer1.Interval = 50
    
End Sub

Private Sub Timer1_Timer()

    Me.Cls
    For i = 1 To 12
        pcolor = IIf(i Mod 2 = 0, vbBlue, vbRed)
    Line (carX - j, carY)-(carX - j + blockw, carY + blockh), pcolor, BF
        carX = IIf(carX + 500 >= 6000, 0, carX + 500)
    Next i
        j = IIf(j + 100 > 900, 0, j + 100)
            Line (x1, y1)-(x1 + 500, y1 + 100), , B
            Me.Circle (x1 + 100, y1 + 150), 50
            Me.Circle (x1 + 380, y1 + 150), 50
            x1 = IIf(x1 + 50 >= 5000, -500, x1 + 50)
            L = Int(Me.Width / 220)
            C = String(L, " ") & Captions & String(L, " ")
            N = N + 1
        If N > Len(C) - L Then N = 1
        Me.Caption = Mid(C, N, L)
    
End Sub


