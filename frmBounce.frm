VERSION 5.00
Begin VB.Form frmBounce 
   BorderStyle     =   0  'None
   Caption         =   "Animated Mouse's Tail"
   ClientHeight    =   3435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   229
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   488
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "About"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   2520
      Top             =   1440
   End
   Begin VB.Label Label1 
      Caption         =   "Double Click to Exit"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "frmBounce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
    frmAbout.Show vbModal
End Sub

Private Sub Form_DblClick()
    Dim i As Integer
    'For i = ImgBall.LBound + 1 To ImgBall.UBound
    '    Unload ImgBall(i)
    'Next i
    Unload Me
    End
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    For i = 1 To 7
        Load ImgBall(i)
        ImgBall(i).Visible = True
        Dim pos As POINTAPI
        GetCursorPos pos
        
        ImgBall(i).Top = pos.X '.ScaleY(ImgBall(i - 1).Top + 11, 3, 1)
        ImgBall(i).Left = pos.Y
    Next i
    Call InitVal(7)
    Call InitBall
    
    Timer1.Interval = 20
    Timer1.Enabled = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'MoveHandler CLng(x), CLng(y)
    'Animate
End Sub

Private Sub Timer1_Timer()

    MoveHandler

    Animate
    DoEvents
End Sub
