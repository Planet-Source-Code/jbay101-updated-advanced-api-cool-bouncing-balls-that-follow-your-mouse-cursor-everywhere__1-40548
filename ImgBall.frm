VERSION 5.00
Begin VB.Form frmImgBall 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   165
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "ImgBall.frx":0000
   ScaleHeight     =   11
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   11
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   0
      Picture         =   "ImgBall.frx":01CE
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   0
      Top             =   0
      Width           =   165
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9000
      Top             =   75
   End
End
Attribute VB_Name = "frmImgBall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

SetRegion Me, vbWhite

'    Dim lResult As Long
'    lResult = SetWindowPos(Me.hWnd, HWND_NOTOPMOST, _
'    0, 0, 0, 0, FLAGS)
StayOnTop Me, True
End Sub

