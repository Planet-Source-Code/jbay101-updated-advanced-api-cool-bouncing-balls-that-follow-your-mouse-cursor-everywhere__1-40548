Attribute VB_Name = "modBounce"
'
'
'
Option Explicit

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Type POINTAPI
        X As Long
        Y As Long
End Type

Global ImgBall(0 To 7) As New frmImgBall

Public Type Vec2D
    X As Long
    Y As Long
End Type

Public Type AnimBall
    Vec As Vec2D
    dx As Double
    dy As Double
    Img As Object
End Type

Dim nBalls As Integer
Dim Xpos, Ypos
Dim DeltaT As Double
Dim SegLen
Dim SpringK
Dim Mass
Dim Gravity
Dim Resistance
Dim StopVel As Double
Dim StopAcc As Double
Dim DotSize As Long
Dim Bounce As Double
Dim bFollowM As Boolean
Dim Balls() As AnimBall
Private Const SWP_NOOWNERZORDER = &H200
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SWP_NOMOVE = &H2
'*******************************************************
'
'This module is all you need to start making your
'own Image Shaped Forms!
'
'*******************************************************

'General Api Declarations
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private hRgn As Long

Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

'This the Main Code to make an Image Shaped Form
'What it does is scan the Image passed to it and then
'remove all lines that correspond to the Transparent
'Color, creating a new virtual image, but without a
'particular color

Public Function GetBitmapRegion(cPicture As StdPicture, cTransparent As Long)
'Variable Declaration
    Dim hRgn As Long, tRgn As Long
    Dim X As Integer, Y As Integer, X0 As Integer
    Dim hDC As Long, BM As BITMAP
'Create a new memory DC, where we will scan the picture
    hDC = CreateCompatibleDC(0)
    If hDC Then
'Let the new DC select the Picture
        SelectObject hDC, cPicture
'Get the Picture dimensions and create a new rectangular
'region
        GetObject cPicture, Len(BM), BM
        hRgn = CreateRectRgn(0, 0, BM.bmWidth, BM.bmHeight)
'Start scanning the picture from top to bottom
        For Y = 0 To BM.bmHeight
            For X = 0 To BM.bmWidth
'Scan a line of non transparent pixels
                While X <= BM.bmWidth And GetPixel(hDC, X, Y) <> cTransparent
                    X = X + 1
                Wend
'Mark the start of a line of transparent pixels
                X0 = X
'Scan a line of transparent pixels
                While X <= BM.bmWidth And GetPixel(hDC, X, Y) = cTransparent
                    X = X + 1
                Wend
'Create a new Region that corresponds to the row of
'Transparent pixels and then remove it from the main
'Region
                If X0 < X Then
                    tRgn = CreateRectRgn(X0, Y, X, Y + 1)
                    CombineRgn hRgn, hRgn, tRgn, 4
'Free the memory used by the new temporary Region
                    DeleteObject tRgn
                End If
            Next X
        Next Y
'Return the memory address to the shaped region
        GetBitmapRegion = hRgn
'Free memory by deleting the Picture
        DeleteObject SelectObject(hDC, cPicture)
    End If
'Free memory by deleting the created DC
    DeleteDC hDC
End Function

Public Sub StayOnTop(frmForm As Form, fOnTop As Boolean)
    
    
    Dim lState As Long
    Dim iLeft As Integer, iTop As Integer, iWidth As Integer, iHeight As Integer


    With frmForm
        iLeft = .Left / Screen.TwipsPerPixelX
        iTop = .Top / Screen.TwipsPerPixelY
        iWidth = .Width / Screen.TwipsPerPixelX
        iHeight = .Height / Screen.TwipsPerPixelY
    End With
    


    If fOnTop Then
        lState = HWND_TOPMOST
    Else
        lState = HWND_NOTOPMOST
    End If
    Call SetWindowPos(frmForm.hwnd, lState, iLeft, iTop, iWidth, iHeight, 0)
End Sub

Public Sub apiMove(frmForm As Form, Left As Long, Top As Long)
    
    
    Dim lState As Long
    Dim iLeft As Integer, iTop As Integer, iWidth As Integer, iHeight As Integer


    With frmForm
        iLeft = Left
        iTop = Top
        iWidth = .Width / Screen.TwipsPerPixelX
        iHeight = .Height / Screen.TwipsPerPixelY
    End With
    


    'If fOnTop Then
        lState = HWND_TOPMOST
   ' Else
   '     lState = HWND_NOTOPMOST
   ' End If
    Call SetWindowPos(frmForm.hwnd, lState, iLeft, iTop, iWidth, iHeight, 0)
End Sub

Sub SetRegion(frmForm As Form, Color As Long)
'Free the memory allocated by the previous Region
    If hRgn Then DeleteObject hRgn
'Scan the Bitmap and remove all transparent pixels from
'it, creating a new region
    hRgn = GetBitmapRegion(frmForm.Picture, Color)
'Set the Forms new Region
    SetWindowRgn frmForm.hwnd, hRgn, True
End Sub

Function InitVal(Balls As Integer)
' Some of the variables are still unknown to me
    nBalls = Balls          ' numbers of ball
    Xpos = Ypos = 0     ' evaluate position
    DeltaT = 0.01       '
    SegLen = 10#        ' it seem like the distance between the
                        ' mouse pointer and the ball
                        ' it's quite intersting to change the value
                        ' and see the effect
    SpringK = 11       ' spring constant,
                       ' if large, the longer and higher the tail
                       ' will swing
    Mass = 1            'mass of the ball
    Gravity = 40        ' gravity coeff,
                        ' if large, the balls are more difficult
                        ' to move upward
    Resistance = 9     ' resistivity of the ball to move itself
                        ' from a location, the larger the more difficult to
                        ' move
    StopVel = 0.1
    StopAcc = 0.1
    DotSize = 11        ' the size of the ball in pixel
    Bounce = 0.95       ' bouncing coeff,
    bFollowM = True     ' animation flag
End Function


' must only be called after load all imgBall
Function InitBall()
    Dim i As Integer
    ReDim Balls(nBalls)

    For i = 0 To nBalls
        Balls(i) = BallSet(ImgBall(i))
    Next i

    For i = 0 To nBalls
        Dim pos As POINTAPI
        GetCursorPos pos
        apiMove Balls(i).Img, pos.X, pos.Y
        'Balls(i).Img.Left = pos.X 'frmImgBall.ScaleX(balls(i).Vec.x, 3, 1)
        'Balls(i).Img.Top = pos.Y 'frmImgBall.ScaleY(balls(i).Vec.y, 3, 1)
    Next i
End Function

' initialize a ball
Function BallSet(Img As Object) As AnimBall
     Dim pos As POINTAPI
    GetCursorPos pos
    
    BallSet.Vec.X = pos.X
    BallSet.Vec.Y = pos.Y
    BallSet.dx = BallSet.dy = 0
    Set BallSet.Img = Img
End Function

' initialize a vector variable
Function VecSet(X As Long, Y As Long) As Vec2D
    VecSet.X = X
    VecSet.Y = Y
End Function

' update position when mouse move
Function MoveHandler()
Dim pos As POINTAPI
GetCursorPos pos

    Xpos = pos.X
    Ypos = pos.Y
End Function

' calculate the spring force of the balls chain
Function SpringForce(i As Integer, j As Integer, ByRef spring As Vec2D)
    Dim tempdx, tempdy, tempLen, springF
    tempdx = Balls(i).Vec.X - Balls(j).Vec.X
    tempdy = Balls(i).Vec.Y - Balls(j).Vec.Y
    tempLen = Sqr(tempdx * tempdx + tempdy * tempdy)
    If (tempLen > SegLen) Then
        springF = SpringK * (tempLen - SegLen)
        spring.X = spring.X + (tempdx / tempLen) * springF
        spring.Y = spring.Y + (tempdy / tempLen) * springF
    End If
End Function

' main routine of this animated balls
' call on mouse move or every 20ms
Function Animate()
    Dim iH, iW
    Dim start As Integer
    Dim i As Integer
    Dim spring As Vec2D
    Dim resist As Vec2D
    Dim accel As Vec2D
    ' enable the animation
    If (bFollowM) Then
        Balls(0).Vec.X = Xpos
        Balls(0).Vec.Y = Ypos
        start = 1
    End If
    
    For i = start To nBalls
        spring = VecSet(0, 0)
        
        If (i > 0) Then
            Call SpringForce(i - 1, i, spring)
        End If
        
        If (i < (nBalls - 1)) Then
            Call SpringForce(i + 1, i, spring)
        End If
        resist = VecSet(-Balls(i).dx * Resistance, -Balls(i).dy * Resistance)
        accel = VecSet((spring.X + resist.X) / Mass, _
                        (spring.Y + resist.Y) / Mass + Gravity)

        Balls(i).dx = Balls(i).dx + DeltaT * accel.X
        Balls(i).dy = Balls(i).dy + DeltaT * accel.Y

        If (Abs(Balls(i).dx) < StopVel And _
            Abs(Balls(i).dy) < StopVel And _
            Abs(accel.X) < StopAcc And _
            Abs(accel.Y) < StopAcc) Then
            Balls(i).dx = 0
            Balls(i).dy = 0
        End If

        Balls(i).Vec.X = Balls(i).Vec.X + Balls(i).dx
        Balls(i).Vec.Y = Balls(i).Vec.Y + Balls(i).dy

        ' checking for boundary conditions
        iW = frmImgBall.ScaleX(Screen.Width, 1, 3) ' frmBounce.ScaleWidth
        iH = frmImgBall.ScaleY(Screen.Height, 1, 3)

        ' check bottom
        If (Balls(i).Vec.Y >= iH - DotSize - 1) Then
            If (Balls(i).dy > 0) Then
                Balls(i).dy = Bounce * (-Balls(i).dy)
            End If
            Balls(i).Vec.Y = iH - DotSize - 1
        End If
        
        ' check right
        If (Balls(i).Vec.X >= iW - DotSize) Then
            If (Balls(i).dx > 0) Then
                Balls(i).dx = Bounce * (-Balls(i).dx)
            End If
            Balls(i).Vec.X = iW - DotSize - 1
        End If

        ' check left
        If (Balls(i).Vec.X < 0) Then
            If (Balls(i).dx < 0) Then
                Balls(i).dx = Bounce * (-Balls(i).dx)
            End If
            Balls(i).Vec.X = 0
        End If
        ' check top
        If (Balls(i).Vec.Y < 0) Then
            If (Balls(i).dy < 0) Then
                Balls(i).dy = Bounce * (-Balls(i).dy)
            End If
            Balls(i).Vec.Y = 0
        End If

        apiMove Balls(i).Img, Balls(i).Vec.X, Balls(i).Vec.Y
        UpdateWindow Balls(i).Img.hwnd
        Balls(i).Img.Refresh
        'Balls(i).Img.Left = frmImgBall.ScaleX(Balls(i).Vec.X, 3, 1)
        'Balls(i).Img.Top = frmImgBall.ScaleY(Balls(i).Vec.Y, 3, 1)
    Next i
    
End Function
