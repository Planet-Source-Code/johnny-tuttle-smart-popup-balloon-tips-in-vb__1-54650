VERSION 5.00
Begin VB.Form frmBalloon 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E1FFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   1830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   ControlBox      =   0   'False
   ForeColor       =   &H80000017&
   Icon            =   "frmBalloon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   122
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   321
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timAutoClose 
      Enabled         =   0   'False
      Left            =   3960
      Top             =   1200
   End
   Begin VB.Image imgIconXP 
      Height          =   240
      Index           =   2
      Left            =   720
      Picture         =   "frmBalloon.frx":000C
      Top             =   1320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIconXP 
      Height          =   240
      Index           =   1
      Left            =   480
      Picture         =   "frmBalloon.frx":0596
      Top             =   1320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIconXP 
      Height          =   240
      Index           =   0
      Left            =   240
      Picture         =   "frmBalloon.frx":0B20
      Top             =   1320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgDisplayIcon 
      Height          =   240
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   240
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "<Title>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   645
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "<Caption>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   795
   End
End
Attribute VB_Name = "frmBalloon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I trimmed off and rewrote almost all of this code.
'But my original inspiration goes to Robert Morris
'for his PopUp Balloons example. You can find it on planet source code
'by searching for "Popup Balloons (2k/XP-style)"
'
'Basically this determines the size and position of the form based on
'the text, and then turns the form into a rounded rectangle with a
'Balloon tip point (Yes I wrote that myself)
'Balloon tip position is based on where on the screen it will be displayed
'
'Sorry this is all the comments, they distract me when I'm programming.

Option Explicit

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type PointType
    Point1 As POINTAPI
    Point2 As POINTAPI
    Point3 As POINTAPI
End Type

Private Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type

Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As PointType, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function FrameRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long

Private Const HWND_TOP = 0
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
   
Private bBottom As Boolean
Private bLeft As Boolean


Friend Sub ShowBalloon(sTitle As String, sText As String, X As Long, Y As Long, Optional IStyle As IconStyle, _
                       Optional iAutoCloseAfter As Integer = 0)

    Dim lHeight&, lWidth&, lTop&, lLeft&
    
    Me.BackColor = vbInfoBackground
    lblTitle.ForeColor = vbInfoText
    lblText.ForeColor = vbInfoText
    
    lblTitle.Caption = sTitle
    lblText.Caption = sText
    
    If lblText.Width + 20 > ScaleX(Screen.Width / 2, vbTwips, vbPixels) Then
        lblText.WordWrap = True
        lblText.Width = ScaleX(Screen.Width / 2, vbTwips, vbPixels) - 20
    End If
    
    If (lblTitle.Width + 20) > lblText.Width Then
        lWidth = 40 + lblTitle.Width
    Else
        lWidth = 20 + lblText.Width
    End If
    lHeight = 45 + lblText.Height

    Select Case IStyle
        Case ISTYLE_INFO: imgDisplayIcon.Picture = imgIconXP(0).Picture
        Case ISTYLE_ERROR: imgDisplayIcon.Picture = imgIconXP(1).Picture
        Case ISTYLE_WARNING: imgDisplayIcon.Picture = imgIconXP(2).Picture
        Case Else
            Me.imgDisplayIcon.Visible = False
            Me.lblTitle.Left = imgDisplayIcon.Left
    End Select
        
    lHeight = lHeight + 20
    
    If (Y - lHeight) > 0 Then
        bBottom = True
        lTop = Y - lHeight
    Else
        bBottom = False
        imgDisplayIcon.Top = 28
        lblTitle.Top = 28
        lblText.Top = 52
        lTop = Y
    End If
    If (X + lWidth) < ScaleX(Screen.Width, vbTwips, vbPixels) Then
        bLeft = True
        If X > 15 Then
            lLeft = X - 15
        Else
            lLeft = 0
        End If
    Else
        bLeft = False
        If X < (ScaleX(Screen.Width, vbTwips, vbPixels) - 15) Then
            lLeft = X - lWidth + 15
        Else
            lLeft = ScaleX(Screen.Width, vbTwips, vbPixels) - lWidth
        End If
    End If
    
    If iAutoCloseAfter = 0 Then
        Me.timAutoClose.Enabled = False
    Else
        Me.timAutoClose.Interval = iAutoCloseAfter
        Me.timAutoClose.Enabled = True
    End If
    
    SetWindowPos Me.hwnd, HWND_TOP, lLeft, lTop, lWidth, lHeight, SWP_NOACTIVATE + SWP_SHOWWINDOW
    DrawForm X, Y
End Sub

Private Sub DrawForm(X, Y)

    Dim X1&, Y1&, X2&, Y2&
    Dim rgn1&, rgn2&
    Dim lBrush&, LB As LOGBRUSH
    Dim Poly As PointType
    
    With Me
        .Cls
        
        X1 = .ScaleLeft
        X2 = .ScaleWidth
        If bBottom = True Then
            Y1 = .ScaleTop
            Y2 = .ScaleHeight - 20
        Else
            Y1 = .ScaleTop + 20
            Y2 = .ScaleHeight
        End If
    End With

    With Poly
        If bLeft Then
            .Point1.X = X1 + 15
            If X < 15 Then
                .Point2.X = ScaleX(X)
            Else
                .Point2.X = .Point1.X
            End If
            .Point3.X = .Point1.X + 19
        Else
            .Point1.X = X2 - 15
            If X > (ScaleX(Screen.Width, vbTwips, vbPixels) - 15) Then
                .Point2.X = Me.ScaleLeft + Me.ScaleWidth
            Else
                .Point2.X = .Point1.X
            End If
            .Point3.X = .Point1.X - 19
        End If
        
        If bBottom = True Then
            .Point1.Y = Y2 - 1
            .Point2.Y = .Point1.Y + 19
            .Point3.Y = .Point1.Y
        Else
            .Point1.Y = Y1 + 1
            .Point2.Y = .Point1.Y - 19
            .Point3.Y = .Point1.Y
        End If
    End With
    
    With LB
        .lbColor = 0
        .lbHatch = 0
        .lbStyle = 0
    End With
    
    rgn1 = CreateRoundRectRgn(X1&, Y1&, X2&, Y2&, 15, 15)
    rgn2 = CreatePolygonRgn(Poly, 3, 2)
    CombineRgn rgn1, rgn1, rgn2, 2
    lBrush = CreateBrushIndirect(LB)
    FrameRgn Me.hdc, rgn1, lBrush, 1, 1
    SetWindowRgn Me.hwnd, rgn1, True
End Sub

Friend Sub HideBalloon()
    Unload Me
End Sub

Private Sub timAutoClose_Timer()
    HideBalloon
End Sub

Private Sub Form_Click()
    HideBalloon
End Sub
Private Sub imgDisplayIcon_Click()
    HideBalloon
End Sub
Private Sub lblText_Click()
    HideBalloon
End Sub
Private Sub lblTitle_Click()
    HideBalloon
End Sub

