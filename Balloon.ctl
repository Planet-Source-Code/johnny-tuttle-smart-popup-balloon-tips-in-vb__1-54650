VERSION 5.00
Begin VB.UserControl Balloon 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   CanGetFocus     =   0   'False
   ClientHeight    =   1380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1425
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Balloon.ctx":0000
   PropertyPages   =   "Balloon.ctx":0974
   ScaleHeight     =   1380
   ScaleWidth      =   1425
   ToolboxBitmap   =   "Balloon.ctx":0986
   Begin VB.Timer BalloonTimer 
      Enabled         =   0   'False
      Left            =   720
      Top             =   720
   End
   Begin VB.Timer CtrlTimer 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   240
      Top             =   720
   End
End
Attribute VB_Name = "Balloon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'I rewrote much of this code. Original credit goes to Evan Toder
'for his CtlRect control. You can find it on planet source code
'by searching for "MouseEnter/Exit for ALL controls 1 line!!"
'
'Basically this creates RECT coordinates for each control,
'and searches if the mouse is over the coordinates
'
'Sorry this is all the comments, they distract me when I'm programming.

Option Explicit

Private Const CCHILDREN_TITLEBAR = 5

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type TITLEBARINFO
    cbSize As Long
    rcTitleBar As RECT
    rgstate(CCHILDREN_TITLEBAR) As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function GetTitleBarInfo Lib "user32.dll" (ByVal hwnd As Long, ByRef pti As TITLEBARINFO) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long

Public Event MouseEnterControl(CtrlName As String, ByRef CancelBalloon As Boolean)
Public Event MouseExitControl(CtrlName As String)

Private WithEvents F As Form
Attribute F.VB_VarHelpID = -1

Public Enum IconStyle
    ISTYLE_NONE = 1
    ISTYLE_INFO = 2
    ISTYLE_WARNING = 3
    ISTYLE_ERROR = 4
End Enum

Private Type BalloonTip
    CtrlName As String
    BalloonTitle As String
    BalloonText As String
    BalloonIcon As IconStyle
    CtrlRect As RECT
End Type

Private tBarHeight&
Private mBarHeight&
Private pt As POINTAPI
Private CurrentIndex As Integer
Private BTip() As BalloonTip
Private bGetHeights As Boolean

Private m_OnTimer As Integer
Private m_OffTimer As Integer
Private m_DefaultIconStyle As IconStyle
Private m_DefaultTitle As String
Private m_AutoHide As Boolean
Private m_Enabled As Boolean

Private Const m_def_OnTimer = 500
Private Const m_def_OffTimer = 0
Private Const m_def_Icon = ISTYLE_INFO
Private Const m_def_DefaultTitle = "Description"
Private Const m_def_AutoHide = True
Private Const m_def_Enabled = True


Public Property Let OnTimer(Interval As Integer)
    m_OnTimer = Interval
    PropertyChanged "OnTimer"
End Property
Public Property Get OnTimer() As Integer
Attribute OnTimer.VB_ProcData.VB_Invoke_Property = ";Balloon"
    OnTimer = m_OnTimer
End Property
Public Property Let OffTimer(Interval As Integer)
    m_OffTimer = Interval
    PropertyChanged "OffTimer"
End Property
Public Property Get OffTimer() As Integer
Attribute OffTimer.VB_ProcData.VB_Invoke_Property = ";Balloon"
    OffTimer = m_OffTimer
End Property
Public Property Let DefaultIconStyle(IStyle As IconStyle)
    m_DefaultIconStyle = IStyle
    PropertyChanged "DefaultIconStyle"
End Property
Public Property Get DefaultIconStyle() As IconStyle
Attribute DefaultIconStyle.VB_ProcData.VB_Invoke_Property = ";Balloon"
    DefaultIconStyle = m_DefaultIconStyle
End Property
Public Property Let DefaultTitle(DefaultTitle As String)
    m_DefaultTitle = DefaultTitle
    PropertyChanged "DefaultTitle"
End Property
Public Property Get DefaultTitle() As String
Attribute DefaultTitle.VB_ProcData.VB_Invoke_Property = ";Balloon"
    DefaultTitle = m_DefaultTitle
End Property
Public Property Let AutoHide(Hide As Boolean)
    m_AutoHide = Hide
    PropertyChanged "AutoHide"
End Property
Public Property Get AutoHide() As Boolean
Attribute AutoHide.VB_ProcData.VB_Invoke_Property = ";Balloon"
    AutoHide = m_AutoHide
End Property
Public Property Let Enabled(IsEnabled As Boolean)
    m_Enabled = IsEnabled
    CtrlTimer.Enabled = m_Enabled
    If m_Enabled = False Then BalloonTimer.Enabled = False
    PropertyChanged "Enabled"
End Property
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Balloon"
    Enabled = m_Enabled
End Property

Private Sub CtrlTimer_Timer()
    Call GetFormPos
End Sub

Private Sub BalloonTimer_Timer()
    BalloonTimer.Enabled = False
    If CurrentIndex < 0 Then Exit Sub
    
    GetCursorPos pt
    
    Dim sTitle As String, IStyle As IconStyle
    If LenB(BTip(CurrentIndex).BalloonTitle) = 0 Then
        sTitle = m_DefaultTitle
    Else
        sTitle = BTip(CurrentIndex).BalloonTitle
    End If
    If BTip(CurrentIndex).BalloonIcon = 0 Then
        IStyle = m_DefaultIconStyle
    Else
        IStyle = BTip(CurrentIndex).BalloonIcon
    End If
    
    frmBalloon.ShowBalloon sTitle, BTip(CurrentIndex).BalloonText, pt.X, pt.Y, IStyle, m_OffTimer
End Sub

Private Function GetTitleBarHeight(lHwnd&) As Long
    Dim TitleInfo As TITLEBARINFO

    TitleInfo.cbSize = Len(TitleInfo)
    GetTitleBarInfo lHwnd, TitleInfo
    GetTitleBarHeight = 4 + (TitleInfo.rcTitleBar.Bottom - TitleInfo.rcTitleBar.Top)
End Function

Private Function GetMenuBarHeight(lHwnd&) As Long
    If GetMenu(lHwnd) <> 0 Then GetMenuBarHeight = 20
End Function

Public Sub Clear()
    ReDim BTip(0)
    CurrentIndex = -1
    CtrlTimer.Enabled = False
    BalloonTimer.Enabled = False
End Sub

Public Sub AddTip(CtrlName As String, Optional BalloonTitle As String, Optional BalloonTip As String, Optional BalloonIcon As IconStyle)
    If bGetHeights Then
        tBarHeight = GetTitleBarHeight(F.hwnd)
        mBarHeight = GetMenuBarHeight(F.hwnd)
        bGetHeights = False
    End If
    
    Dim Index As Integer, CName As String
    Dim C As Control
    CName = CtrlName
    Set C = GetCtrl(CName)
    
    If C Is Nothing Then
        Error 5
        Exit Sub
    End If
    
    If Not (UBound(BTip) = 0 And LenB(BTip(0).CtrlName) = 0) Then
        Index = UBound(BTip) + 1
        ReDim Preserve BTip(0 To Index)
    End If
        
    With BTip(Index)
        .CtrlName = CtrlName
        .BalloonTitle = BalloonTitle
        .BalloonText = BalloonTip
        .BalloonIcon = BalloonIcon
    End With
    SetCtrlRect C, Index
    If m_Enabled Then CtrlTimer.Enabled = True
End Sub

Public Sub ChangeTip(CtrlName As String, Optional BalloonTitle As String, Optional BalloonTip As String, Optional BalloonIcon As IconStyle)
    Dim Index As Integer
    Index = CtrlIndex(CtrlName)
    If Index > -1 Then
        With BTip(Index)
            .BalloonTitle = BalloonTitle
            .BalloonText = BalloonTip
            .BalloonIcon = BalloonIcon
        End With
    End If
End Sub

Public Sub RemoveTip(CtrlName As String)
    Dim Index As Integer, i As Integer
    Index = CtrlIndex(CtrlName)
    If Index > -1 Then
        If UBound(BTip) = 0 Then
            Clear
        Else
            For i = (Index + 1) To UBound(BTip)
                BTip(i - 1) = BTip(i)
            Next i
            ReDim Preserve BTip(0 To UBound(BTip) - 1)
        End If
    End If
End Sub

Public Sub Refresh()
    Dim Index As Integer
    For Index = 0 To UBound(BTip)
        SetCtrlRect GetCtrl(BTip(Index).CtrlName), Index
    Next Index
End Sub

Public Sub RefreshTip(CtrlName As String)
    Dim Index As Integer
    Index = CtrlIndex(CtrlName)
    If Index > -1 Then
        SetCtrlRect GetCtrl(CtrlName), Index
    End If
End Sub

Private Sub SetCtrlRect(Ctrl As Control, Index As Integer)
    Dim lPos&, rPos&, tPos&, bPos&, sX&, sY&
    Dim C As Control
    Dim bDefault As Boolean

    sX = Screen.TwipsPerPixelX
    sY = Screen.TwipsPerPixelY

    Set C = Ctrl
    lPos = C.Left + 20
    tPos = C.Top

    Do Until C.Container Is F
        Set C = C.Container
        lPos = lPos + C.Left + 20
        tPos = tPos + C.Top
    Loop
    
    With Ctrl
        lPos = (lPos / sX)
        rPos = (.Width / sX) + lPos
        tPos = (tPos / sY)
        bPos = (.Height / sY) + tPos
    End With
    
    SetRect BTip(Index).CtrlRect, lPos, tPos, rPos, bPos
End Sub

Private Function GetFormPos() As RECT
    Dim L As Integer, GCR As RECT, lR As RECT
    Dim CancelBalloon As Boolean
    Dim C As Control, bSkip As Boolean, bKeepLooking As Boolean

    If GetForegroundWindow <> F.hwnd Then
        BalloonTimer.Enabled = False
        Unload frmBalloon
        Exit Function
    End If
    If F.Visible = False Then Exit Function
    
    GetWindowRect F.hwnd, GetFormPos
    GetCursorPos pt

    For L = 0 To UBound(BTip)
        lR = BTip(L).CtrlRect
        OffsetRect lR, GetFormPos.Left, (GetFormPos.Top + tBarHeight + mBarHeight)
        
        If PtInRect(lR, pt.X, pt.Y) Then
            If L <> CurrentIndex Then
                bSkip = False
                Set C = GetCtrl(BTip(L).CtrlName)
                If (C.Visible = False Or C.Left < 0) Then
                    bSkip = True
                Else
                    Do Until C.Container Is F
                        Set C = C.Container
                        If (C.Visible = False Or C.Left < 0) Then
                            bSkip = True
                            Exit Do
                        End If
                    Loop
                End If
                If bSkip = False Then
                    If CurrentIndex > -1 Then RaiseEvent MouseExitControl(BTip(CurrentIndex).CtrlName)
                    RaiseEvent MouseEnterControl(BTip(L).CtrlName, CancelBalloon)
                    CurrentIndex = L
                    If CancelBalloon Then
                        BalloonTimer.Enabled = False
                    Else
                        BalloonTimer.Interval = m_OnTimer
                        BalloonTimer.Enabled = True
                    End If
                    Exit Function
                End If
                bKeepLooking = True
            Else
                If Not bKeepLooking Then Exit Function
            End If
        End If
    Next L
    
    If CurrentIndex > -1 Then
       BalloonTimer.Enabled = False
       RaiseEvent MouseExitControl(BTip(CurrentIndex).CtrlName)
       CurrentIndex = -1
       If m_AutoHide Then Unload frmBalloon
    End If
End Function

Private Function GetCtrl(ByVal CName As String) As Control
    Dim C As Control, i As Integer, Index As Integer
    
    i = InStr(CName, "(")
    If i > 0 And Right$(CName, 1) = ")" Then
        Index = Val(Mid$(CName, i + 1, Len(CName) - (i + 1)))
        CName = Left$(CName, i - 1)
    Else
        Index = -1
    End If
    
    On Error Resume Next
    For Each C In F.Controls
        If UCase$(C.Name) = UCase$(CName) Then
            If Index = -1 Then
                Set GetCtrl = C
                Exit Function
            Else
                If C.Index = Index Then
                    Set GetCtrl = C
                    Exit Function
                End If
            End If
        End If
    Next
    Set GetCtrl = Nothing
End Function

Private Function CtrlIndex(CtrlName As String) As Integer
    Dim Index As Integer
    For Index = 0 To UBound(BTip)
        If UCase$(BTip(Index).CtrlName) = UCase$(CtrlName) Then
            CtrlIndex = Index
            Exit Function
        End If
    Next Index
    CtrlIndex = -1
End Function

Public Sub ShowBalloon(BalloonTitle As String, BalloonTip As String, _
                       Optional BalloonIcon As IconStyle, Optional CloseTimer As Integer, _
                       Optional X As Long = -1, Optional Y As Long = -1)
    If X = -1 And Y = -1 Then
        GetCursorPos pt
        X = pt.X
        Y = pt.Y
    End If
    Unload frmBalloon
    frmBalloon.ShowBalloon BalloonTitle, BalloonTip, X, Y, BalloonIcon, CloseTimer
End Sub

Public Sub HideBalloon()
    Unload frmBalloon
End Sub

Private Sub UserControl_InitProperties()
    m_OnTimer = m_def_OnTimer
    m_OffTimer = m_def_OffTimer
    m_DefaultIconStyle = m_def_Icon
    m_DefaultTitle = m_def_DefaultTitle
    m_AutoHide = m_def_AutoHide
    m_Enabled = m_def_Enabled
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ReDim BTip(0)
    
    m_OnTimer = PropBag.ReadProperty("OnTimer", m_def_OnTimer)
    m_OffTimer = PropBag.ReadProperty("OffTimer", m_def_OffTimer)
    m_DefaultIconStyle = PropBag.ReadProperty("DefaultIconStyle", m_def_Icon)
    m_DefaultTitle = PropBag.ReadProperty("DefaultTitle", m_def_DefaultTitle)
    m_AutoHide = PropBag.ReadProperty("AutoHide", m_def_AutoHide)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
        
    If Ambient.UserMode = True Then
        Set F = UserControl.Extender.Parent
        CtrlTimer.Enabled = True
        bGetHeights = True
    Else
        CtrlTimer.Enabled = False
        Set F = Nothing
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "OnTimer", m_OnTimer, m_def_OnTimer
    PropBag.WriteProperty "OffTimer", m_OffTimer, m_def_OffTimer
    PropBag.WriteProperty "DefaultIconStyle", m_DefaultIconStyle, m_def_Icon
    PropBag.WriteProperty "DefaultTitle", m_DefaultTitle, m_def_DefaultTitle
    PropBag.WriteProperty "AutoHide", m_AutoHide, m_def_AutoHide
    PropBag.WriteProperty "Enabled", m_Enabled, m_def_Enabled
End Sub

Private Sub UserControl_Resize()
   Size 420, 420
End Sub

Private Sub UserControl_Terminate()
    Unload frmBalloon
End Sub

