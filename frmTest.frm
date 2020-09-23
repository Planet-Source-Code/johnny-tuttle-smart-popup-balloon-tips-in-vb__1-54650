VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{27230248-28BF-4F04-9FA9-496AC1465C22}#1.0#0"; "BalloonTips.ocx"
Begin VB.Form frmTest 
   Caption         =   "Test Balloon Tips"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Control Arrays"
      Height          =   495
      Left            =   2040
      TabIndex        =   14
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   3840
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3135
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5530
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmTest.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Option1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Check1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmTest.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmTest.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Shape2"
      Tab(2).Control(1)=   "Combo1"
      Tab(2).Control(2)=   "Slider1"
      Tab(2).Control(3)=   "VScroll1"
      Tab(2).Control(4)=   "HScroll1"
      Tab(2).ControlCount=   5
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   -72480
         TabIndex        =   13
         Top             =   1680
         Width           =   1575
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1815
         Left            =   -70680
         TabIndex        =   12
         Top             =   840
         Width           =   255
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   495
         Left            =   -74400
         TabIndex        =   11
         Top             =   1920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         _Version        =   393216
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -74400
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   960
         Width           =   1455
      End
      Begin VB.PictureBox Picture1 
         Height          =   2295
         Left            =   -74160
         ScaleHeight     =   2235
         ScaleWidth      =   3195
         TabIndex        =   7
         Top             =   600
         Width           =   3255
         Begin VB.Frame Frame1 
            Caption         =   "Frame1"
            Height          =   1695
            Left            =   240
            TabIndex        =   8
            Top             =   240
            Width           =   2655
            Begin VB.ListBox List1 
               Height          =   1035
               Left            =   480
               TabIndex        =   9
               Top             =   360
               Width           =   1695
            End
         End
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Tag             =   "test"
         Text            =   "Text1"
         Top             =   840
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Left            =   720
         TabIndex        =   4
         Tag             =   $"frmTest.frx":0054
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   375
         Left            =   3000
         TabIndex        =   3
         Tag             =   $"frmTest.frx":006B
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   615
         Left            =   2280
         TabIndex        =   2
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   -72120
         Shape           =   3  'Circle
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Tag             =   "test"
         Top             =   720
         Width           =   1335
      End
      Begin VB.Shape Shape1 
         Height          =   495
         Left            =   360
         Tag             =   "test"
         Top             =   1920
         Width           =   1455
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   1111
      ButtonWidth     =   1270
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Button 1"
            Key             =   "Button1"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Button 2"
            Key             =   "Button2"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Button 3"
            Key             =   "Button3"
         EndProperty
      EndProperty
      Begin BalloonTips.Balloon Balloon1 
         Left            =   3360
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private iIndex As Integer


Private Sub Form_Load()
    'This shows you how easy it is to add tips for standard controls
    With Balloon1
        .AddTip "Label1", "Label", "Don't you hate being labeled?"
        .AddTip "Command1", , "I Command You!"
        .AddTip "Check1", , "Check it out!!"
        .AddTip "Option1", "Optional Title", "Just another option!"
        .AddTip "Text1", , "Text, what would we do without it?"
        .AddTip "Shape1", , "Rectangles are cool!"
        
        'Also note to add objects inside other objects,
        'add the inner most objects first
        .AddTip "List1", , "This is the the list inside a frame, inside a picturebox!"
        .AddTip "Frame1", , "This is a frame inside a picture box!"
        .AddTip "Picture1", , "This is a picture box!"
        
        .AddTip "Combo1", "Combo", "This is a combo box."
        .AddTip "Shape2", "Circle!", "This is a circle!"
        .AddTip "VScroll1", "Scroll Bar", "This is a vertical scroll bar"
        .AddTip "HScroll1", "Scroll Bar", "This is a horizontal scroll bar"
        .AddTip "Slider1", "Weeee!", "This is a slider!" & vbNewLine & "Rhymes with cyder?"
        .AddTip "Command2", "Show Array Example", "Show Example of control arrays on another form."
    End With
End Sub

Private Sub Form_Activate()
    Balloon1.Enabled = True 'enable balloon tips on this form when it is active
End Sub

Private Sub Form_Deactivate()
    'You have to disable the balloon tips when you have multiple forms
    'using balloon tips because they will both check to see if they are
    'in the foreground window and the one that is not will hide the balloon.
    Balloon1.HideBalloon
    Balloon1.Enabled = False
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    'You need to do this when changing tabs because the ssTab control
    'actually moves controls lefts way off screen so they are not visible
    Balloon1.Refresh
End Sub

Private Sub Toolbar1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'This is an example of how you could use the balloon manually
    'for more complex controls
    Dim iNewIndex As Integer, i As Integer
    Dim iLeft As Integer, iRight As Integer, iTop As Integer, iBottom As Integer
    
    iNewIndex = -1
    With Toolbar1
        For i = 1 To .Buttons.Count
            iLeft = .Buttons(i).Left
            iRight = .Buttons(i).Left + .Buttons(i).Width
            iTop = .Buttons(i).Top
            iBottom = .Buttons(i).Top + .Buttons(i).Height
            
            If (x > iLeft And x < iRight) Then
                If (y > iTop And y < iBottom) Then
                    iNewIndex = i
                    Exit For
                End If
            End If
        Next i
    End With
    
    If iNewIndex = -1 Then
        Timer1.Enabled = False
        Balloon1.HideBalloon
    ElseIf iNewIndex <> iIndex Then
        Timer1.Enabled = True
    End If
    iIndex = iNewIndex
End Sub

Private Sub Timer1_Timer()
    'iIndex was determined in the mouse_move event
    Timer1.Enabled = False
    Select Case iIndex
        Case 1
            Balloon1.ShowBalloon "Button 1", "First Button!", ISTYLE_INFO
        Case 2
            Balloon1.ShowBalloon "Button 2", "Second Button!", ISTYLE_WARNING
        Case 3
            Balloon1.ShowBalloon "Button 3", "Third Button!", ISTYLE_ERROR
        Case Else
            Balloon1.HideBalloon
    End Select
End Sub

Private Sub Command2_Click()
    frmArray.Show
End Sub

