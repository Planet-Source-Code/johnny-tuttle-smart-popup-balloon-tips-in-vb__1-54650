VERSION 5.00
Object = "{27230248-28BF-4F04-9FA9-496AC1465C22}#1.0#0"; "BalloonTips.ocx"
Begin VB.Form frmArray 
   Caption         =   "Form1"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Remove Text Box"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Text Box"
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox txtArray 
      Height          =   285
      Index           =   0
      Left            =   2280
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin BalloonTips.Balloon Balloon1 
      Left            =   120
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Label lblArray 
      Caption         =   "Label Array 3"
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblArray 
      Caption         =   "Label Array 2"
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblArray 
      Caption         =   "Label Array 1"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblArray 
      Caption         =   "Label Array 0"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    With Balloon1
        .AddTip "lblArray(0)", "Label Array 0", "This is index 0 of the array!"
        .AddTip "lblArray(1)", "Label Array 1", "This is index 1 of the array!"
        .AddTip "lblArray(2)", "Label Array 2", "This is index 2 of the array!"
        .AddTip "lblArray(3)", "Label Array 3", "This is index 3 of the array!"
        .AddTip "txtArray(0)", "Text Array 0", "This is index 0 of the array!"
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


Private Sub Command1_Click()
    'add a dynamic control at runtime
    Dim i As Integer
    i = txtArray.UBound + 1
    If i = 7 Then Exit Sub
    Load txtArray(i)
    txtArray(i).Top = txtArray(txtArray.UBound - 1).Top + 400
    txtArray(i).Visible = True
    'add a balloon tip for the new control
    Balloon1.AddTip "txtArray(" & i & ")", "Text Array " & i, "This is index " & i & " of the array!"
End Sub

Private Sub Command2_Click()
    'remove dynamic control
    Dim i As Integer
    i = txtArray.UBound
    If i = 0 Then Exit Sub
    Unload txtArray(i)
    'remove balloon tip for control
    Balloon1.RemoveTip "txtArray(" & i & ")"
End Sub
