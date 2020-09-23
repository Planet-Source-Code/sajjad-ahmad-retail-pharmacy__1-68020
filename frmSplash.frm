VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   ClientHeight    =   4500
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   4500
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   3720
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   840
      Top             =   2760
   End
   Begin VB.Label lblCompany 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Company"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label lblPlatform 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "License Information:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   2400
      TabIndex        =   0
      Top             =   1560
      Width           =   3375
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
    frmLogin.Show
End Sub

Private Sub Frame1_Click()
    Unload Me
    frmLogin.Show
End Sub

Private Sub Timer1_Timer()
If Me.ProgressBar1.Value = 100 Then
    Unload Me
    frmLogin.Show
Else
    Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
End If

End Sub

