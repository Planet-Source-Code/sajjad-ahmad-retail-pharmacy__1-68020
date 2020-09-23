VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "PRJCHAMELEON.OCX"
Begin VB.Form frmRegister 
   BackColor       =   &H00FFFFFF&
   Caption         =   "License Information"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5520
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   Picture         =   "Register.frx":0000
   ScaleHeight     =   3645
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1950
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   2430
      Width           =   2655
   End
   Begin prjChameleon.chameleonButton Command1 
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   3000
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Validate License Information"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   4210752
      FCOLO           =   255
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "Register.frx":433CE
      UMCOL           =   0   'False
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton Command2 
      Height          =   375
      Left            =   3840
      TabIndex        =   9
      Top             =   3000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   4210752
      FCOLO           =   255
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "Register.frx":436E8
      UMCOL           =   0   'False
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Software Code:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   7
      Top             =   1950
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "License Key :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   6
      Top             =   2430
      Width           =   1695
   End
   Begin VB.Line Line2 
      DrawMode        =   10  'Mask Pen
      X1              =   360
      X2              =   5040
      Y1              =   1365
      Y2              =   1365
   End
   Begin VB.Image Image1 
      Height          =   1125
      Left            =   3840
      Picture         =   "Register.frx":43A02
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1170
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2003-2005  IBASOFT. All rights reserved."
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   1440
      Width           =   3855
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pharmacy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   840
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   3465
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1635
      TabIndex        =   0
      Top             =   1005
      Width           =   1275
   End
   Begin VB.Line Line1 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   360
      X2              =   5040
      Y1              =   1365
      Y2              =   1365
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pharmacy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   840
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   165
      Width           =   3465
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text2.Text = "" Then Exit Sub

'ActiveLock1.Register (Text2)

'If ActiveLock1.RegisteredUser = True Then
 '   MsgBox "Your Registration has been accepted. Thank you!", vbInformation, "Pharmacy Registration"
'Else
 '   MsgBox "You didn't type the right key, sorry.", vbExclamation, "Pharmacy Registration"
'End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
'Me.Text1 = ActiveLock1.SoftwareCode
End Sub

