VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "PRJCHAMELEON.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Login Information"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6585
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   2280
      Width           =   3975
   End
   Begin prjChameleon.chameleonButton Command1 
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   3360
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
      FCOL            =   0
      FCOLO           =   255
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "Login.frx":0000
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
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&OK"
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
      FCOL            =   0
      FCOLO           =   255
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "Login.frx":031A
      UMCOL           =   0   'False
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password to Access System"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   5
      Top             =   1920
      Width           =   3975
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
      Left            =   3360
      TabIndex        =   4
      Top             =   1200
      Width           =   1275
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
      Left            =   1760
      TabIndex        =   3
      Top             =   320
      Width           =   3465
   End
   Begin VB.Image Image1 
      Height          =   1125
      Left            =   5280
      Picture         =   "Login.frx":0634
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1170
   End
   Begin VB.Line Line2 
      DrawMode        =   10  'Mask Pen
      X1              =   1800
      X2              =   6240
      Y1              =   1560
      Y2              =   1560
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
      Left            =   1800
      TabIndex        =   6
      Top             =   360
      Width           =   3465
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   1800
      X2              =   6240
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Image Image4 
      Height          =   4065
      Left            =   1920
      Picture         =   "Login.frx":1830
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4680
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   4215
      Left            =   0
      Picture         =   "Login.frx":1C84
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   2055
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
If MsgBox("Are you sure to Exit?", vbYesNo, "Exit Confirmation") = vbYes Then
    End
Else
    Exit Sub
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
If CN.State = adStateClosed Then OpenDB
If Text1.Text = "" Then
    Text1.SetFocus
    Exit Sub
End If

rs.Open "Select * from Setup", CN, adOpenKeyset, adLockOptimistic

If Text1.Text = rs.Fields(4) Then
    With CURRENT_COMPANY
    .COMPANY_NAME = rs.Fields(0)
    .COMPANY_ADDRESS = rs.Fields(1)
    .COMPANY_PHONE = rs.Fields(2)
    .LOCK_KEY = rs.Fields(3)
    .DB_PASSWORD = rs.Fields(4)
    .ADMIN_PASS = rs.Fields(5)
    End With
    Unload Me
    MainMDI.Show
ElseIf Text1.Text = rs.Fields(5) Then
    With CURRENT_COMPANY
    .COMPANY_NAME = rs.Fields(0)
    .COMPANY_ADDRESS = rs.Fields(1)
    .COMPANY_PHONE = rs.Fields(2)
    .LOCK_KEY = rs.Fields(3)
    .DB_PASSWORD = rs.Fields(4)
    .ADMIN_PASS = rs.Fields(5)
    End With
    Unload Me
    
    MainMDI.mpricelist.Visible = True
    MainMDI.mconfiguration.Visible = True
    MainMDI.lop2.Visible = True
    MainMDI.Show
Else
   MsgBox "Access Denied!", vbCritical, "Login Failure"
End If
    
Set rs = Nothing

End Sub

Private Sub Form_Load()
OpenDB
End Sub



