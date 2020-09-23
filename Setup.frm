VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "PRJCHAMELEON.OCX"
Begin VB.Form frmSetup 
   Caption         =   "Company Information"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6120
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   4560
   ScaleWidth      =   6120
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Compnay Setup Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   4335
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5895
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   240
         TabIndex        =   11
         Top             =   2760
         Width           =   5295
         Begin VB.TextBox Text6 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   3000
            MaxLength       =   10
            TabIndex        =   14
            ToolTipText     =   "Please Enter a character between A-Z, 0-9 for unlocking the Lock Screen."
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox Text5 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   480
            MaxLength       =   10
            TabIndex        =   12
            ToolTipText     =   "Please Enter a character between A-Z, 0-9 for unlocking the Lock Screen."
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Admin Password:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   5
            Left            =   3000
            TabIndex        =   15
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Login Password:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   4
            Left            =   480
            TabIndex        =   13
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   720
         IMEMode         =   3  'DISABLE
         Left            =   1440
         TabIndex        =   1
         Top             =   1320
         Width           =   4095
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1440
         TabIndex        =   0
         Top             =   720
         Width           =   4095
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1440
         TabIndex        =   2
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   4920
         MaxLength       =   1
         TabIndex        =   3
         ToolTipText     =   "Please Enter a character between A-Z, 0-9 for unlocking the Lock Screen."
         Top             =   2280
         Width           =   615
      End
      Begin prjChameleon.chameleonButton Command1 
         Height          =   375
         Left            =   600
         TabIndex        =   4
         Top             =   3840
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Save Information"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         MICON           =   "Setup.frx":0000
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
         Left            =   4080
         TabIndex        =   5
         Top             =   3840
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
            Weight          =   700
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
         MICON           =   "Setup.frx":031A
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
         Caption         =   "Lock Key:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   3960
         TabIndex        =   10
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
If Text1.Text = "" Then Exit Sub
If Text2.Text = "" Then Exit Sub

With rs
    .Fields(0) = Text1.Text
    .Fields(1) = Text2.Text
    .Fields(2) = Text3.Text
    .Fields(3) = Text4.Text
    .Fields(4) = Text5.Text
    .Fields(5) = Text6.Text
    .Update
End With

With CURRENT_COMPANY
    .COMPANY_NAME = Text1.Text
    .COMPANY_ADDRESS = Text2.Text
    .COMPANY_PHONE = Text3.Text
    .LOCK_KEY = Text4.Text
    .DB_PASSWORD = Text5.Text
    .ADMIN_PASS = Text6.Text
End With
MsgBox "Changes has been successfully saved.", vbInformation
Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next

rs.Open "SELECT * FROM Setup", cn, adOpenStatic, adLockOptimistic

Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
Text4.Text = rs.Fields(3)
Text5.Text = rs.Fields(4)
Text6.Text = rs.Fields(5)

End Sub

Private Sub Form_Resize()
CenterCtrl Me, Frame1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rs = Nothing
End Sub


