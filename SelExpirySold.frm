VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "PRJCHAMELEON.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmExpirySold 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Report Selection"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3825
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3615
      Begin VB.OptionButton Option2 
         BackColor       =   &H00800000&
         Caption         =   "Medicines that were not sold from"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H000000FF&
         Caption         =   "Medicines that will be Expired After"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   3375
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   1020
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckBox        =   -1  'True
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   20709379
      CurrentDate     =   38431
   End
   Begin prjChameleon.chameleonButton Command1 
      Default         =   -1  'True
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
      _ExtentX        =   1931
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   255
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "SelExpirySold.frx":0000
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
      Left            =   2040
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
      _ExtentX        =   1931
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   255
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "SelExpirySold.frx":031A
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
      Caption         =   "Select Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "frmExpirySold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Private Sub Command1_Click()
If Option1.Value = True Then
    rs.Open "Select * from ProductESQry Where expirydate < #" & Format(DTPicker1.Value, "dd/mm/yyyy") & "#" & " And Unit_Stock>0", cn, adOpenForwardOnly, adLockReadOnly
    If rs.EOF Or rs.BOF Then
        MsgBox "No data for this Report! Canceling...", vbInformation
        Set rs = Nothing
        Exit Sub
    Else
        Set ExpirySold.DataSource = rs
        ExpirySold.Sections(2).Controls.Item("Label1").Caption = "Products That Will be Expired After"
        ExpirySold.Sections(2).Controls.Item("Label2").Caption = Format(DTPicker1.Value, "dd/mm/yyyy")
        Me.Hide
        ExpirySold.Show
        Set rs = Nothing
    End If
ElseIf Option2.Value = True Then
    If DTPicker1.Value > Date Then
        MsgBox "Please select a date before Today", vbInformation
        Exit Sub
    Else
    rs.Open "Select * from ProductESQry Where LastSold Is Null OR LastSold< #" & Format(DTPicker1.Value, "dd/mm/yyyy") & "#", cn, adOpenForwardOnly, adLockReadOnly
    If rs.EOF Or rs.BOF Then
        MsgBox "No data for this Report! Canceling...", vbInformation
        Set rs = Nothing
        Exit Sub
    Else
        Set ExpirySold.DataSource = rs
        ExpirySold.Sections(2).Controls.Item("Label1").Caption = "Products That Were Not Sold Since"
        ExpirySold.Sections(2).Controls.Item("Label2").Caption = Format(DTPicker1.Value, "dd/mm/yyyy")
        Me.Hide
        ExpirySold.Show
        Set rs = Nothing
    End If
    End If
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
DTPicker1.Value = Date
End Sub
