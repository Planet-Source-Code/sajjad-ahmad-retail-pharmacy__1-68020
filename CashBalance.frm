VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "PRJCHAMELEON.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCashBalance 
   Caption         =   "Daily Cash Balance Statement"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   5910
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Cash Balance Statement"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   5535
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5535
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   6
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   4320
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   5
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   3720
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   4
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   3240
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   2760
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2280
         TabIndex        =   0
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   20709379
         CurrentDate     =   38416
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Index           =   2
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   2160
         Width           =   1935
      End
      Begin prjChameleon.chameleonButton Command1 
         Default         =   -1  'True
         Height          =   375
         Left            =   4320
         TabIndex        =   1
         Top             =   480
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&View"
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
         BCOL            =   14867671
         BCOLO           =   14867671
         FCOL            =   0
         FCOLO           =   255
         MCOL            =   9469052
         MPTR            =   99
         MICON           =   "CashBalance.frx":0000
         UMCOL           =   0   'False
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton chameleonButton1 
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   4920
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Print Statement"
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
         MCOL            =   9469052
         MPTR            =   99
         MICON           =   "CashBalance.frx":031A
         UMCOL           =   0   'False
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton chameleonButton2 
         Height          =   375
         Left            =   3240
         TabIndex        =   3
         Top             =   4920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Close"
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
         MCOL            =   9469052
         MPTR            =   99
         MICON           =   "CashBalance.frx":0634
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
         BorderStyle     =   1  'Fixed Single
         Height          =   45
         Index           =   1
         Left            =   720
         TabIndex        =   21
         Top             =   4200
         Width           =   4335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   45
         Index           =   0
         Left            =   720
         TabIndex        =   20
         Top             =   2640
         Width           =   4335
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Returned"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   840
         TabIndex        =   18
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   17
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Received"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   840
         TabIndex        =   16
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   15
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   14
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Date to View:"
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
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Return Total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   9
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Net Cash Balance :"
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
         Height          =   375
         Index           =   4
         Left            =   840
         TabIndex        =   8
         Top             =   4320
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmCashBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Private Sub chameleonButton1_Click()
Printer.FontName = "Arial"
Printer.FontSize = 12
Printer.Print Space(90)
Printer.Print Space(20) + "================================================================"
Printer.Print Space(70) + "Cash Balance Statement"
Printer.Print Space(80) + "For: " & Format(DTPicker1.Value, "dd/mm/yyyy")
Printer.Print Space(20) + "================================================================"
Printer.Print Space(15)
Printer.Print Space(20) + Label1(1).Caption + " :" + Space(100) + Format(Text1(0).Text, "0.00")
Printer.Print Space(3)
Printer.Print Space(20) + Label1(3).Caption + " :" + Space(97) + Format(Text1(1).Text, "0.00")
Printer.Print Space(3)
Printer.Print Space(20) + Label1(6).Caption + " :" + Space(94) + Format(Text1(2).Text, "0.00")
Printer.Print Space(3)
Printer.Print Space(20) + Label1(0).Caption + " :" + Space(100) + Format(Text1(3).Text, "0.00")
Printer.Print Space(3)
Printer.Print Space(20) + Label1(2).Caption + " :" + Space(98) + Format(Text1(4).Text, "0.00")
Printer.Print Space(3)
Printer.Print Space(20) + Label1(7).Caption + " :" + Space(95) + Format(Text1(5).Text, "0.00")
Printer.Print Space(15)
Printer.Print Space(20) + "================================================================"
Printer.Print Space(20) + Label1(4).Caption + " :" + Space(86) + Format(Text1(6).Text, "0.00")
Printer.Print Space(20) + "================================================================"

Printer.EndDoc

End Sub

Private Sub chameleonButton2_Click()
Unload Me
End Sub

Private Sub Command1_Click()
GetCashBal
End Sub

Private Sub Form_Load()
Me.DTPicker1.Value = Date
GetCashBal
End Sub

Private Sub GetCashBal()
On Error Resume Next

rs.Open "Select * from 3SaleDiscNet where billdate=#" & DTPicker1.Value & "#", cn, adOpenForwardOnly, adLockReadOnly
rs1.Open "Select * from 4ReturnDiscNet where billdate=#" & DTPicker1.Value & "#", cn, adOpenForwardOnly, adLockReadOnly

If rs.EOF Or rs.BOF Then
Me.Text1(0).Text = 0
Me.Text1(1).Text = 0
Me.Text1(2).Text = 0
Else
Me.Text1(0).Text = rs.Fields(1)
Me.Text1(1).Text = rs.Fields(2)
Me.Text1(2).Text = rs.Fields(3)
End If

If rs1.EOF Or rs1.BOF Then
Me.Text1(3).Text = 0
Me.Text1(4).Text = 0
Me.Text1(5).Text = 0

Else
Me.Text1(3).Text = rs1.Fields(1)
Me.Text1(4).Text = rs1.Fields(2)
Me.Text1(5).Text = rs1.Fields(3)

End If

Text1(6) = Format(Text1(2) - Text1(5), "0.00")

Set rs = Nothing
Set rs1 = Nothing
End Sub

Private Sub Form_Resize()
CenterCtrl Me, Frame1
End Sub

