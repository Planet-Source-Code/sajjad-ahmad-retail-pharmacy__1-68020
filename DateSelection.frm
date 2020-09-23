VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "PRJCHAMELEON.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDateSelection 
   Caption         =   "Bills Backlog Selection"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3360
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   1440
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckBox        =   -1  'True
      CustomFormat    =   "dd/MMM/yyyy"
      Format          =   58720259
      CurrentDate     =   38410
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detail Options:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3135
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Sale of Medical Store"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   360
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Return of Purchased Medicine"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Return of Sold Medicine"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Purchase for Medical Store"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   2895
      End
   End
   Begin prjChameleon.chameleonButton Command1 
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   2040
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
      MICON           =   "DateSelection.frx":0000
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
      Left            =   1800
      TabIndex        =   3
      Top             =   2040
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
      MICON           =   "DateSelection.frx":031A
      UMCOL           =   0   'False
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   345
      Left            =   1440
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckBox        =   -1  'True
      CustomFormat    =   "dd/MMM/yyyy"
      Format          =   58720259
      CurrentDate     =   38410
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   " Start Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008000&
      Caption         =   " End Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1335
   End
End
Attribute VB_Name = "frmDateSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset

Private Sub Command1_Click()

If Me.DTPicker1.Value > Date Or Me.DTPicker2.Value > Date Then
    MsgBox "You must pick a date before today!", vbInformation
    Exit Sub
ElseIf Me.DTPicker1.Value > Me.DTPicker2.Value Then
    MsgBox "StarDate must be less than end date!", vbInformation
    Exit Sub
Else
If Option1.Value = True Then
    rs.Open "Select * from SalebyDate Where BillDate >= #" & DTPicker1.Value & "#" & "and BillDate <= #" & DTPicker2.Value & "#", cn, adOpenKeyset, adLockOptimistic
    If rs.EOF Or rs.BOF Then
        MsgBox "No Data for this Report! Canceling...", vbInformation
        Set rs = Nothing
        Exit Sub
    Else
    Set OrdersBacklog.DataSource = rs
    OrdersBacklog.Sections(2).Controls.Item("Label8").Caption = "Product Sale Report"
    OrdersBacklog.Sections(2).Controls.Item("Label9").Caption = "From " & DTPicker1.Value & " To " & DTPicker2.Value
    OrdersBacklog.Show
    Set rs = Nothing
    Unload frmDateSelection
    End If
ElseIf Option2.Value = True Then
    rs1.Open "Select * from PurchasebyDate Where BillDate >= #" & DTPicker1.Value & "#" & "and BillDate <= #" & DTPicker2.Value & "#", cn, adOpenKeyset, adLockOptimistic
    If rs1.EOF Or rs1.BOF Then
        MsgBox "No Data for this Report! Canceling...", vbInformation
        Set rs1 = Nothing
        Exit Sub
    Else
    Set POrdersBacklog.DataSource = rs1
    POrdersBacklog.Sections(2).Controls.Item("Label8").Caption = "Purchase For Medical Store"
    POrdersBacklog.Sections(2).Controls.Item("Label9").Caption = "From " & DTPicker1.Value & "  To  " & DTPicker2.Value
    POrdersBacklog.Show
    Set rs1 = Nothing
    Unload frmDateSelection
    End If
ElseIf Option3.Value = True Then
    rs.Open "Select * from RefundbyDate Where BillDate >= #" & DTPicker1.Value & "#" & "and BillDate <= #" & DTPicker2.Value & "#", cn, adOpenKeyset, adLockOptimistic
    If rs.EOF Or rs.BOF Then
        MsgBox "No Data for this Report! Canceling...", vbInformation
        Set rs = Nothing
        Exit Sub
    Else
    Set DailyRefundSale.DataSource = rs
    DailyRefundSale.Sections(2).Controls.Item("Label8").Caption = "Report of Return Medicine"
    DailyRefundSale.Sections(2).Controls.Item("Label9").Caption = "From " & DTPicker1.Value & " To " & DTPicker2.Value
    DailyRefundSale.Show
    Set rs = Nothing
    Unload frmDateSelection
    End If
ElseIf Option4.Value = True Then
    rs1.Open "Select * from PRefundbyDate Where BillDate >= #" & DTPicker1.Value & "#" & "and BillDate <= #" & DTPicker2.Value & "#", cn, adOpenKeyset, adLockOptimistic
    If rs1.EOF Or rs1.BOF Then
        MsgBox "No Data for this Report! Canceling...", vbInformation
        Set rs1 = Nothing
        Exit Sub
    Else
    Set POrdersBacklog.DataSource = rs1
    POrdersBacklog.Sections(2).Controls.Item("Label8").Caption = "Report of Return Medicine to Supplier"
    POrdersBacklog.Sections(2).Controls.Item("Label9").Caption = "From " & DTPicker1.Value & " To " & DTPicker2.Value
    POrdersBacklog.Show
    Set rs1 = Nothing
    Unload frmDateSelection
    End If
End If
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.DTPicker1.Value = Date
Me.DTPicker2.Value = Date
End Sub


