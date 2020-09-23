VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "PRJCHAMELEON.OCX"
Begin VB.Form frmSupplier 
   Caption         =   "Supplier Information"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7950
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   4905
   ScaleWidth      =   7950
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Supplier Information"
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
      Height          =   4695
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7455
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   0
         TabIndex        =   11
         Top             =   3960
         Width           =   7455
         Begin prjChameleon.chameleonButton cmdFirst 
            Height          =   375
            Left            =   240
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   240
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   ""
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
            FOCUSR          =   0   'False
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   255
            MCOL            =   14867671
            MPTR            =   99
            MICON           =   "Supplier.frx":0000
            PICN            =   "Supplier.frx":031A
            PICH            =   "Supplier.frx":06B4
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   -1  'True
            FX              =   1
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdPrev 
            Height          =   375
            Left            =   720
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   240
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   ""
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
            FOCUSR          =   0   'False
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   255
            MCOL            =   14867671
            MPTR            =   99
            MICON           =   "Supplier.frx":0A4E
            PICN            =   "Supplier.frx":0D68
            PICH            =   "Supplier.frx":1102
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   -1  'True
            FX              =   1
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdNext 
            Height          =   375
            Left            =   6240
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   240
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   ""
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
            FOCUSR          =   0   'False
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   255
            MCOL            =   14867671
            MPTR            =   99
            MICON           =   "Supplier.frx":149C
            PICN            =   "Supplier.frx":17B6
            PICH            =   "Supplier.frx":1B50
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   -1  'True
            FX              =   1
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdLast 
            Height          =   375
            Left            =   6720
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   240
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   ""
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
            FOCUSR          =   0   'False
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   255
            MCOL            =   14867671
            MPTR            =   99
            MICON           =   "Supplier.frx":1EEA
            PICN            =   "Supplier.frx":2204
            PICH            =   "Supplier.frx":259E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   -1  'True
            FX              =   1
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdClose 
            Height          =   375
            Left            =   5160
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "&Close"
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
            MCOL            =   9469052
            MPTR            =   99
            MICON           =   "Supplier.frx":2938
            UMCOL           =   0   'False
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   -1  'True
            FX              =   0
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdAdd 
            Height          =   375
            Left            =   1320
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   240
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "&New"
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
            MCOL            =   9469052
            MPTR            =   99
            MICON           =   "Supplier.frx":2C52
            UMCOL           =   0   'False
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   -1  'True
            FX              =   0
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdEdit 
            Height          =   375
            Left            =   2280
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   240
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "&Edit"
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
            MCOL            =   9469052
            MPTR            =   99
            MICON           =   "Supplier.frx":2F6C
            UMCOL           =   0   'False
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   -1  'True
            FX              =   0
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdUpdate 
            Height          =   375
            Left            =   1320
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   240
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "&Save"
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
            MCOL            =   9469052
            MPTR            =   99
            MICON           =   "Supplier.frx":3286
            UMCOL           =   0   'False
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   -1  'True
            FX              =   0
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdCancel 
            Height          =   375
            Left            =   2280
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   240
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "&Undo"
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
            MCOL            =   9469052
            MPTR            =   99
            MICON           =   "Supplier.frx":35A0
            UMCOL           =   0   'False
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   -1  'True
            FX              =   0
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdDelete 
            Height          =   375
            Left            =   3240
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   240
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "&Delete"
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
            MCOL            =   9469052
            MPTR            =   99
            MICON           =   "Supplier.frx":38BA
            UMCOL           =   0   'False
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   -1  'True
            FX              =   0
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdFind 
            Height          =   375
            Left            =   4200
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   240
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "&Find"
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
            MCOL            =   9469052
            MPTR            =   99
            MICON           =   "Supplier.frx":3BD4
            UMCOL           =   0   'False
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   -1  'True
            FX              =   0
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Fax"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   350
         Index           =   5
         Left            =   2880
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   3480
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Phone"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   350
         Index           =   2
         Left            =   2880
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Address"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   825
         Index           =   1
         Left            =   2880
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1800
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         DataField       =   "SupplierID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   350
         Index           =   0
         Left            =   2880
         MaxLength       =   4
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         DataField       =   "SupplierName"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   350
         Index           =   4
         Left            =   2880
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   1200
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   5
         Left            =   1440
         TabIndex        =   10
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   9
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   8
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier ID:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   6
         Top             =   1200
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents rs As Recordset
Attribute rs.VB_VarHelpID = -1

Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean


Private Sub cmdFind_Click()
On Error Resume Next
Dim spname As String
spname = InputBox("Enter Supplier ID to Search?", "Find Supplier")
If spname = "" Then Exit Sub

rs.MoveFirst
rs.Find "SupplierID='" & spname & "'"

If rs.EOF Or rs.BOF Then
    MsgBox "Sorry! No record Found.", vbInformation
    rs.MoveLast
End If

End Sub

Private Sub Form_Load()
On Error Resume Next
  'OpenDB
  cn.CursorLocation = adUseClient
  
  Set rs = New Recordset
  rs.Open "select * from Suppliers", cn, adOpenStatic, adLockOptimistic
  
  Dim otext As TextBox
  'Bind the text boxes to the data provider
  For Each otext In Me.Text1
    Set otext.DataSource = rs
  Next
  mbDataChanged = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      cmdClose_Click
    Case vbKeyEnd
      cmdLast_Click
    Case vbKeyHome
      cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
      Else
        cmdPrev_Click
      End If
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
        cmdLast_Click
      Else
        cmdNext_Click
      End If
  End Select
End Sub

Private Sub Form_Resize()
CenterCtrl Me, Frame1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
  Set rs = Nothing
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  With rs
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    
    'Call GenSeq(rs, Me.Text1(0))
    .AddNew
    mbAddNewFlag = True
    SetButtons False
   
  End With

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  
  If MsgBox("Are you sure to delete this record?", vbYesNo) = vbYes Then
    With rs
        .Delete
        .MoveNext
        If .EOF Then .MoveLast
    End With
    Exit Sub
  Else
    Exit Sub
  End If
  
DeleteErr:
  MsgBox "You cannot Delete this Record Because it has Reference!", vbInformation
  rs.CancelUpdate
  Exit Sub
  
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr
 
  mbEditFlag = True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  rs.CancelUpdate
  If mvBookMark > 0 Then
    rs.Bookmark = mvBookMark
  Else
    rs.MoveFirst
  End If
  mbDataChanged = False

End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr
If Me.Text1(1) = "" Or Me.Text1(2) = "" Then
MsgBox "Incomplete Record....Transaction Failed!", vbInformation
Me.Text1(1).SetFocus
Exit Sub
Else
  rs.UpdateBatch adAffectAll
  MsgBox "Transaction Successfull...Records Updated!", vbInformation
  
  If mbAddNewFlag Then
    rs.MoveLast              'move to the new record
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False

  Exit Sub
End If

UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  rs.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
  Resume Next
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  rs.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
  Resume Next
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError

  If Not rs.EOF Then rs.MoveNext
  If rs.EOF And rs.RecordCount > 0 Then
    Beep
     'moved off the end so go back
    rs.MoveLast
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrev_Click()
  On Error GoTo GoPrevError

  If Not rs.BOF Then rs.MovePrevious
  If rs.BOF And rs.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    rs.MoveFirst
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bval As Boolean)
  cmdAdd.Visible = bval
  cmdEdit.Visible = bval
  cmdUpdate.Visible = Not bval
  cmdCancel.Visible = Not bval
 
  cmdDelete.Visible = bval
  cmdFind.Visible = bval
  cmdClose.Visible = bval
  cmdNext.Enabled = bval
  cmdFirst.Enabled = bval
  cmdLast.Enabled = bval
  cmdPrev.Enabled = bval
End Sub

