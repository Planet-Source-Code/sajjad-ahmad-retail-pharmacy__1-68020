VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "PRJCHAMELEON.OCX"
Object = "{23F895D7-45A6-4886-931B-89D88C2857ED}#1.2#0"; "IGRID250_75B4A91C.OCX"
Begin VB.Form frmCategories 
   Caption         =   "Categories"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4410
   ScaleWidth      =   4710
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin iGrid250_75B4A91C.iGrid iGrid1 
         Height          =   3015
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   5318
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjChameleon.chameleonButton cmdDelete 
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   3720
         Width           =   1095
         _ExtentX        =   1931
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
         MCOL            =   16777215
         MPTR            =   99
         MICON           =   "Categories.frx":0000
         UMCOL           =   0   'False
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdClose 
         Height          =   375
         Left            =   3000
         TabIndex        =   2
         Top             =   3720
         Width           =   1095
         _ExtentX        =   1931
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
         MCOL            =   16777215
         MPTR            =   99
         MICON           =   "Categories.frx":031A
         UMCOL           =   0   'False
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdNew 
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   3720
         Width           =   1095
         _ExtentX        =   1931
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
         MCOL            =   16777215
         MPTR            =   99
         MICON           =   "Categories.frx":0634
         UMCOL           =   0   'False
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdUndo 
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   3720
         Width           =   1095
         _ExtentX        =   1931
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
         MCOL            =   16777215
         MPTR            =   99
         MICON           =   "Categories.frx":094E
         UMCOL           =   0   'False
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdSave 
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   3720
         Width           =   1095
         _ExtentX        =   1931
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
         MCOL            =   16777215
         MPTR            =   99
         MICON           =   "Categories.frx":0C68
         UMCOL           =   0   'False
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Categories"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   120
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim TR As Integer

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
If iGrid1.CurRow < 1 Then Exit Sub
If MsgBox("Are you sure to delete this Category?", vbQuestion + vbYesNo) = vbYes Then
    cn.Execute "Delete * from categories Where Category='" & iGrid1.CellValue(iGrid1.CurRow, 1) & "'"
    iGrid1.RemoveRow iGrid1.CurRow
Else
    Exit Sub
End If
End Sub

Private Sub cmdNew_Click()
iGrid1.RowCount = iGrid1.RowCount + 1
iGrid1.SetCurCell iGrid1.RowCount, 1
SetButtons False
End Sub

Private Sub cmdSave_Click()
If TR = iGrid1.RowCount Then Exit Sub

With rs
For i = TR + 1 To iGrid1.RowCount
    .AddNew
    .Fields(0) = iGrid1.CellValue(i, 1)
    .Fields(1) = iGrid1.CellValue(i, 2)
    .Update
Next i
SetButtons True
End With
End Sub

Private Sub cmdUndo_Click()
If iGrid1.RowCount > TR Then
    iGrid1.RemoveRow iGrid1.RowCount
    SetButtons True
Else
    SetButtons True
End If
End Sub

Private Sub Form_Load()
'OpenDB
rs.Open "Categories", cn, adOpenKeyset, adLockOptimistic
iGrid1.FillFromRS rs
iGrid1.ColWidth(1) = 60
iGrid1.ColWidth(2) = 180
TR = iGrid1.RowCount
SetButtons True
End Sub

Private Sub SetButtons(bval As Boolean)
cmdNew.Visible = bval
cmdClose.Visible = bval
cmdDelete.Visible = bval
cmdSave.Visible = Not bval
cmdUndo.Visible = Not bval
End Sub

Private Sub Form_Resize()
CenterCtrl Me, Frame1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rs = Nothing
End Sub

Private Sub iGrid1_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid250_75B4A91C.ETextEditFlags)
If lCol = 1 Then
    lMaxLength = 6
ElseIf lCol = 2 Then
    lMaxLength = 50
End If
End Sub
