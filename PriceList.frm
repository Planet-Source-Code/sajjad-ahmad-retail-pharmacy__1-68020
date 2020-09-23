VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "PRJCHAMELEON.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{23F895D7-45A6-4886-931B-89D88C2857ED}#1.2#0"; "IGRID250_75B4A91C.OCX"
Begin VB.Form frmPriceList 
   Caption         =   "Price List"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6360
   ScaleWidth      =   9060
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Product Price List"
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
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      Begin iGrid250_75B4A91C.iGrid iGrid1 
         Height          =   5775
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   10186
         BackColorOddRows=   16777152
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
         KeyPressBehaviour=   3
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   6420
         Visible         =   0   'False
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin prjChameleon.chameleonButton cmdSave 
         Height          =   495
         Left            =   5160
         TabIndex        =   1
         Top             =   6360
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   " &Save Price List Changes"
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
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   99
         MICON           =   "PriceList.frx":0000
         PICN            =   "PriceList.frx":031A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdClose 
         Height          =   495
         Left            =   8400
         TabIndex        =   2
         Top             =   6360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   " &Exit without Saving"
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
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   99
         MICON           =   "PriceList.frx":0850
         PICN            =   "PriceList.frx":0B6A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
End
Attribute VB_Name = "frmPriceList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim oldval As Variant

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()

If MsgBox("Are you sure to save all changes?", vbQuestion + vbYesNo) = vbYes Then
Me.MousePointer = vbHourglass
Me.ProgressBar1.Visible = True

For i = 1 To iGrid1.RowCount
    cn.Execute "Update Products set Pack_TP=" & iGrid1.CellValue(i, 5) & _
    ", Pack_RP=" & iGrid1.CellValue(i, 6) & ", Unit_TP=" & iGrid1.CellValue(i, 7) & _
    ", Unit_RP=" & iGrid1.CellValue(i, 8) & ", Unit_Stock=" & iGrid1.CellValue(i, 9) & " Where Code=" & iGrid1.CellValue(i, 1)
    
    If Not IsNull(iGrid1.CellValue(i, 10)) Then
    cn.Execute "Update Products Set Expirydate=#" & iGrid1.CellValue(i, 10) & "#" & " Where Code=" & iGrid1.CellValue(i, 1)
    End If

    Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
Next i
    
MsgBox "Records updated Sucessfully", vbInformation
Me.MousePointer = vbDefault
Me.ProgressBar1.Visible = False
Me.cmdClose.SetFocus
Me.cmdSave.Enabled = False
Call LoadPriceList
Else
    Exit Sub
End If

End Sub

Private Sub Form_Load()
'OpenDB

Call LoadPriceList

Me.cmdSave.Enabled = False
Me.ProgressBar1.Min = 0
Me.ProgressBar1.Max = iGrid1.RowCount

End Sub

Private Sub Form_Resize()
CenterCtrl Me, Frame1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rs = Nothing
End Sub

Private Sub iGrid1_AfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
If lCol = 5 Then
    iGrid1.CellValue(iGrid1.CurRow, 6) = iGrid1.CellValue(iGrid1.CurRow, 5) + (iGrid1.CellValue(iGrid1.CurRow, 5) * 0.15)
    iGrid1.CellValue(iGrid1.CurRow, 7) = iGrid1.CellValue(iGrid1.CurRow, 5) / iGrid1.CellValue(iGrid1.CurRow, 4)
    iGrid1.CellValue(iGrid1.CurRow, 8) = iGrid1.CellValue(iGrid1.CurRow, 6) / iGrid1.CellValue(iGrid1.CurRow, 4)
ElseIf lCol = 6 Then
    iGrid1.CellValue(iGrid1.CurRow, 8) = iGrid1.CellValue(iGrid1.CurRow, 6) / iGrid1.CellValue(iGrid1.CurRow, 4)
End If
End Sub

Private Sub iGrid1_BeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid250_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
If lCol = 5 Or lCol = 6 Or lCol = 7 Or lCol = 8 Or lCol = 9 Then
    If IsNumeric(vNewValue) = False Then
        MsgBox "Enetr Numeric Value", vbInformation
        eResult = igEditResProceed
        iGrid1.RequestEdit vbKeyReturn
        Exit Sub
    ElseIf vNewValue <> oldval Then
        For i = 1 To iGrid1.ColCount
            iGrid1.CellBackColor(iGrid1.CurRow, i) = vbGreen
        Next
        Me.cmdSave.Enabled = True
    End If
ElseIf lCol = 10 Then
    If IsDate(vNewValue) = False Then
        MsgBox "Enter a Valid Date!", vbInformation
        eResult = igEditResProceed
        Exit Sub
    Else
        For i = 1 To iGrid1.ColCount
            iGrid1.CellBackColor(iGrid1.CurRow, i) = vbGreen
        Next
        Me.cmdSave.Enabled = True
    End If
End If
End Sub

Private Sub iGrid1_GotFocus()
iGrid1.SetCurCell 1, 2
End Sub

Private Sub iGrid1_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid250_75B4A91C.ETextEditFlags)
oldval = iGrid1.CellValue(iGrid1.CurRow, iGrid1.CurCol)
If lCol < 5 Then
    bCancel = True
Else
    bCancel = False
End If

End Sub

Private Sub LoadPriceList()
rs.Open "Select Code,Product,Category,Packing,Pack_TP,Pack_RP,Unit_TP,Unit_RP,Unit_Stock,ExpiryDate from products", cn, adOpenKeyset, adLockOptimistic
iGrid1.FillFromRS rs
iGrid1.ColWidth(1) = 45
iGrid1.ColWidth(2) = 150
iGrid1.ColWidth(3) = 55
iGrid1.ColWidth(4) = 55
iGrid1.ColWidth(9) = 65
iGrid1.ColWidth(10) = 80
iGrid1.ColHeaderText(10) = "Expiry Date"

iGrid1.Header.Height = 25

Set rs = Nothing
End Sub

