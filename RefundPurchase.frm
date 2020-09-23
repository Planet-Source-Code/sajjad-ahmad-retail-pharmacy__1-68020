VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "PRJCHAMELEON.OCX"
Object = "{23F895D7-45A6-4886-931B-89D88C2857ED}#1.2#0"; "IGRID250_75B4A91C.OCX"
Begin VB.Form frmRefundPurchase 
   Caption         =   "Refund Purchase Invoice"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6945
   ScaleWidth      =   10080
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   5760
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   9120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   0
         Left            =   7680
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   480
         Width           =   2895
      End
      Begin prjChameleon.chameleonButton cmdsave 
         Height          =   495
         Left            =   2400
         TabIndex        =   6
         Top             =   5640
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "  &Save"
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
         MICON           =   "RefundPurchase.frx":0000
         PICN            =   "RefundPurchase.frx":031A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdundo 
         Height          =   495
         Left            =   3840
         TabIndex        =   7
         Top             =   5640
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "  &Undo"
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
         MICON           =   "RefundPurchase.frx":0850
         PICN            =   "RefundPurchase.frx":0B6A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdclose 
         Height          =   495
         Left            =   5280
         TabIndex        =   8
         Top             =   5640
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "   &Close"
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
         MICON           =   "RefundPurchase.frx":1006
         PICN            =   "RefundPurchase.frx":1320
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin iGrid250_75B4A91C.iGrid iGrid1 
         Height          =   4575
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   8070
         BorderStyle     =   1
         DefaultRowHeight=   18
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         KeyPressBehaviour=   0
         Begin iGrid250_75B4A91C.iGrid iGrid2 
            Height          =   4455
            Left            =   0
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   7858
            BackColor       =   -2147483624
            BorderStyle     =   0
            DefaultRowHeight=   18
            Editable        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
            GridLinesExtend =   1
            KeyPressBehaviour=   3
            RowMode         =   -1  'True
         End
      End
      Begin prjChameleon.chameleonButton cmdnewpro 
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   5640
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Reload Products"
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
         MICON           =   "RefundPurchase.frx":17DD
         PICN            =   "RefundPurchase.frx":1AF7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         Caption         =   "Supplier Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Bill No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   7680
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Return to Company"
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
         Left            =   3960
         TabIndex        =   14
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Grand Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Left            =   6840
         TabIndex        =   13
         Top             =   5760
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   9120
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmRefundPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rspro As New ADODB.Recordset
Dim rs As New ADODB.Recordset

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdnewpro_Click()
LoadProducts
End Sub

Private Sub cmdSave_Click()
If Text1(2).Text = "" Then
    MsgBox "Please Select Supplier Information", vbInformation, "Supplier Error"
    Text1(2).SetFocus
    Exit Sub
ElseIf iGrid1.CellValue(1, 1) = "" Or Val(iGrid1.CellValue(1, 9)) = 0 Then
    MsgBox "Invalid Purchase Detail Entries!", vbInformation
    iGrid1.SetFocus
    iGrid1.SetCurCell 1, 1
    Exit Sub
ElseIf MsgBox("Are you sure to save invoice?", vbQuestion + vbYesNo) = vbYes Then
    Dim rsPO As New ADODB.Recordset
    Dim rsPOD As New ADODB.Recordset
    
    'Update PORDERS Table
    rsPO.Open "RPorders", cn, adOpenKeyset, adLockOptimistic
    With rsPO
        .AddNew
        .Fields(0) = Text1(0)
        .Fields(1) = Text1(1)
        .Fields(2) = Text1(2)
        .Fields(3) = Text1(3)
        .Update
    End With
    'Update Purchase Order Details Table
    rsPOD.Open "RPorderDetails", cn, adOpenKeyset, adLockOptimistic
    With rsPOD
        For i = 1 To iGrid1.RowCount - 1
        .AddNew
        .Fields(0) = Text1(0)
        .Fields(1) = iGrid1.CellValue(i, 1)
        .Fields(2) = iGrid1.CellValue(i, 5)
        .Fields(3) = iGrid1.CellValue(i, 6)
        .Fields(4) = iGrid1.CellValue(i, 7)
        .Fields(5) = iGrid1.CellValue(i, 8) / 100
        .Update
        Next
    End With
    
    'Update Product Table Unit_Stock
    With iGrid1
    For i = 1 To .RowCount - 1
        cn.Execute "Update Products Set Pack_stock=pack_stock- " & .CellValue(i, 5) + .CellValue(i, 6) & ", Unit_stock=unit_stock- " & .CellValue(i, 5) + .CellValue(i, 6) / .CellValue(i, 4) & " Where Code=" & .CellValue(i, 1)
    Next
    End With

Set rsPO = Nothing
Set rsPOD = Nothing

    'Print Invoice
    If MsgBox("Records Updated Successfully...Do you want to Print?", vbInformation + vbYesNo) = vbYes Then
        'Report Generation
        Dim rsINV As New ADODB.Recordset
        rsINV.Open "Select * from RPInvoiceqry where BillNo=" & Me.Text1(0), cn, adOpenStatic, adLockOptimistic

        Set PInvoiceRPT.DataSource = rsINV

        'Head Section
        PInvoiceRPT.Sections(2).Controls.Item("lblBillNo").Caption = rsINV.Fields("BillNo")
        PInvoiceRPT.Sections(2).Controls.Item("lbldate").Caption = rsINV.Fields("BillDate")
        PInvoiceRPT.Sections(2).Controls.Item("lblsid").Caption = rsINV.Fields("SupplierID")
        PInvoiceRPT.Sections(2).Controls.Item("lblsname").Caption = rsINV.Fields("SupplierName")
        PInvoiceRPT.Sections(2).Controls.Item("label4").Caption = "PURCHASE INVOICE"

        'Totals Section
        PInvoiceRPT.Sections(5).Controls.Item("lblsubtotal").Caption = Format(rsINV.Fields("GrandTotal"), "0.00")

        'PInvoiceRPT.Show
        PInvoiceRPT.Visible = False
        PInvoiceRPT.PrintReport False, rptRangeAllPages
        Set rsINV = Nothing
        Call NewMode
        Call LoadPGrid
        Call LoadProducts
        Text1(2).SetFocus
    Else
        Call NewMode
        Call LoadPGrid
        Call LoadProducts
        Text1(2).SetFocus
    End If
End If
End Sub

Private Sub cmdUndo_Click()
If MsgBox("Are you sure to Reset the Invoice?", vbQuestion + vbYesNo) = vbYes Then
'Set to new Mode
Call NewMode
Call LoadPGrid
Call LoadProducts
Text1(2).SetFocus
Else
    Exit Sub
End If
End Sub

Private Sub Form_Load()
'OpenDB
NewMode
LoadPGrid
LoadProducts
End Sub

Private Sub Form_Resize()
CenterCtrl Me, Frame1
End Sub

Private Sub iGrid1_AfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
With iGrid1
    lRow = .CurRow
    .CellValue(lRow, 9) = Format(.CellValue(lRow, 5) * .CellValue(lRow, 7) - (.CellValue(lRow, 5) * .CellValue(lRow, 7) * (.CellValue(lRow, 8) / 100)), "0.00")
End With
CalculateTotals
End Sub

Private Sub iGrid1_BeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid250_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
If lCol = 5 Then
    If vNewValue = "" Or vNewValue = 0 Then
        eResult = igEditResProceed
        Exit Sub
    Else
        eResult = igEditResCommit
    End If
ElseIf lCol = 8 Then
    If vNewValue > 99 Then
        MsgBox "Discount must be less than 100%", vbInformation
        eResult = igEditResProceed
    End If
End If
End Sub

Private Sub iGrid1_GotFocus()
If iGrid1.RowCount = 1 And iGrid1.CellValue(1, 1) = "" Then
iGrid1.SetCurCell 1, 1
End If
End Sub

Private Sub iGrid1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If iGrid1.CurCol = 1 And KeyAscii = vbKeyReturn Then
    iGrid2.Visible = True
    iGrid2.SetFocus
ElseIf KeyAscii = vbKeySpace Then
    iGrid1.CancelEdit
    iGrid1.SetCurCell iGrid1.RowCount, 1
ElseIf KeyAscii = vbKeyEscape Then
    If iGrid1.CellValue(iGrid1.CurRow, 1) = "" Then Exit Sub
    If MsgBox("Are you sure to delete this row?", vbQuestion + vbYesNo) = vbYes Then
        iGrid1.RemoveRow iGrid1.CurRow
    Else
        Exit Sub
    End If
End If
End Sub

Private Sub iGrid1_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid250_75B4A91C.ETextEditFlags)
'Columns Cell Alignment
Call ColAlign
'Lock Entry Fields
If lCol = 1 Or lCol = 2 Or lCol = 3 Or lCol = 9 Then
    bCancel = True
ElseIf lCol = 5 Then
    eTextEditOpt = igTextEditNumberOnly
End If
End Sub

Private Sub iGrid2_GotFocus()
iGrid2.SetCurCell 1, 1
End Sub

Private Sub iGrid2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    iGrid1.SetFocus
    iGrid1.SetCurCell iGrid1.RowCount, 1
    iGrid2.Visible = False
ElseIf KeyAscii = vbKeyReturn Then
For i = 1 To iGrid1.RowCount
If iGrid1.CellValue(i, 1) = iGrid2.CellValue(iGrid2.CurRow, 6) Then
    MsgBox "Product Cannot be Duplicated, It is Already Listed in Order!", vbInformation
    iGrid2.SetFocus
    Exit Sub
End If
Next
    iGrid1.CellValue(iGrid1.RowCount, 1) = iGrid2.CellValue(iGrid2.CurRow, 6)
    iGrid1.CellValue(iGrid1.RowCount, 2) = iGrid2.CellValue(iGrid2.CurRow, 1)
    iGrid1.CellValue(iGrid1.RowCount, 3) = iGrid2.CellValue(iGrid2.CurRow, 2)
    iGrid1.CellValue(iGrid1.RowCount, 4) = iGrid2.CellValue(iGrid2.CurRow, 3)
    iGrid1.CellValue(iGrid1.RowCount, 5) = 0
    iGrid1.CellValue(iGrid1.RowCount, 6) = 0
    iGrid1.CellValue(iGrid1.RowCount, 7) = iGrid2.CellValue(iGrid2.CurRow, 4)
    iGrid1.CellValue(iGrid1.RowCount, 8) = 0
    iGrid1.SetCurCell iGrid1.RowCount, 5
    iGrid1.RowCount = iGrid1.RowCount + 1
    iGrid1.RequestEdit vbKeyReturn
    iGrid2.Visible = False
End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 2 And KeyAscii = vbKeyReturn Then
frmSelSupplier.Show
End If
End Sub

Private Sub LoadPGrid()
With iGrid1
    .AddCol "Code", "Code", igTextCenter, , 50
    .AddCol "Product", "Product", , , 190
    .AddCol "Type", "Type", igTextCenter, , 55
    .AddCol "Pack", "Pack", igTextCenter, , 50
    .AddCol "Qty", "Qty", igTextCenter, , 50
    .AddCol "Bonus", "Bonus", igTextCenter, , 50
    .AddCol "Rate", "Rate", igTextRight, , 65
    .AddCol "Disc%", "Disc%", igTextCenter, , 52
    .AddCol "Line Total", "Line Total", igTextRight, , 100
    
    .RowCount = 1
    .Header.Flat = True
    .Header.Buttons = False
    .Header.HotTrack = False
    .Header.BackColor = &H808000
    .Header.ForeColor = vbWhite
    .Header.Font = "Tahoma"
    .Header.Font.Bold = True
    .Header.Font.Size = 10
    .Header.AutoHeight
End With
End Sub

Private Sub LoadProducts()
'Filling the Product Grid
iGrid2.Clear True
rspro.Open "Select Product,Category,Packing,Pack_TP,Unit_Stock,Code from Products Order By Product", cn, adOpenForwardOnly, adLockReadOnly
iGrid2.FillFromRS rspro
With iGrid2
    .Header.Flat = True
    .Header.Buttons = False
    .Header.HotTrack = False
    .Header.BackColor = vbBlue
    .Header.ForeColor = vbWhite
    .Header.Font = "Tahoma"
    .Header.Font.Bold = True
    .Header.Font.Size = 10
    .Header.AutoHeight
    
    .ColHeaderText(1) = "Product"
    .ColHeaderText(2) = "Type"
    .ColHeaderText(3) = "Pack"
    .ColHeaderText(4) = "Pack-TP"
    .ColHeaderText(5) = "Stock"
    .ColHeaderText(6) = "Code"
    
    .ColWidth(1) = 250
    .ColWidth(2) = 60
    .ColWidth(4) = 80
End With

Set rspro = Nothing

End Sub

Private Sub ColAlign()
With iGrid1
For i = 1 To iGrid1.RowCount
    .CellTextFlags(i, 1) = igTextCenter
    .CellTextFlags(i, 4) = igTextCenter
    .CellTextFlags(i, 5) = igTextCenter
    .CellTextFlags(i, 6) = igTextCenter
    .CellTextFlags(i, 7) = igTextRight
    .CellTextFlags(i, 8) = igTextRight
    .CellTextFlags(i, 9) = igTextRight
Next
End With
End Sub

Private Sub CalculateTotals()
Text1(3) = 0
For i = 1 To iGrid1.RowCount
    Text1(3) = Format(Val(Text1(3)) + Val(iGrid1.CellValue(i, 9)), "0.00")
Next
End Sub

Private Sub NewMode()
rs.Open "Select Billno from RPOrders", cn, adOpenKeyset, adLockOptimistic

If rs.EOF Or rs.BOF Then
Me.Text1(0).Text = 1
Else
rs.MoveLast
Text1(0).Text = rs.Fields(0) + 1
End If

Text1(1) = Date
Text1(2) = ""
Text2 = ""
Text1(3) = 0
iGrid1.Clear True
Set rs = Nothing
End Sub



