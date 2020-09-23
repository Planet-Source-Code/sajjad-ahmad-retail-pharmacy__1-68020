VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "PRJCHAMELEON.OCX"
Object = "{23F895D7-45A6-4886-931B-89D88C2857ED}#1.2#0"; "IGRID250_75B4A91C.OCX"
Begin VB.Form frmRefundSale 
   Caption         =   "Refund Sale Invoice"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6780
   ScaleWidth      =   11340
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   6495
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10935
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   0
         Left            =   7080
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   8880
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   240
         TabIndex        =   0
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
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
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   5080
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Index           =   5
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   5955
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Height          =   360
         Index           =   4
         Left            =   8520
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "0"
         Top             =   5520
         Width           =   735
      End
      Begin prjChameleon.chameleonButton cmdsave 
         Height          =   495
         Left            =   600
         TabIndex        =   3
         Top             =   5640
         Width           =   1455
         _ExtentX        =   2566
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
         MICON           =   "RefundSale.frx":0000
         PICN            =   "RefundSale.frx":031A
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
         Left            =   2400
         TabIndex        =   4
         Top             =   5640
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "  &Reset"
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
         MICON           =   "RefundSale.frx":0850
         PICN            =   "RefundSale.frx":0B6A
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
         Left            =   4200
         TabIndex        =   5
         Top             =   5640
         Width           =   1455
         _ExtentX        =   2566
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
         MICON           =   "RefundSale.frx":1006
         PICN            =   "RefundSale.frx":1320
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
         Height          =   3735
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   6588
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
            Height          =   3615
            Left            =   0
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   6376
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
            KeyPressBehaviour=   3
            RowMode         =   -1  'True
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Total:"
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
         Left            =   7080
         TabIndex        =   19
         Top             =   5080
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Discount:"
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
         Index           =   5
         Left            =   7080
         TabIndex        =   18
         Top             =   5520
         Width           =   1095
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
         Index           =   6
         Left            =   7080
         TabIndex        =   17
         Top             =   6000
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9360
         TabIndex        =   16
         Top             =   5573
         Width           =   255
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Refund Sale"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   495
         Left            =   3840
         TabIndex        =   15
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Bill No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   7080
         TabIndex        =   14
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   8880
         TabIndex        =   13
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Customer/Patient Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   3615
      End
   End
End
Attribute VB_Name = "frmRefundSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim stock As Integer
'Dim UnitTP As Integer
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdNew_Click()
Call NewMode
Call LoadGrids
Text1(2).SetFocus
cmdNew.Enabled = False
cmdSave.Enabled = True
End Sub

Private Sub cmdSave_Click()
If iGrid1.RowCount = 0 Then Exit Sub
If iGrid1.CellValue(1, 1) = "" Then Exit Sub

If MsgBox("Are you sure to save this record?", vbQuestion + vbYesNo) = vbYes Then
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset

'Add Order Head
rs1.Open "ROrders", cn, adOpenKeyset, adLockOptimistic
With rs1
    .AddNew
    .Fields(0) = Text1(0)
    .Fields(1) = Text1(1)
    .Fields(2) = Text1(2)
    .Fields(3) = Text1(3)
    .Fields(4) = Val(Text1(4) / 100)
    .Fields(5) = Text1(5)
    .Update
End With

'Add Order Details
rs2.Open "ROrderDetails", cn, adOpenKeyset, adLockOptimistic
With rs2
    For i = 1 To iGrid1.RowCount - 1
        .AddNew
        .Fields(0) = Text1(0)
        .Fields(1) = iGrid1.CellValue(i, 1)
        .Fields(2) = iGrid1.CellValue(i, 3)
        .Fields(3) = iGrid1.CellValue(i, 4)
        .Update
    Next
End With

'Update Stock
For i = 1 To iGrid1.RowCount - 1
    cn.Execute "Update Products Set Unit_Stock=Unit_Stock + " & iGrid1.CellValue(i, 3) & " Where Code=" & iGrid1.CellValue(i, 1)
Next

'Update Last Sold Date
'For i = 1 To iGrid1.RowCount - 1
 '   cn.Execute "Update Products Set LastSold=#" & Date & "#" & " Where Code=" & iGrid1.CellValue(i, 1)
'Next

Set rs1 = Nothing
Set rs2 = Nothing

'Question whether to print invoice or not
If MsgBox("Invoice Issued! Do you want to print invoice?", vbQuestion + vbYesNo) = vbYes Then
'Report Generation
Dim rsINV As New ADODB.Recordset
rsINV.Open "Select * from RInvoiceqry where BillNo=" & Me.Text1(0), cn, adOpenStatic, adLockOptimistic

Set InvoiceRPT.DataSource = rsINV

'Head Section
InvoiceRPT.Sections(2).Controls.Item("lblBillNo").Caption = rsINV.Fields("BillNo")
InvoiceRPT.Sections(2).Controls.Item("lbldate").Caption = rsINV.Fields("BillDate")
InvoiceRPT.Sections(2).Controls.Item("lblcustomername").Caption = rsINV.Fields("CustomerName")
InvoiceRPT.Sections(2).Controls.Item("label4").Caption = "REFUND INVOICE"
'Totals Section
InvoiceRPT.Sections(5).Controls.Item("lblsubtotal").Caption = rsINV.Fields("SubTotal")
InvoiceRPT.Sections(5).Controls.Item("lblDiscount").Caption = FormatPercent(rsINV.Fields("Discount"), 0)
InvoiceRPT.Sections(5).Controls.Item("lblgrandtotal").Caption = rsINV.Fields("GrandTotal")

'InvoiceRPT.Show
InvoiceRPT.Hide
InvoiceRPT.PrintReport False, rptRangeAllPages

'Set to new Mode
Call NewMode
Call LoadGrids
Text1(2).SetFocus

Set rsINV = Nothing
Else
'Set to new Mode
    Call NewMode
    Call LoadGrids
    Text1(2).SetFocus
End If

Else
    Exit Sub
End If
End Sub

Private Sub cmdUndo_Click()
If MsgBox("Are you sure to Reset the Invoice?", vbQuestion + vbYesNo) = vbYes Then
'Set to new Mode
Call NewMode
Call LoadGrids
Text1(2).SetFocus
Else
    Exit Sub
End If
End Sub

Private Sub Form_Load()
'OpenDB
'New Mode Clearing Grid & Textboxes
Call NewMode
'Product Selection grid
LoadGrids

End Sub

Private Sub Form_Resize()
CenterCtrl Me, Frame1
End Sub

Private Sub iGrid1_GotFocus()
If iGrid1.RowCount = 1 And iGrid1.CellValue(1, 1) = "" Then
    iGrid1.SetCurCell 1, 1
End If
End Sub

Private Sub iGrid2_GotFocus()
iGrid2.SetCurCell 1, 1
End Sub

Private Sub Text1_Change(Index As Integer)
If Index = 2 Then
    Text1(2) = UCase(Text1(2))
    Text1(2).SelStart = Len(Text1(2))
End If

If Text1(4) = "" Then Exit Sub
If Index = 4 And Val(Text1(4)) > 99 Then
    MsgBox "Discount must be less than 99%", vbInformation
    Text1(4).Text = 0
    Text1(4).SetFocus
Else
    Text1(5) = Format(Val(Text1(3)) - (Val(Text1(3) * Val(Text1(4)) / 100)), "0.00")
End If

End Sub

Private Sub Text1_GotFocus(Index As Integer)
If Index = 4 Then
Text1(4).SelStart = 0
Text1(4).SelLength = Len(Text1(4))
End If
End Sub

Private Sub NewMode()
rs.Open "Select Billno from ROrders", cn, adOpenKeyset, adLockOptimistic

If rs.EOF Or rs.BOF Then
Me.Text1(0).Text = 1
Else
rs.MoveLast
Text1(0).Text = rs.Fields(0) + 1
End If

Text1(1) = Date
Text1(2) = ""
Text1(3) = 0
Text1(4) = 0
Text1(5) = 0
iGrid1.Clear True
Set rs = Nothing
End Sub
Private Sub LoadGrids()
With iGrid1
    .AddCol "Code", "Code", igTextCenter, , 70
    .AddCol "Product", "Product", , , 300
    .AddCol "Qty", "Qty", igTextCenter, , 60
    .AddCol "Rate", "Rate", igTextRight, , 80
    .AddCol "Line Total", "Line Total", igTextRight, , 150
    .RowCount = 1
    
    .Header.Flat = True
    .Header.Buttons = False
    .Header.HotTrack = False
    .Header.BackColor = &H800080
    .Header.ForeColor = vbWhite
    .Header.Font = "Tahoma"
    .Header.Font.Bold = True
    .Header.Font.Size = 10
    .Header.AutoHeight

    
End With

'Filling the Product Grid
iGrid2.Clear True
rs.Open "Select Product,Category,Unit_RP,Unit_Stock,Code,Unit_TP from Products Where Unit_Stock>0 Order By Product", cn, adOpenStatic, adLockOptimistic
iGrid2.FillFromRS rs
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
    .ColHeaderText(3) = "Unit Price"
    .ColHeaderText(4) = "Stock"
    .ColHeaderText(5) = "Code"
    
    .ColWidth(1) = 250
    .ColWidth(2) = 60
    .ColWidth(3) = 80
End With
Set rs = Nothing

End Sub

Private Sub iGrid1_AfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
If lCol = 3 Or lCol = 4 Then
iGrid1.CellValue(iGrid1.CurRow, 5) = Format(iGrid1.CellValue(iGrid1.CurRow, 3) * iGrid1.CellValue(iGrid1.CurRow, 4), "0.00")
CalculateTotal
If iGrid1.CurRow = iGrid1.RowCount Then
iGrid1.RowCount = iGrid1.RowCount + 1
iGrid1.SetCurCell iGrid1.RowCount, 1
End If
End If
End Sub

Private Sub iGrid1_BeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid250_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)

If lCol = 3 And vNewValue = 0 Or vNewValue = "" Then
    eResult = igEditResProceed
'ElseIf lCol = 3 And vNewValue > stock Then
  '  MsgBox "Quantity cannot be greater than stock", vbInformation
  '  eResult = igEditResProceed
End If
End Sub

Private Sub iGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
'If iGrid1.RowCount = 1 And iGrid1.CellValue(1, 1) = "" Then Exit Sub
If iGrid1.CellValue(iGrid1.CurRow, 1) = "" Then Exit Sub
If MsgBox("Are you sure to delete this row?", vbQuestion + vbYesNo) = vbYes Then
    If iGrid1.RowCount = 1 Or iGrid1.CurRow = iGrid1.RowCount Then
        iGrid1.RemoveRow iGrid1.CurRow
        iGrid1.RowCount = iGrid1.RowCount + 1
        iGrid1.SetCurCell iGrid1.RowCount, 1
        CalculateTotal
    Else
        iGrid1.RemoveRow iGrid1.CurRow
        iGrid1.SetCurCell iGrid1.RowCount, 1
        CalculateTotal
    End If
Else
    Exit Sub
End If
End If
End Sub

Private Sub iGrid1_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid250_75B4A91C.ETextEditFlags)
If lCol = 1 Then
    bCancel = True
    iGrid2.Visible = True
    iGrid2.SetFocus
    iGrid2.SetCurCell 1, 1
ElseIf lCol = 2 Or lCol = 4 Or lCol = 5 Then
    bCancel = True
ElseIf lCol = 3 Then
   eTextEditOpt = igTextEditNumberOnly
End If
End Sub

Private Sub iGrid2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then

For i = 1 To iGrid1.RowCount
If iGrid1.CellValue(i, 1) = iGrid2.CellValue(iGrid2.CurRow, 5) Then
    MsgBox "Product Cannot be Duplicated, It is Already Listed in Order!", vbInformation
    iGrid2.SetFocus
    Exit Sub
End If
Next
    iGrid1.CellValue(iGrid1.CurRow, 1) = iGrid2.CellValue(iGrid2.CurRow, 5)
    iGrid1.CellValue(iGrid1.CurRow, 2) = iGrid2.CellValue(iGrid2.CurRow, 1) & " " & iGrid2.CellValue(iGrid2.CurRow, 2)
    iGrid1.CellValue(iGrid1.CurRow, 4) = iGrid2.CellValue(iGrid2.CurRow, 3)
    CellAlign1
    iGrid1.SetFocus
    iGrid2.Visible = False
    iGrid1.SetCurCell iGrid1.CurRow, 3
    iGrid1.RequestEdit vbKeyReturn
ElseIf KeyAscii = vbKeyEscape Then
    iGrid1.SetFocus
    iGrid2.Visible = False
End If
End Sub

Private Sub CellAlign1()
With iGrid1
For i = 1 To .RowCount
    .CellTextFlags(i, 3) = igTextCenter
    .CellTextFlags(i, 4) = igTextRight
    .CellTextFlags(i, 5) = igTextRight
    .CellFmtString(i, 4) = "0.00"
    .CellFmtString(i, 5) = "0.00"
Next
End With
End Sub
Private Sub CheckPriceStock()
rs.Open "select * from products where code=" & iGrid1.CellValue(iGrid1.CurRow, 1), cn, adOpenForwardOnly, adLockReadOnly
'UnitTP = rs.Fields("Unit_TP")
stock = rs.Fields("unit_stock")
Set rs = Nothing
End Sub

Private Sub CalculateTotal()
Text1(3).Text = 0
For i = 1 To iGrid1.RowCount
Text1(3) = Format(Val(Text1(3)) + Val(iGrid1.CellValue(i, 5)), "0.00")
Next
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 4 Then
Select Case KeyAscii
Case 46 To 57
    Exit Sub
Case 8
    DoEvents
    Exit Sub
Case Else
    KeyAscii = 0
End Select

ElseIf Index = 2 And KeyAscii = vbKeyReturn Then
    iGrid1.SetFocus
End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
If Index = 4 And Text1(4).Text = "" Then
Text1(4).Text = 0
End If
End Sub

