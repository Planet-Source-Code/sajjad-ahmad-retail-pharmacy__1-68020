VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "PRJCHAMELEON.OCX"
Object = "{23F895D7-45A6-4886-931B-89D88C2857ED}#1.2#0"; "IGRID250_75B4A91C.OCX"
Begin VB.Form frmDupSelection 
   Caption         =   "Duplicate Bills Selection"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3345
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
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
      Height          =   2775
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   3135
      Begin iGrid250_75B4A91C.iGrid iGrid1 
         Height          =   2415
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   4260
         BackColorOddRows=   12648447
         BorderStyle     =   1
         Editable        =   0   'False
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
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Refund Purchase Bill"
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
         Left            =   360
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1560
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Refund Sale Bill"
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
         Left            =   360
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1200
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Sale Bill"
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
         Left            =   360
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   480
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Purchase Bill"
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
         Left            =   360
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   2175
      End
   End
   Begin VB.TextBox Text1 
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
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin prjChameleon.chameleonButton Command1 
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   2880
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
      MICON           =   "DuplicateSelection.frx":0000
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
      TabIndex        =   2
      Top             =   2880
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
      MICON           =   "DuplicateSelection.frx":031A
      UMCOL           =   0   'False
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmDupSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Private Sub Command1_Click()
If Me.Text1.Text = "" Or IsNumeric(Me.Text1) = False Then
        MsgBox "Please Enter a Number Value", vbInformation
        Text1.SetFocus
        Exit Sub
End If
    
If Option1.Value = True Then
    Dim rsBill As New ADODB.Recordset
    rsBill.Open "Select * from InvoiceQry where billno=" & Val(Text1.Text), cn, adOpenKeyset, adLockOptimistic

        If rsBill.EOF Or rsBill.BOF Then
            MsgBox "No Record for this Bill Found!", vbInformation
            Exit Sub
        Else
            'Set Data Source
                Set InvoiceRPT.DataSource = rsBill
            'Head Section
            InvoiceRPT.Sections(2).Controls.Item("lblBillNo").Caption = rsBill.Fields("BillNo")
            InvoiceRPT.Sections(2).Controls.Item("lbldate").Caption = rsBill.Fields("BillDate")
            InvoiceRPT.Sections(2).Controls.Item("lblcustomername").Caption = rsBill.Fields("CustomerName")
            InvoiceRPT.Sections(2).Controls.Item("label4").Caption = "DUPLICATE INVOICE"

            'Totals Section
            InvoiceRPT.Sections(5).Controls.Item("lblsubtotal").Caption = rsBill.Fields("SubTotal")
            InvoiceRPT.Sections(5).Controls.Item("lblDiscount").Caption = FormatPercent(rsBill.Fields("Discount"), 0)
            InvoiceRPT.Sections(5).Controls.Item("lblgrandtotal").Caption = rsBill.Fields("GrandTotal")
            
            InvoiceRPT.Show
            Set rsBill = Nothing
            Me.Hide
        End If
        
ElseIf Option2.Value = True Then
    Dim rspBill As New ADODB.Recordset
    rspBill.Open "Select * from pinvoiceQry where billno=" & Val(Text1.Text), cn, adOpenKeyset, adLockOptimistic

        If rspBill.EOF Or rspBill.BOF Then
            MsgBox "No Record for this Bill Found!", vbInformation
            Exit Sub
        Else
            'Set Data Source
            Set PInvoiceRPT.DataSource = rspBill

            'Head Section
            PInvoiceRPT.Sections(2).Controls.Item("lblBillNo").Caption = rspBill.Fields("BillNo")
            PInvoiceRPT.Sections(2).Controls.Item("lbldate").Caption = rspBill.Fields("BillDate")
            PInvoiceRPT.Sections(2).Controls.Item("lblsid").Caption = rspBill.Fields("SupplierID")
            PInvoiceRPT.Sections(2).Controls.Item("lblsname").Caption = rspBill.Fields("SupplierName")
            PInvoiceRPT.Sections(2).Controls.Item("label4").Caption = "DUPLICATE PURCHASE INVOICE"

            'Totals Section
            PInvoiceRPT.Sections(5).Controls.Item("lblsubtotal").Caption = Format(rspBill.Fields("GrandTotal"), "0.00")
            PInvoiceRPT.Show
            Me.Hide
            Set rspBill = Nothing
        End If
        
ElseIf Option3.Value = True Then
    Dim rs1 As New ADODB.Recordset
    rs1.Open "Select * from RInvoiceqry where Billno=" & Val(Text1.Text), cn, adOpenKeyset, adLockOptimistic

        If rs1.EOF Or rs1.BOF Then
            MsgBox "No Record for this Bill Found!", vbInformation
            Exit Sub
        Else
            'Set Data Source
            Set InvoiceRPT.DataSource = rs1
            'Head Section
            InvoiceRPT.Sections(2).Controls.Item("lblBillNo").Caption = rs1.Fields("BillNo")
            InvoiceRPT.Sections(2).Controls.Item("lbldate").Caption = rs1.Fields("BillDate")
            InvoiceRPT.Sections(2).Controls.Item("lblcustomername").Caption = rs1.Fields("CustomerName")
            InvoiceRPT.Sections(2).Controls.Item("label4").Caption = "DUPLICATE REFUND INVOICE"

            'Totals Section
            InvoiceRPT.Sections(5).Controls.Item("lblsubtotal").Caption = rs1.Fields("SubTotal")
            InvoiceRPT.Sections(5).Controls.Item("lblDiscount").Caption = FormatPercent(rs1.Fields("Discount"), 0)
            InvoiceRPT.Sections(5).Controls.Item("lblgrandtotal").Caption = rs1.Fields("GrandTotal")
            
            InvoiceRPT.Show
            Me.Hide
            Set rs1 = Nothing
        End If
ElseIf Option4.Value = True Then
    Dim rs2 As New ADODB.Recordset
    rs2.Open "Select * from RPInvoiceqry where Billno=" & Val(Text1.Text), cn, adOpenKeyset, adLockOptimistic

        If rs2.EOF Or rs2.BOF Then
            MsgBox "No Record for this Bill Found!", vbInformation
            Exit Sub
        Else
            'Set Data Source
            Set PInvoiceRPT.DataSource = rs2

            'Head Section
            PInvoiceRPT.Sections(2).Controls.Item("lblBillNo").Caption = rs2.Fields("BillNo")
            PInvoiceRPT.Sections(2).Controls.Item("lbldate").Caption = rs2.Fields("BillDate")
            PInvoiceRPT.Sections(2).Controls.Item("lblsid").Caption = rs2.Fields("SupplierID")
            PInvoiceRPT.Sections(2).Controls.Item("lblsname").Caption = rs2.Fields("SupplierName")
            PInvoiceRPT.Sections(2).Controls.Item("label4").Caption = "DUPLICATE PURCHASE REFUND INVOICE"

            'Totals Section
            PInvoiceRPT.Sections(5).Controls.Item("lblsubtotal").Caption = Format(rs2.Fields("GrandTotal"), "0.00")
            PInvoiceRPT.Show
            Me.Hide
            Set rs2 = Nothing
        End If
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Public Sub GetOrders()
iGrid1.Clear True
rs.Open "Select BillNo,BillDate from orders", cn, adOpenForwardOnly, adLockReadOnly
iGrid1.FillFromRS rs
iGrid1.ColWidth(1) = 60
iGrid1.ColWidth(2) = 100
Set rs = Nothing
End Sub

Private Sub iGrid1_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)
On Error Resume Next
Text1.Text = iGrid1.CellValue(iGrid1.CurRow, 1)
End Sub

Private Sub iGrid1_GotFocus()
On Error Resume Next
iGrid1.SetCurCell 1, 1
End Sub

Public Sub GetPOrders()
iGrid1.Clear True
rs.Open "Select BillNo,BillDate from Porders", cn, adOpenForwardOnly, adLockReadOnly
iGrid1.FillFromRS rs
iGrid1.ColWidth(1) = 60
iGrid1.ColWidth(2) = 100
Set rs = Nothing
End Sub

Public Sub GetROrders()
iGrid1.Clear True
rs.Open "Select BillNo,BillDate from Rorders", cn, adOpenForwardOnly, adLockReadOnly
iGrid1.FillFromRS rs
iGrid1.ColWidth(1) = 60
iGrid1.ColWidth(2) = 100
Set rs = Nothing
End Sub

Public Sub GetRPOrders()
iGrid1.Clear True
rs.Open "Select BillNo,BillDate from RPorders", cn, adOpenForwardOnly, adLockReadOnly
iGrid1.FillFromRS rs
iGrid1.ColWidth(1) = 60
iGrid1.ColWidth(2) = 100
Set rs = Nothing
End Sub

