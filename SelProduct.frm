VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "PRJCHAMELEON.OCX"
Object = "{23F895D7-45A6-4886-931B-89D88C2857ED}#1.2#0"; "IGRID250_75B4A91C.OCX"
Begin VB.Form frmSelProduct 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Product Selection"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin iGrid250_75B4A91C.iGrid iGrid1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   5106
      BackColorOddRows=   12648384
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
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3600
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
   End
   Begin prjChameleon.chameleonButton Command1 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   3120
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
      MICON           =   "SelProduct.frx":0000
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
      Default         =   -1  'True
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   255
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "SelProduct.frx":031A
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
Attribute VB_Name = "frmSelProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
If Text1.Text = "SingleProPur" Then
    rs.Open "Select * from PurchasebyDate Where Product='" & iGrid1.CellValue(iGrid1.CurRow, 1) & "'", cn, adOpenForwardOnly, adLockReadOnly
    If rs.BOF Or rs.EOF Then
        MsgBox "No Data for this Product! Canceling...", vbInformation
        Set rs = Nothing
        Exit Sub
    Else
        Set SingleProPur.DataSource = rs
        SingleProPur.Sections(2).Controls.Item("label9").Caption = iGrid1.CellValue(iGrid1.CurRow, 1)
        SingleProPur.Show
        Set rs = Nothing
    End If
ElseIf Text1.Text = "History" Then
    rs.Open "Select * from SalebyDate Where Product='" & iGrid1.CellValue(iGrid1.CurRow, 1) & "'", cn, adOpenForwardOnly, adLockReadOnly
    If rs.EOF Or rs.BOF Then
        MsgBox "No Data for this Report! Canceling...", vbInformation
        Set rs = Nothing
        Exit Sub
    Else
    Set OrdersBacklog.DataSource = rs
    OrdersBacklog.Sections(2).Controls.Item("Label8").Caption = "History of a Product Sale"
    OrdersBacklog.Sections(2).Controls.Item("Label9").Caption = Format(Date, "dd/mm/yyyy")
    OrdersBacklog.Sections(5).Controls.Item("Function2").Visible = True
    OrdersBacklog.Show
    Set rs = Nothing
    End If
Else
    Exit Sub
End If

End Sub

Private Sub Form_Load()
rs.Open "Select Product,Category,Unit_Stock,Code from Products Order By Product", cn, adOpenStatic, adLockOptimistic
iGrid1.FillFromRS rs
With iGrid1
    .ColWidth(1) = 170
    .ColWidth(2) = 40
    .ColWidth(3) = 40
    .ColVisible(4) = False
    
    .ColHeaderText(2) = "Type"
    .ColHeaderText(3) = "Stock"
'For i = 1 To .RowCount
    '.CellTextFlags(i, 3) = igTextRight
    '.CellTextFlags(i, 4) = igTextRight
    '.CellBackColor(i, 2) = vbYellow
    '.CellBackColor(i, 3) = vbCyan
    '.CellBackColor(i, 4) = vbMagenta
'Next
End With
Set rs = Nothing
End Sub

Private Sub iGrid1_GotFocus()
iGrid1.SetCurCell 1, 1
End Sub
