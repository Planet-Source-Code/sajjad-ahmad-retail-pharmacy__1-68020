VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "PRJCHAMELEON.OCX"
Object = "{23F895D7-45A6-4886-931B-89D88C2857ED}#1.2#0"; "IGRID250_75B4A91C.OCX"
Begin VB.Form frmSelSupplier 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "                Select Supplier"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3660
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   3660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin iGrid250_75B4A91C.iGrid iGrid1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   5530
      BackColor       =   -2147483624
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
   Begin prjChameleon.chameleonButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   3360
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "SelSupplier.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   3360
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "SelSupplier.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmSelSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If iGrid1.CurRow = 0 Then Exit Sub
If IsLoaded("frmPorder") Then
    frmPOrder.Text1(2).Text = Me.iGrid1.CellValue(Me.iGrid1.CurRow, 1)
    frmPOrder.Text2.Text = iGrid1.CellValue(iGrid1.CurRow, 2)
    frmPOrder.iGrid1.SetFocus
    Unload Me
ElseIf IsLoaded("frmRefundPurchase") Then
    frmRefundPurchase.Text1(2).Text = Me.iGrid1.CellValue(Me.iGrid1.CurRow, 1)
    frmRefundPurchase.Text2.Text = iGrid1.CellValue(iGrid1.CurRow, 2)
    frmRefundPurchase.iGrid1.SetFocus
    Unload Me
End If
End Sub

Private Sub Form_Load()
rs.Open "Select SupplierID,SupplierName from Suppliers", cn, adOpenForwardOnly, adLockReadOnly
iGrid1.FillFromRS rs

iGrid1.ColWidth(1) = 50
iGrid1.ColWidth(2) = 150
iGrid1.ColHeaderText(1) = "ID"
iGrid1.ColHeaderText(2) = "Supplier Name"
Set rs = Nothing
End Sub

Private Sub iGrid1_GotFocus()
On Error Resume Next
iGrid1.SetCurCell 1, 1
End Sub
