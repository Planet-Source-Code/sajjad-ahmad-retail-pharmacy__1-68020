VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "PRJCHAMELEON.OCX"
Object = "{23F895D7-45A6-4886-931B-89D88C2857ED}#1.2#0"; "IGRID250_75B4A91C.OCX"
Begin VB.Form frmSelCat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Categories"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   2520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin iGrid250_75B4A91C.iGrid iGrid1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   5318
      BackColorOddRows=   12640511
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
   Begin prjChameleon.chameleonButton chameleonButton1 
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3240
      Width           =   975
      _ExtentX        =   1720
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
      MICON           =   "SelCat.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton chameleonButton2 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   3240
      Width           =   975
      _ExtentX        =   1720
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
      MICON           =   "SelCat.frx":001C
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
Attribute VB_Name = "frmSelCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub chameleonButton1_Click()
On Error GoTo sel_err
If IsLoaded("frmproducts") Then
frmProducts.Text1(1).Text = iGrid1.CellValue(iGrid1.CurRow, 1)
frmProducts.Text1(2).SetFocus
Unload Me
End If
sel_err:
    Exit Sub
End Sub

Private Sub chameleonButton2_Click()
Unload Me
End Sub

Private Sub Form_Load()
rs.Open "Select Category from Categories", cn, adOpenForwardOnly, adLockReadOnly
iGrid1.FillFromRS rs
iGrid1.ColWidth(1) = 130
Set rs = Nothing
End Sub

Private Sub iGrid1_GotFocus()
iGrid1.SetCurCell 1, 1
End Sub

