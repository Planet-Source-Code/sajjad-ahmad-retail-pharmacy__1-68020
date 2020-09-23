VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "PRJCHAMELEON.OCX"
Begin VB.MDIForm MainMDI 
   BackColor       =   &H8000000C&
   Caption         =   "Pharmacy Managment System - Ver 1.0"
   ClientHeight    =   6030
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9705
   Icon            =   "MainMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   380
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   9705
      TabIndex        =   0
      Top             =   0
      Width           =   9705
      Begin prjChameleon.chameleonButton chameleonButton1 
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   8
         TX              =   "Sale Bill"
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
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   14867671
         MPTR            =   99
         MICON           =   "MainMDI.frx":0CCA
         PICN            =   "MainMDI.frx":0FE4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton chameleonButton2 
         Height          =   375
         Left            =   1215
         TabIndex        =   2
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   8
         TX              =   "Refund Bill"
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
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   14606046
         MPTR            =   99
         MICON           =   "MainMDI.frx":137E
         PICN            =   "MainMDI.frx":1698
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton chameleonButton3 
         Height          =   375
         Left            =   2430
         TabIndex        =   3
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   8
         TX              =   "Medicines"
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
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   14606046
         MPTR            =   99
         MICON           =   "MainMDI.frx":1A32
         PICN            =   "MainMDI.frx":1D4C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton chameleonButton4 
         Height          =   375
         Left            =   3645
         TabIndex        =   4
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   8
         TX              =   "Backup"
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
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   14606046
         MPTR            =   99
         MICON           =   "MainMDI.frx":20E6
         PICN            =   "MainMDI.frx":2400
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton chameleonButton5 
         Height          =   375
         Left            =   4860
         TabIndex        =   5
         Top             =   0
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   8
         TX              =   "Lock System"
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
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   14606046
         MPTR            =   99
         MICON           =   "MainMDI.frx":279A
         PICN            =   "MainMDI.frx":2AB4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton chameleonButton6 
         Height          =   375
         Left            =   6195
         TabIndex        =   6
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   8
         TX              =   "Exit System"
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
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   14606046
         MPTR            =   99
         MICON           =   "MainMDI.frx":2E4E
         PICN            =   "MainMDI.frx":3168
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
   Begin VB.Menu mfile 
      Caption         =   "&File"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu mlock 
         Caption         =   "&System Lock"
         Shortcut        =   {F11}
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu mexit 
         Caption         =   "&Exit System"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu msetup 
      Caption         =   "&Setup"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu mcategory 
         Caption         =   "&Category"
         Shortcut        =   ^C
      End
      Begin VB.Menu mproducts 
         Caption         =   "&Medicine"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnurefresh 
         Caption         =   "&Refresh Stock"
         Shortcut        =   ^R
      End
      Begin VB.Menu lop3 
         Caption         =   "-"
      End
      Begin VB.Menu msuppliers 
         Caption         =   "&Suppliers"
         Shortcut        =   {F3}
      End
      Begin VB.Menu lop2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mpricelist 
         Caption         =   "&Price List"
         Shortcut        =   ^{F2}
         Visible         =   0   'False
      End
      Begin VB.Menu mconfiguration 
         Caption         =   "&Store Information"
         Shortcut        =   {F4}
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mtransaction 
      Caption         =   "&Transactions"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu msaleinv 
         Caption         =   "&Sale Invoice"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mrefund 
         Caption         =   "&Return Medicine"
         Shortcut        =   {F6}
      End
      Begin VB.Menu line101 
         Caption         =   "-"
      End
      Begin VB.Menu mpin 
         Caption         =   "&Purchase Medicine"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mretcomp 
         Caption         =   "Return to &Company"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mreports 
      Caption         =   "&Reports"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu mdailysale 
         Caption         =   "&Daily Sale Summary"
      End
      Begin VB.Menu mdrefunrpt 
         Caption         =   "D&aily Return Summary"
      End
      Begin VB.Menu mstk 
         Caption         =   "&Stock"
         Begin VB.Menu mstkrpt 
            Caption         =   "&Current Stock Report"
         End
         Begin VB.Menu mreorder 
            Caption         =   "&Auto Demand Generator"
         End
         Begin VB.Menu mpending 
            Caption         =   "&Pending Medicine List"
         End
      End
      Begin VB.Menu mpurrpt 
         Caption         =   "&Purchase"
         Begin VB.Menu mduppbill 
            Caption         =   "&Duplicate Purchase Bill"
         End
         Begin VB.Menu mdetpbill 
            Caption         =   "D&etail of Purchase by Date"
         End
         Begin VB.Menu line002 
            Caption         =   "-"
         End
         Begin VB.Menu mdetamed 
            Caption         =   "Detail of &A Medicine Purchase"
         End
      End
      Begin VB.Menu msalerpt 
         Caption         =   "S&ale"
         Begin VB.Menu mdupsbill 
            Caption         =   "&Duplicate of Sale Bill"
         End
         Begin VB.Menu mdetsbill 
            Caption         =   "D&etail of Sale By Date"
         End
         Begin VB.Menu line001 
            Caption         =   "-"
         End
         Begin VB.Menu mprowise 
            Caption         =   "&Productwise Sales Report"
         End
         Begin VB.Menu mhistoryms 
            Caption         =   "&History of Medicine Sale"
         End
      End
      Begin VB.Menu mretsale 
         Caption         =   "&Return Sale"
         Begin VB.Menu mdupretbill 
            Caption         =   "&Duplicate of Return Bill"
         End
         Begin VB.Menu mdetretbill 
            Caption         =   "D&etail of Return Medicine by Date"
         End
         Begin VB.Menu line003 
            Caption         =   "-"
         End
         Begin VB.Menu mdupretcomp 
            Caption         =   "&Duplicate of Return to Company"
         End
      End
      Begin VB.Menu line4 
         Caption         =   "-"
      End
      Begin VB.Menu mexpiryrpt 
         Caption         =   "&Expiry Report"
      End
      Begin VB.Menu mcashbala 
         Caption         =   "&Cash Balance"
      End
   End
   Begin VB.Menu mutilities 
      Caption         =   "&Utilities"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu mbackup 
         Caption         =   "&Backup Database"
         Shortcut        =   {F9}
      End
      Begin VB.Menu line9 
         Caption         =   "-"
      End
      Begin VB.Menu mnotepad 
         Caption         =   "&Notepad"
      End
      Begin VB.Menu mcalc 
         Caption         =   "&Calculator"
      End
   End
   Begin VB.Menu mhelp 
      Caption         =   "&Help"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu mphelp 
         Caption         =   "&Pharmacy Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu line10 
         Caption         =   "-"
      End
      Begin VB.Menu mabout 
         Caption         =   "About Pharmacy"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "MainMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chameleonButton1_Click()
frmOrder.Show
End Sub

Private Sub chameleonButton2_Click()
frmRefundSale.Show
End Sub

Private Sub chameleonButton3_Click()
frmProducts.Show
End Sub

Private Sub chameleonButton4_Click()
frmBackup.Show 1
End Sub

Private Sub chameleonButton5_Click()
frmLock.Show 1
End Sub

Private Sub chameleonButton6_Click()
Unload Me
End Sub

Private Sub mabout_Click()
frmAbout.Show
End Sub

Private Sub mbackup_Click()
frmBackup.Show 1
End Sub

Private Sub mbklog_Click()
frmDateSelection.Show
End Sub

Private Sub mcalc_Click()
On Error GoTo errHandle
    Dim a As Double
    a = Shell("C:\Windows\calc.exe", vbNormalFocus)
    Exit Sub
errHandle:
    MsgBox "Unable to run Calculator Utility on your computer", vbInformation, "Error in opening!!!"
    Resume Next
End Sub

Private Sub mcashbal_Click()
frmCashBalance.Show
End Sub

Private Sub mcashbala_Click()
frmCashBalance.Show
End Sub

Private Sub mcategory_Click()
frmCategories.Show
End Sub

Private Sub mconfiguration_Click()
frmSetup.Show
End Sub

Private Sub mdailysale_Click()
Dim rs As New ADODB.Recordset
rs.Open "Select * from dailysalesummary where billdate=#" & Date & "#", cn, adOpenKeyset, adLockOptimistic

If rs.EOF Or rs.BOF Then
    MsgBox "No data for report. Cancelling....", vbInformation
    Exit Sub
Else
    Set DailySale.DataSource = rs
    DailySale.Sections(2).Controls("Label9").Caption = "For: " & Format$(Date, "dd-mmm-yyyy")
    DailySale.Show
End If
Set rs = Nothing
End Sub

Private Sub mdetamed_Click()
frmSelProduct.Show
frmSelProduct.Text1 = "SingleProPur"
End Sub

Private Sub mdetpbill_Click()
frmDateSelection.Show
frmDateSelection.Option2.Visible = True
frmDateSelection.Option2.Value = True
End Sub

Private Sub mdetretbill_Click()
frmDateSelection.Show
frmDateSelection.Option3.Visible = True
frmDateSelection.Option3.Value = True
End Sub

Private Sub mdetsbill_Click()
frmDateSelection.Show
frmDateSelection.Option1.Visible = True
frmDateSelection.Option1.Value = True
End Sub

Private Sub MDIForm_Load()
'Refresh Pack Stock
Refresh_Stock

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
If MsgBox("Are you sure to Exit System?", vbQuestion + vbYesNo) = vbYes Then
    Dim Form As Form
    For Each Form In Forms
        Unload Form
        Set Form = Nothing
    Next Form
    cn.Close
    End
Else
    Cancel = True
End If
End Sub
Private Sub mdrefunrpt_Click()
Dim rs As New ADODB.Recordset
rs.Open "Select * from DailyReturnSummary where Billdate= #" & Date & "#", cn, adOpenKeyset, adLockOptimistic

If rs.BOF Or rs.EOF Then
    MsgBox "No Data for this Report. Canceling...", vbInformation
    Exit Sub
Else
    Set DailySale.DataSource = rs
    DailySale.Sections(2).Controls("Label8").Caption = "Daily Summary of Return Medicine"
    DailySale.Sections(2).Controls("Label9").Caption = "For: " & Format$(Date, "dd-mmm-yyyy")
    DailySale.Show
End If

Set rs = Nothing

End Sub

Private Sub mduppbill_Click()
frmDupSelection.Show
frmDupSelection.Option2.Visible = True
frmDupSelection.Option2.Value = True
frmDupSelection.GetPOrders
End Sub

Private Sub mdupretbill_Click()
frmDupSelection.Show
frmDupSelection.Option3.Visible = True
frmDupSelection.Option3.Value = True
frmDupSelection.GetROrders
End Sub

Private Sub mdupretcomp_Click()
frmDupSelection.Show
frmDupSelection.Option4.Visible = True
frmDupSelection.Option4.Value = True
frmDupSelection.GetRPOrders
End Sub

Private Sub mdupsbill_Click()
frmDupSelection.Show
frmDupSelection.Option1.Visible = True
frmDupSelection.Option1.Value = True
frmDupSelection.GetOrders
End Sub

Private Sub mexit_Click()
Unload Me
End Sub

Private Sub mexpiryrpt_Click()
frmExpirySold.Show
frmExpirySold.Option1.Visible = True
frmExpirySold.Option1.Value = True
End Sub

Private Sub mhistoryms_Click()
frmSelProduct.Show
frmSelProduct.Text1.Text = "History"
End Sub

Private Sub mlock_Click()
frmLock.Show 1
End Sub

Private Sub mnotepad_Click()
On Error GoTo errHandle
    Dim a As Double
    a = Shell("c:\windows\notepad.exe", vbNormalFocus)
    Exit Sub
errHandle:
    MsgBox "Unable to run Notepad Utility on your computer", vbInformation, "Error in opening!!!"
    Resume Next
End Sub

Private Sub mnurefresh_Click()
Call Refresh_Stock
End Sub

Private Sub mpending_Click()
frmExpirySold.Show
frmExpirySold.Option2.Visible = True
frmExpirySold.Option2.Value = True
End Sub

Private Sub mpin_Click()
frmPOrder.Show
End Sub

Private Sub mpricelist_Click()
frmPriceList.Show
End Sub

Private Sub mproducts_Click()
frmProducts.Show
End Sub

Private Sub mprowise_Click()
Dim rs As New ADODB.Recordset
rs.Open "Select * from Productwisesale", cn, adOpenForwardOnly, adLockOptimistic

If rs.EOF Or rs.BOF Then
    MsgBox "No Data for this report. Canceling ....", vbInformation
    Exit Sub
Else
    Set ProductwiseSale.DataSource = rs
    ProductwiseSale.Sections(2).Controls.Item("label2").Caption = Format(Date, "dd/mm/yyyy")
    ProductwiseSale.Show
End If

Set rs = Nothing

End Sub

Private Sub mrefund_Click()
frmRefundSale.Show
End Sub

Private Sub mreorder_Click()
Dim rs As New ADODB.Recordset
rs.Open "Select * from AutoDemand", cn, adOpenKeyset, adLockOptimistic

If rs.EOF Or rs.BOF Then
    MsgBox "No Data for this report. Canceling ....", vbInformation
    Exit Sub
Else
    Set AutoDemand.DataSource = rs
    AutoDemand.Show
End If

Set rs = Nothing
End Sub

Private Sub mrepair_Click()
'On Error GoTo Repair_Error

    'RepairDatabase App.Path & "\Pharmacy.mdb"
    'Screen.MousePointer = 0
    'MsgBox "Database repaired successfully", 64, "Repair"
    
'Repair_Error:
    'MsgBox "Error when repairing database", 16, "Error"
    'Screen.MousePointer = 0
    'Exit Sub
End Sub

Private Sub mretcomp_Click()
frmRefundPurchase.Show
End Sub

Private Sub msaleinv_Click()
frmOrder.Show
End Sub

Private Sub mstkrpt_Click()
Dim rs As New ADODB.Recordset
rs.Open "Select * from CurrentStock", cn, adOpenKeyset, adLockOptimistic

If rs.BOF Or rs.EOF Then
    MsgBox "No Data for this Report. Canceling...", vbInformation
    Exit Sub
End If

Set CurrentStock.DataSource = rs
CurrentStock.Sections(2).Controls.Item("label2").Caption = Now
CurrentStock.Show

Set rs = Nothing

End Sub

Private Sub msuppliers_Click()
frmSupplier.Show
End Sub

Public Sub Refresh_Stock()
On Error Resume Next
cn.Execute "Update Products Set Pack_Stock=Unit_Stock/Packing "
End Sub

