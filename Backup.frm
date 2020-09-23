VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "PRJCHAMELEON.OCX"
Begin VB.Form frmBackup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backup Database"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameCurrBackUp 
      Caption         =   "Choose Path for BackUp"
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
      Height          =   2655
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   6015
      Begin VB.FileListBox File1 
         Height          =   675
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.DirListBox Dir1 
         Height          =   1665
         Left            =   480
         TabIndex        =   9
         Top             =   840
         Width           =   3015
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   480
         TabIndex        =   8
         Top             =   480
         Width           =   3015
      End
      Begin prjChameleon.chameleonButton cmdCancel 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   3960
         TabIndex        =   11
         Top             =   1680
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "  &Cancel"
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
         FCOL            =   4210752
         FCOLO           =   255
         MCOL            =   16777215
         MPTR            =   99
         MICON           =   "Backup.frx":0000
         PICN            =   "Backup.frx":031A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   1
         NGREY           =   0   'False
         FX              =   1
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdSave 
         Height          =   645
         Left            =   3960
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Backup Database"
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1138
         BTYPE           =   3
         TX              =   "  &Backup"
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
         FCOL            =   4210752
         FCOLO           =   255
         MCOL            =   16777215
         MPTR            =   99
         MICON           =   "Backup.frx":076C
         PICN            =   "Backup.frx":0A86
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   1
         NGREY           =   0   'False
         FX              =   1
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Last BackUp Details"
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
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Path:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Time:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblLastPath 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Backup Path"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   720
         TabIndex        =   3
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label lblLastDate 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Backup Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   1320
         Width           =   5175
      End
      Begin VB.Label lblLastTime 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Backup Time"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   1
         Top             =   1680
         Width           =   5175
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Backup Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   4920
      Width           =   6015
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Fsys As New FileSystemObject
Dim bckupFile As File

'Reading Previously Backup Details
Private Sub Form_Load()
    Dim lastPath As String
    Dim lastDate As String
    Dim lastTime As String
  
    'Read Registry for previous settings stored
    lastPath = GetSetting(App.Title, "Settings", "BackupPath")
    lastDate = GetSetting(App.Title, "Settings", "BackupDate")
    lastTime = GetSetting(App.Title, "Settings", "BackupTime")
    
    If lastPath = "" Then
        lblLastPath.Caption = "No Backup made previously"
        lblLastDate.Caption = " "
        lblLastTime.Caption = " "
    Else
        lblLastPath.Caption = lastPath
        lblLastDate.Caption = lastDate & "  (mm-dd-yy)"
        lblLastTime.Caption = lastTime
    End If
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Backup Cmd Btn
Private Sub cmdSave_Click()
On Error Resume Next
If File1.Path = App.Path Then
MsgBox "Please Choose another Destination Path for Backup!", vbInformation
Exit Sub
End If
    cmdSave.Enabled = False
    Label1.Caption = "Please Wait, Backup in Progress..."
    
    Dim destination As String
    Dim Source As String
    Dim currDate, currTime As String
    currDate = Format$(Now, "mm - dd - yy")
    currTime = Format$(Now, "hh:mm:ss AM/PM")
    destination = File1.Path & "\" & "Pharmacy.mdb"
    Source = App.Path & "\SaleManager.mdb"
    'MsgBox "Source : " & source
    'MsgBox "Destination : " & destination
    Set bckupFile = Fsys.GetFile(finalpath)
    bckupFile.Attributes = Compressed
    Fsys.CopyFile Source, destination, True
    'Saving Current Backup Details
    SaveSetting App.Title, "Settings", "BackupPath", destination
    SaveSetting App.Title, "Settings", "BackupDate", currDate
    SaveSetting App.Title, "Settings", "BackupTime", currTime
    
    Label1.Caption = "Backup Process Completed Successfully..."
    cmdSave.Enabled = True
    MsgBox "Baackup Process Completed Successfully!", vbInformation, "Backup"
    Unload Me
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub


