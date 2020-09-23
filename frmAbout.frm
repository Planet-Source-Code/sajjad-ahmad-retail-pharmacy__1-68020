VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "PRJCHAMELEON.OCX"
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   Caption         =   "About Pharmacy"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5535
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   3450
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin prjChameleon.chameleonButton Command2 
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Register"
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
      BCOL            =   14867671
      BCOLO           =   14867671
      FCOL            =   0
      FCOLO           =   255
      MCOL            =   9469052
      MPTR            =   99
      MICON           =   "frmAbout.frx":433CE
      UMCOL           =   0   'False
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton Command1 
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   2880
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
      BCOL            =   14867671
      BCOLO           =   14867671
      FCOL            =   0
      FCOLO           =   255
      MCOL            =   9469052
      MPTR            =   99
      MICON           =   "frmAbout.frx":436E8
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
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sajjad Ahmad  0300-5504930 Email: sajjad@inbox.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   1
      Left            =   2040
      TabIndex        =   8
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "For Help and Quries Contact:"
      BeginProperty Font 
         Name            =   "Verdana"
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
      Left            =   2160
      TabIndex        =   7
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "IMPORTANT INFORMATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1560
      Width           =   5655
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1710
      TabIndex        =   2
      Top             =   885
      Width           =   1275
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pharmacy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   840
      Index           =   1
      Left            =   435
      TabIndex        =   1
      Top             =   0
      Width           =   3465
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2003-2007 Tecknix Solutions. All rights reserved."
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pharmacy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   840
      Index           =   0
      Left            =   435
      TabIndex        =   3
      Top             =   45
      Width           =   3465
   End
   Begin VB.Image Image1 
      Height          =   1125
      Left            =   3915
      Picture         =   "frmAbout.frx":43A02
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1170
   End
   Begin VB.Line Line2 
      DrawMode        =   10  'Mask Pen
      X1              =   435
      X2              =   5115
      Y1              =   1245
      Y2              =   1245
   End
   Begin VB.Line Line1 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   435
      X2              =   5115
      Y1              =   1245
      Y2              =   1245
   End
   Begin VB.Image Image2 
      Height          =   1170
      Left            =   120
      Picture         =   "frmAbout.frx":44BFE
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   1740
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
frmRegister.Show 1
End Sub

