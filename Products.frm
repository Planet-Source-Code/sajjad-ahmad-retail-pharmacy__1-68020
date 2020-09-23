VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "PRJCHAMELEON.OCX"
Object = "{23F895D7-45A6-4886-931B-89D88C2857ED}#1.2#0"; "IGRID250_75B4A91C.OCX"
Begin VB.Form frmProducts 
   Caption         =   "Products"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6195
   ScaleWidth      =   9030
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8055
      Begin VB.Frame Frame4 
         Height          =   1695
         Left            =   3240
         TabIndex        =   24
         Top             =   3120
         Width           =   4815
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            DataField       =   "unit_rp"
            Height          =   325
            Index           =   13
            Left            =   3360
            TabIndex        =   44
            Text            =   "Text1"
            Top             =   1200
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            DataField       =   "unit_st"
            Height          =   325
            Index           =   12
            Left            =   3360
            TabIndex        =   43
            Text            =   "Text1"
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            DataField       =   "unit_tp"
            Height          =   325
            Index           =   11
            Left            =   3360
            TabIndex        =   42
            Text            =   "Text1"
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            DataField       =   "pack_rp"
            Height          =   325
            Index           =   10
            Left            =   1080
            TabIndex        =   41
            Text            =   "Text1"
            Top             =   1200
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            DataField       =   "Pack_st"
            Height          =   325
            Index           =   9
            Left            =   1080
            TabIndex        =   40
            Text            =   "Text1"
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            DataField       =   "pack_tp"
            Height          =   325
            Index           =   8
            Left            =   1080
            TabIndex        =   39
            Text            =   "Text1"
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Unit_RP:"
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
            Index           =   11
            Left            =   2520
            TabIndex        =   30
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Unit_ST:"
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
            Index           =   10
            Left            =   2520
            TabIndex        =   29
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Unit_TP:"
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
            Index           =   9
            Left            =   2520
            TabIndex        =   28
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Pack_RP:"
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
            Index           =   7
            Left            =   240
            TabIndex        =   27
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Pack_ST:"
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
            Index           =   6
            Left            =   240
            TabIndex        =   26
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Pack_TP:"
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
            Index           =   5
            Left            =   240
            TabIndex        =   25
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   0
         TabIndex        =   12
         Top             =   4920
         Width           =   8055
         Begin prjChameleon.chameleonButton cmdFirst 
            Height          =   375
            Left            =   240
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   240
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   ""
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
            FOCUSR          =   0   'False
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   255
            MCOL            =   14867671
            MPTR            =   99
            MICON           =   "Products.frx":0000
            PICN            =   "Products.frx":031A
            PICH            =   "Products.frx":06B4
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   -1  'True
            FX              =   1
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdPrev 
            Height          =   375
            Left            =   720
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   240
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   ""
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
            FOCUSR          =   0   'False
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   255
            MCOL            =   14867671
            MPTR            =   99
            MICON           =   "Products.frx":0A4E
            PICN            =   "Products.frx":0D68
            PICH            =   "Products.frx":1102
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   -1  'True
            FX              =   1
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdNext 
            Height          =   375
            Left            =   6960
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   240
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   ""
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
            FOCUSR          =   0   'False
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   255
            MCOL            =   14867671
            MPTR            =   99
            MICON           =   "Products.frx":149C
            PICN            =   "Products.frx":17B6
            PICH            =   "Products.frx":1B50
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   -1  'True
            FX              =   1
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdLast 
            Height          =   375
            Left            =   7440
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   240
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   ""
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
            FOCUSR          =   0   'False
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   255
            MCOL            =   14867671
            MPTR            =   99
            MICON           =   "Products.frx":1EEA
            PICN            =   "Products.frx":2204
            PICH            =   "Products.frx":259E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   -1  'True
            FX              =   1
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdClose 
            Height          =   375
            Left            =   5640
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "&Close"
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
            FOCUSR          =   0   'False
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   255
            MCOL            =   9469052
            MPTR            =   99
            MICON           =   "Products.frx":2938
            UMCOL           =   0   'False
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   1
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdAdd 
            Height          =   375
            Left            =   1680
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "&New"
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
            FOCUSR          =   0   'False
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   255
            MCOL            =   9469052
            MPTR            =   99
            MICON           =   "Products.frx":2C52
            UMCOL           =   0   'False
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   1
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdEdit 
            Height          =   375
            Left            =   3000
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "&Edit"
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
            FOCUSR          =   0   'False
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   255
            MCOL            =   9469052
            MPTR            =   99
            MICON           =   "Products.frx":2F6C
            UMCOL           =   0   'False
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   1
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdUpdate 
            Height          =   375
            Left            =   1680
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "&Save"
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
            FOCUSR          =   0   'False
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   255
            MCOL            =   9469052
            MPTR            =   99
            MICON           =   "Products.frx":3286
            UMCOL           =   0   'False
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   1
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdCancel 
            Height          =   375
            Left            =   3000
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "&Undo"
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
            FOCUSR          =   0   'False
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   255
            MCOL            =   9469052
            MPTR            =   99
            MICON           =   "Products.frx":35A0
            UMCOL           =   0   'False
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   1
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdDelete 
            Height          =   375
            Left            =   4320
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "&Delete"
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
            FOCUSR          =   0   'False
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   255
            MCOL            =   9469052
            MPTR            =   99
            MICON           =   "Products.frx":38BA
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
      Begin VB.Frame Frame3 
         Height          =   1335
         Left            =   3240
         TabIndex        =   8
         Top             =   600
         Width           =   4815
         Begin VB.TextBox Text1 
            DataField       =   "Product"
            Height          =   325
            Index           =   2
            Left            =   1080
            TabIndex        =   33
            Text            =   "Text1"
            Top             =   840
            Width           =   3615
         End
         Begin VB.TextBox Text1 
            DataField       =   "Category"
            Height          =   325
            Index           =   1
            Left            =   3360
            TabIndex        =   32
            Text            =   "Text1"
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            DataField       =   "Code"
            Height          =   325
            Index           =   0
            Left            =   1080
            TabIndex        =   31
            Text            =   "Text1"
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Code:"
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
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Product:"
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
            Left            =   240
            TabIndex        =   10
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Category:"
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
            Left            =   2400
            TabIndex        =   9
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1215
         Left            =   3240
         TabIndex        =   2
         Top             =   1920
         Width           =   4815
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            DataField       =   "ro_lvl"
            Height          =   325
            Index           =   7
            Left            =   3480
            TabIndex        =   38
            Text            =   "Text1"
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            DataField       =   "expirydate"
            Height          =   325
            Index           =   6
            Left            =   1080
            TabIndex        =   37
            Text            =   "Text1"
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            DataField       =   "pack_stock"
            Height          =   325
            Index           =   5
            Left            =   0
            TabIndex        =   36
            Text            =   "Text1"
            Top             =   480
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            DataField       =   "unit_stock"
            Height          =   325
            Index           =   4
            Left            =   3480
            TabIndex        =   35
            Text            =   "Text1"
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            DataField       =   "packing"
            Height          =   325
            Index           =   3
            Left            =   1080
            TabIndex        =   34
            Text            =   "Text1"
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Pack_Stock:"
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
            Index           =   8
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Unit_Stock:"
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
            Index           =   12
            Left            =   2400
            TabIndex        =   6
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "RO_lvl:"
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
            Index           =   4
            Left            =   2400
            TabIndex        =   5
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Packing:"
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
            Index           =   3
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Expiry:"
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
            Index           =   14
            Left            =   240
            TabIndex        =   3
            Top             =   720
            Width           =   735
         End
      End
      Begin iGrid250_75B4A91C.iGrid iGrid1 
         Height          =   4095
         Left            =   120
         TabIndex        =   0
         Top             =   720
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   7223
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
         RowMode         =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Products Form"
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
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   7935
      End
   End
End
Attribute VB_Name = "frmProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents rs As Recordset
Attribute rs.VB_VarHelpID = -1

Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean


Private Sub Form_Load()
On Error Resume Next
  'OpenDB
  cn.CursorLocation = adUseClient
  
  Set rs = New Recordset
  rs.Open "select * from Products", cn, adOpenStatic, adLockOptimistic
    
  Dim otext As TextBox
  'Bind the text boxes to the data provider
  For Each otext In Me.Text1
    Set otext.DataSource = rs
  Next
 
  LoadGrid1
  rs.MoveFirst
  mbAddNewFlag = False
  mbEditFlag = False
  mbDataChanged = False
  
End Sub

Private Sub Form_Resize()
CenterCtrl Me, Frame1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
  Set rs = Nothing
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  With rs
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    
    'GenSeq rs, Me.Text1(0)
    '.AddNew
    NewRecord
    mbAddNewFlag = True
    SetButtons False
    Text1(1).SetFocus
  End With

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  
  If MsgBox("Are you sure to delete this record?", vbYesNo) = vbYes Then
  With rs
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  LoadGrid1
  Exit Sub
  Else
  Exit Sub
  End If
DeleteErr:
  MsgBox Err.Description
  rs.CancelUpdate
  Exit Sub
  
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr
  mbEditFlag = True
  SetButtons False
  Me.Text1(1).SetFocus
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  rs.CancelUpdate
  If mvBookMark > 0 Then
    rs.Bookmark = mvBookMark
  Else
    rs.MoveFirst
  End If
  mbDataChanged = False

End Sub

Private Sub cmdUpdate_Click()
'On Error Resume Next
If Me.Text1(1).Text = "" Or Me.Text1(2).Text = "" Then
    MsgBox "Product & Category Fileds must be Enterd!", vbInformation
    Me.Text1(1).SetFocus
    Exit Sub
Else
    With rs
        For i = 0 To 13
        If Not Text1(i).Text = "" Then
           .Fields(i) = Text1(i).Text
        End If
        Next
        '.Fields(0) = Text1(0).Text
        '.Fields(1) = Text1(1).Text
        '.Fields(2) = Text1(2).Text
        '.Fields(3) = Text1(3).Text
        'If Not Text1(4).Text = "" Then
        '   .Fields(4) = Text1(4).Text
        'End If
        '.Fields(5) = Text1(5).Text
        '.Fields(6) = Text1(6).Text
        '.Fields(7) = Text1(7).Text
        '.Fields(8) = Text1(8).Text
        '.Fields(9) = Text1(9).Text
        '.Fields(10) = Text1(10).Text
        '.Fields(11) = Text1(11).Text
        '.Fields(12) = Text1(12).Text
        '.Fields(13) = Text1(13).Text
        .Update
    End With
    MsgBox "Record Saved Successfully!", vbInformation
  
 If mbAddNewFlag Then
    rs.MoveLast              'move to the new record
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  LoadGrid1
  Exit Sub
End If

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  rs.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
  Resume Next
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  rs.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
  Resume Next
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError

  If Not rs.EOF Then rs.MoveNext
  If rs.EOF And rs.RecordCount > 0 Then
    Beep
     'moved off the end so go back
    rs.MoveLast
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrev_Click()
  On Error GoTo GoPrevError

  If Not rs.BOF Then rs.MovePrevious
  If rs.BOF And rs.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    rs.MoveFirst
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bval As Boolean)
  cmdAdd.Visible = bval
  cmdEdit.Visible = bval
  cmdUpdate.Visible = Not bval
  cmdCancel.Visible = Not bval
  
  cmdDelete.Visible = bval
  cmdClose.Visible = bval
  cmdNext.Enabled = bval
  cmdFirst.Enabled = bval
  cmdLast.Enabled = bval
  cmdPrev.Enabled = bval
End Sub
Private Sub LoadGrid1()
Dim rs1 As New ADODB.Recordset
iGrid1.Clear True
rs1.Open "Select PRODUCT from products order by product asc", cn, adOpenForwardOnly, adLockReadOnly
iGrid1.FillFromRS rs1
iGrid1.ColWidth(1) = 180
Set rs1 = Nothing
End Sub

Private Sub iGrid1_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)
On Error Resume Next
If mbAddNewFlag = True Or mbEditFlag = True Then Exit Sub
rs.Requery
rs.MoveFirst
rs.Find "Product='" & iGrid1.CellValue(iGrid1.CurRow, 1) & "'"
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If mbAddNewFlag = True Or mbEditFlag = True Then
    If Index = 1 And KeyCode = vbKeyReturn Then
        frmSelCat.Show
    End If
End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
On Error Resume Next
If mbEditFlag = True Or mbAddNewFlag = True Then
Select Case Index
Case 2
    If mbAddNewFlag = True Then
    CheckDupName
    End If
Case 3
    CalculatePrice
Case 8
    CalculatePrice
Case 9
    CalculatePrice
Case Else
    Exit Sub
End Select

End If
End Sub

Private Sub CheckDupName()
Dim rs0 As New ADODB.Recordset
rs0.Open "Select Code,Product from Products", cn, adOpenForwardOnly, adLockReadOnly
rs0.MoveFirst
rs0.Find "Product='" & Me.Text1(2).Text & "'"
If rs0.EOF Or rs0.BOF Then
    Set rs0 = Nothing
    Exit Sub
Else
    MsgBox "You have a Product with this Name Already" & vbCrLf & "Product Description :" & rs0.Fields("Code") & "=" & rs0.Fields("Product"), vbInformation
    Me.Text1(2).SetFocus
    Set rs0 = Nothing
    Exit Sub
End If

End Sub

Private Sub NewRecord()
If rs.EOF Or rs.BOF Then
    NextID = 1
Else
    rs.MoveLast
    NextID = rs.Fields(0) + 1
End If
    rs.AddNew
    Text1(0).Text = NextID
End Sub

Private Sub CalculatePrice()
    Text1(10).Text = Format$((Val(Text1(8).Text) + Val(Text1(9).Text)) + ((Val(Text1(8).Text) + Val(Text1(9).Text)) * 0.15), "0.00")
    Text1(11).Text = Format$(Val(Text1(8).Text) / Val(Text1(3).Text), "0.00")
    Text1(12).Text = Format$(Val(Text1(9).Text) / Val(Text1(3).Text), "0.00")
    Text1(13).Text = Format$(Val(Text1(10).Text) / Val(Text1(3).Text), "0.00")
End Sub
