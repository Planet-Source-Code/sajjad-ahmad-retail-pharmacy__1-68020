VERSION 5.00
Begin VB.Form frmLock 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   9105
   ClientLeft      =   120
   ClientTop       =   900
   ClientWidth     =   9540
   LinkTopic       =   "Form2"
   Picture         =   "Lock.frx":0000
   ScaleHeight     =   9105
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
If CURRENT_COMPANY.LOCK_KEY = "" Then
    Unload Me
ElseIf KeyAscii = Asc(CURRENT_COMPANY.LOCK_KEY) Then
    Unload Me
End If
End Sub

