Attribute VB_Name = "Module1"
'CONNECTION VARIABLE
Public cn As New ADODB.Connection
Public NextID As Integer

'COMPANY INFO TYPE CLASS
Public CURRENT_COMPANY As COMPANY_INFO

Public Type COMPANY_INFO
    COMPANY_NAME        As String
    COMPANY_ADDRESS     As String
    COMPANY_PHONE       As String
    LOCK_KEY            As String
    DB_PASSWORD         As String
    ADMIN_PASS          As String
End Type
Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Public Const SEE_MASK_INVOKEIDLIST = &HC
Public Const SEE_MASK_NOCLOSEPROCESS = &H40
Public Const SEE_MASK_FLAG_NO_UI = &H400

Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long

Public Sub ShowProps(FileName As String, OwnerhWnd As Long)

    Dim SEI As SHELLEXECUTEINFO
    Dim lngReturn As Long
     
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or _
         SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .hwnd = OwnerhWnd
        .lpVerb = "properties"
        .lpFile = FileName
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = 0
        .lpIDList = 0
    End With
     
    lngReturn = ShellExecuteEX(SEI)
    
End Sub
Sub Main()
'On Error Resume Next
Dim IRD As String
Dim LRD As String

If App.PrevInstance = True Then
    MsgBox "Pharmacy is Already Runnig on this system!", vbInformation
    End
Else
    Dim frm As Form
    For Each frm In Forms
        Load frm
    Next
'With frmRegister.ActiveLock1
'    If .RegisteredUser = True Then
 '       frmSplash.lblCompany = "Registered Version"
 '       frmAbout.Command2.Visible = False
 '       frmAbout.lblStatus.Caption = "REGISTERED VERSION"
        frmSplash.Show
  '  ElseIf .RegisteredUser = False Then
   '     If IsNull(.LastRunDate) = True Then
    '        MsgBox "SecurityBUG has detected that you've changed the Settings!", vbCritical
     '       End
     '   ElseIf .LastRunDate > Now Then
     '    MsgBox "SecurityBUG has detected that you've changed the clock backwards!", vbCritical
     '    End
       'Check the evaluation period
      '  ElseIf .UsedDays < 20 Then
       '     frmAbout.lblStatus.Caption = "You have " & 20 - .UsedDays & " Days Left to use this software!"
            'MsgBox "You have " & 20 - .UsedDays & " Days Left to use this software!"
        '    frmSplash.Show
         '   frmSplash.lblCompany = "Trial Version"
       ' ElseIf .UsedDays > 20 Then
            ' If the evaluation period has expired...
        '    MsgBox "Your evaluation period has expired!"
         '   End
        'End If
    'End If
'End With
End If

End Sub

Public Sub OpenDB()
On Error GoTo cn_err
    cn.Open "File Name=" & App.Path & "\pharmacy.udl;"
cn_err:
If Err.Number < 0 Then
    Call ShowProps(App.Path & "\pharmacy.udl", frmLogin.hwnd)
End If
End Sub
Function FuncPrev(rs As ADODB.Recordset)

If Not rs.BOF Then rs.MovePrevious
    If rs.BOF And rs.RecordCount > 0 Then
    Beep
    rs.MoveFirst
End If

End Function
Function FuncNext(rs As ADODB.Recordset)

If Not rs.EOF Then rs.MoveNext
    If rs.EOF And rs.RecordCount > 0 Then
    Beep
    rs.MoveLast
End If

End Function


Function CenterCtrl(frm As Form, Ctrl As Control)
    Ctrl.Left = (frm.ScaleWidth - Ctrl.Width) / 2
    Ctrl.Top = (frm.ScaleHeight - Ctrl.Height) / 2
End Function


' DECODE PASSWORD
Function Mcode_Pass(p_str As String) As String
    For i = 1 To Len(p_str) Step 1
        strs = strs + Chr(Asc(Mid(p_str, i, 1)) * 2)
    Next i
        Mcode_Pass = strs
End Function

' CODE PASSWORD
Function DCode_Pass(p_str As String) As String
    For i = 1 To Len(p_str) Step 1
        strs = strs + Chr(Asc(Mid(p_str, i, 1)) / 2)
    Next i
        DCode_Pass = strs
End Function

Function IsLoaded(ByVal frmName As String) As Long
    Dim frm As Form
    For Each frm In Forms
        If StrComp(frm.Name, frmName, vbTextCompare) = 0 Then
            IsLoaded = IsLoaded + 1
        End If
    Next
End Function

Function GenSeq(rec As Recordset, Text1 As TextBox)
If rec.EOF Or rec.BOF Then
    NextID = 1
Else
    rec.MoveLast
    NextID = rec.Fields(0) + 1
End If
    rec.AddNew
    Text1 = NextID
    
End Function

Function HoverText(ctl As Control, shapoo As Shape)

'Size the border to fit around textbox
shapoo.Left = ctl.Left - 30
shapoo.Top = ctl.Top - 30
shapoo.Height = ctl.Height + 65
shapoo.Width = ctl.Width + 65
ctl.BackColor = &HC0FFFF
'ctl.ForeColor = vbWhite

'Select whole text in textbox
ctl.SelStart = 0
ctl.SelLength = Len(ctl.Text)

'Show the border around text
shapoo.Visible = True

End Function
Function ClearHover(ctl As Control, shapoo As Shape)
ctl.BackColor = vbWhite
'ctl.ForeColor = vbBlack

shapoo.Visible = False
End Function



