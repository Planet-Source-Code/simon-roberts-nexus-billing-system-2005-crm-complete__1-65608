VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmIMport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import CSV File"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Import Process"
      Height          =   2325
      Left            =   150
      TabIndex        =   2
      Top             =   4650
      Width           =   8985
      Begin VB.CommandButton Command5 
         Caption         =   "Do Import"
         Height          =   435
         Left            =   150
         TabIndex        =   9
         Top             =   1710
         Width           =   2145
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Select Sysop"
         Height          =   435
         Left            =   150
         TabIndex        =   5
         Top             =   1230
         Width           =   2145
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Select ViSP"
         Height          =   435
         Left            =   150
         TabIndex        =   4
         Top             =   750
         Width           =   2145
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Select Plan to Import As"
         Height          =   435
         Left            =   150
         TabIndex        =   3
         Top             =   270
         Width           =   2145
      End
      Begin VB.Label lblStatus 
         Height          =   195
         Index           =   2
         Left            =   2475
         TabIndex        =   8
         Top             =   1290
         Width           =   555
      End
      Begin VB.Label lblStatus 
         Height          =   195
         Index           =   1
         Left            =   2475
         TabIndex        =   7
         Top             =   840
         Width           =   405
      End
      Begin VB.Label lblStatus 
         Height          =   195
         Index           =   0
         Left            =   2475
         TabIndex        =   6
         Top             =   390
         Width           =   585
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   8250
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Load CSV File"
      Height          =   405
      Left            =   150
      TabIndex        =   1
      Top             =   120
      Width           =   2115
   End
   Begin MSComctlLib.ListView lvImport 
      Height          =   3975
      Left            =   150
      TabIndex        =   0
      Top             =   570
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   7011
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Account Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Username"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Password"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "yes/no to extra services"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Comment"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmIMport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fAddPlan As New frmAddPlan

Private Sub Command1_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Command1_Click"
    Const ContainerName = "frmIMport"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
'**  This code is not to be distributed, reverse engineered or simulated in any way without   **
'**  Premission from the author. The authors of this code is as follows: Simon Antony Roberts **                                                     **
'**Jarrett Cliff Costi, these two are the only people you can communicate with about this code**
'***********************************************************************************************
'**  Project Alpha is a product of Exitstencil Press Australia                                **
'***********************************************************************************************
'**                                                                                           **
'**  Routine:                                                                                 **
'**  Arguments:                                                                               **
'**  Description:    Subroutine, Function or Property of The Nexus                        **
'**  Author:         Simon Roberts                                                            **
'**  Date Last mod:  19-01-2004                                                               **
'**                                                                                           **
'********************************************** Copyright © 2004 Exitstencil Press Australia ***
'
'
'
    If bDebug = -1 Then
        On Error GoTo 0
    ElseIf bDebug = 1 Then
        On Error Resume Next
    Else
        On Error GoTo ErrorOccur
    End If


    cd.Filename = ""
    cd.Filter = "CSV File (*.CSV)|*.CSV"
    cd.FilterIndex = 1
    cd.ShowOpen
    
    If cd.Filename = "" Then Exit Sub
    
    Dim lFileNum As Long
    
    lFileNum = FreeFile
    
    Open cd.Filename For Input As #lFileNum
    
    Dim Field1 As String
    Dim Field2 As String
    Dim Field3 As String
    Dim Field4 As String
    Dim Field5 As String
    
    Dim itmX As ListItem
    
    lvImport.ListItems.Clear
    
    On Error Resume Next
    
    Input #lFileNum, Field1, Field2, Field3, Field4, Field5
    
    While Not EOF(lFileNum)
    
        Input #lFileNum, Field1, Field2, Field3, Field4, Field5
    
        Set itmX = lvImport.ListItems.Add(, , Field1)
        itmX.SubItems(1) = Field2
        itmX.SubItems(2) = Field3
        itmX.SubItems(3) = Field4
        itmX.SubItems(4) = Field4
        
    Wend
    
Exit Sub



ErrorOccur:
Select Case oErr.chkError(directConn, Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Sub

Private Sub Command2_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Command2_Click"
    Const ContainerName = "frmIMport"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
'**  This code is not to be distributed, reverse engineered or simulated in any way without   **
'**  Premission from the author. The authors of this code is as follows: Simon Antony Roberts **                                                     **
'**Jarrett Cliff Costi, these two are the only people you can communicate with about this code**
'***********************************************************************************************
'**  Project Alpha is a product of Exitstencil Press Australia                                **
'***********************************************************************************************
'**                                                                                           **
'**  Routine:                                                                                 **
'**  Arguments:                                                                               **
'**  Description:    Subroutine, Function or Property of The Nexus                        **
'**  Author:         Simon Roberts                                                            **
'**  Date Last mod:  19-01-2004                                                               **
'**                                                                                           **
'********************************************** Copyright © 2004 Exitstencil Press Australia ***
'
'
'
    If bDebug = -1 Then
        On Error GoTo 0
    ElseIf bDebug = 1 Then
        On Error Resume Next
    Else
        On Error GoTo ErrorOccur
    End If


    If bDebug = True Then On Error GoTo 0 Else On Error Resume Next
    
    
    fAddPlan.Show 1
    
    lblStatus(0).Caption = "" & fAddPlan.ptRecID
    
    'FlagID = 8.43874378437844E+22
    
    Screen.MousePointer = vbDefault
        
    
Exit Sub



ErrorOccur:
Select Case oErr.chkError(directConn, Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Sub

Private Sub Command3_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Command3_Click"
    Const ContainerName = "frmIMport"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
'**  This code is not to be distributed, reverse engineered or simulated in any way without   **
'**  Premission from the author. The authors of this code is as follows: Simon Antony Roberts **                                                     **
'**Jarrett Cliff Costi, these two are the only people you can communicate with about this code**
'***********************************************************************************************
'**  Project Alpha is a product of Exitstencil Press Australia                                **
'***********************************************************************************************
'**                                                                                           **
'**  Routine:                                                                                 **
'**  Arguments:                                                                               **
'**  Description:    Subroutine, Function or Property of The Nexus                        **
'**  Author:         Simon Roberts                                                            **
'**  Date Last mod:  19-01-2004                                                               **
'**                                                                                           **
'********************************************** Copyright © 2004 Exitstencil Press Australia ***
'
'
'
    If bDebug = -1 Then
        On Error GoTo 0
    ElseIf bDebug = 1 Then
        On Error Resume Next
    Else
        On Error GoTo ErrorOccur
    End If


    lblStatus(1).Caption = "" & Login.lVirtualID
    
    
Exit Sub



ErrorOccur:
Select Case oErr.chkError(directConn, Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Sub

Private Sub Command4_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Command4_Click"
    Const ContainerName = "frmIMport"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
'**  This code is not to be distributed, reverse engineered or simulated in any way without   **
'**  Premission from the author. The authors of this code is as follows: Simon Antony Roberts **                                                     **
'**Jarrett Cliff Costi, these two are the only people you can communicate with about this code**
'***********************************************************************************************
'**  Project Alpha is a product of Exitstencil Press Australia                                **
'***********************************************************************************************
'**                                                                                           **
'**  Routine:                                                                                 **
'**  Arguments:                                                                               **
'**  Description:    Subroutine, Function or Property of The Nexus                        **
'**  Author:         Simon Roberts                                                            **
'**  Date Last mod:  19-01-2004                                                               **
'**                                                                                           **
'********************************************** Copyright © 2004 Exitstencil Press Australia ***
'
'
'
    If bDebug = -1 Then
        On Error GoTo 0
    ElseIf bDebug = 1 Then
        On Error Resume Next
    Else
        On Error GoTo ErrorOccur
    End If


    lblStatus(2).Caption = "" & Login.lSysopID
    
Exit Sub



ErrorOccur:
Select Case oErr.chkError(directConn, Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Sub

'Private Sub Command5_Click()
'
'
'    '*[ Error Checking Variables ]**********************************************************************************
'
'
'    Const RoutineName = "Command5_Click"
'    Const ContainerName = "frmIMport"
'    '***************************************************************************************************************
'
'
''
''***********************************************************************************************
''**  Project Alpha ® 2003, 2004 +                                                             **
''***********************************************************************************************
''**  This code is not to be distributed, reverse engineered or simulated in any way without   **
''**  Premission from the author. The authors of this code is as follows: Simon Antony Roberts **                                                     **
''**Jarrett Cliff Costi, these two are the only people you can communicate with about this code**
''***********************************************************************************************
''**  Project Alpha is a product of Exitstencil Press Australia                                **
''***********************************************************************************************
''**                                                                                           **
''**  Routine:                                                                                 **
''**  Arguments:                                                                               **
''**  Description:    Subroutine, Function or Property of The Nexus                        **
''**  Author:         Simon Roberts                                                            **
''**  Date Last mod:  19-01-2004                                                               **
''**                                                                                           **
''********************************************** Copyright © 2004 Exitstencil Press Australia ***
''
''
''
'    If bDebug = -1 Then
'        On Error GoTo 0
'    ElseIf bDebug = 1 Then
'        On Error Resume Next
'    Else
'        On Error GoTo ErrorOccur
'    End If
'
'
'
'    Dim lRecID As Long
'    Dim ix As Long
'
'    Dim itmy As ListItem
'
'
'     If bDebug = True Then On Error GoTo 0 Else On Error Resume Next
'
'    Dim fAddPlan As New frmAddPlan
'    Dim frmPayments As New frmAccPayment
'    Dim fDomain As New frmDNS
'    Dim fRadius As New frmRadiusAccount
'
'    Dim fDSL As New frmBroadband
'    Dim fAlias As New frmAlias
'
'
'    Dim rsRadius As ADODB.Recordset
'    Dim rsCheck As ADODB.Recordset
'    Dim rsinvoiceout As ADODB.Recordset
'    Dim rsAccInfo As ADODB.Recordset
'    Dim rsDomains As ADODB.Recordset
'
'    Dim fPOP3 As frmPOP3Account
'    Dim fFTP As frmFTPAccount
'    Dim bx As Byte
'    Dim sPassword As String
'
'    Dim l5meoDMT As Variant
'
'    Dim FlagID As Integer
'
'    Dim bSkip As Boolean
'    Dim lRecID_acci_services As Long
'
'    Set fFTP = New frmFTPAccount
'    Set fPOP3 = New frmPOP3Account
'    Set rsinvoiceout = New ADODB.Recordset
'    Set rsAccInfo = New ADODB.Recordset
'
'    Dim InvRecID  As Variant
'    Dim bResult As Boolean
'    Dim lSubRecID As Long
'    Dim rsSave As ADODB.Recordset
'    Dim RSsUBS As ADODB.Recordset
'
'    Dim oSvCount As Long
'
'    Dim acci_services As type_acci_services
'
'    If lvImport.ListItems.Count > 0 Then
'        For ix = 1 To lvImport.ListItems.Count
'            lRecID = 0
'            Set itmy = lvImport.ListItems(ix)
'
'                If lRecID = 0 Then
''                    dtpBilling.Value = sysNow
'                    Call MySQL.OpenTable(directConn, rsSave, , "select * from accountinfo Limit 1")
'                    rsSave.AddNew
'                    rsSave!SysopID = Login.lSysopID
'                    rsSave!VirtualID = Login.lVirtualID
'                    'rsSave!AccountTypeID = CLng(Mid(itmX.Key, 2))
'                    bsmtpCreation = True
'                    bSetRecID = True
'                End If
'
'                rsSave!AccountName = itmy.Text
'                rsSave!DOB = sysNOW
'                rsSave!Cancelled = 0
'                rsSave!ProcessFlag = 1
'            '    rsSave!Realm = txtRealm.Text
'                rsSave!BillingDate = "2003-8-01 12:00:00"
'
'                If IsNull(rsSave!BillingDate) Then rsSave!BillingDate = DateAdd("m", 1, sysNOW)
'
'
'                    rsSave!FlagA_RecID = 0
'                    rsSave!FlagASet = sysNOW
'                    rsSave!FlagB_RecID = 0
'                    rsSave!FlagBSet = sysNOW
'                    rsSave!AboutUS = 0
'
'                    rsSave!PayIntervalType = "d"
'                    rsSave!PayInterval = 14
'
'                    rsSave!Classification = 2
'
'
'                    MySQL.Execute directConn, "UPDATE virtualisp Set NoSub=NoSub+1 where RecID =" & Login.lVirtualID
'
'                lRecID = MySQL.SetRecID(rsSave, "accountinfo", directConn)
'
'
'
'
'
'                MySQL.Execute directConn, "UPDATE accountinfo SET gPassword=MD5('" & itmy.SubItems(2) & "'), gUsername='" & itmy.SubItems(1) & "' Where RecID = " & lRecID
'
'               Call MySQL.OpenTable(directConn, rsSave, , "select * from acci_emailaddresses Limit 0")
'
'                rsSave.AddNew
'                rsSave!acci_RecID = lRecID
'                rsSave!ContactName = itmy.Text
'                rsSave!EmailAddress = itmy.SubItems(1) & "@" & Login.sVISPDomain
'                rsSave!Checked = 0
'
'                rsSave.Update
'
'                Call MySQL.OpenTable(directConn, rsSave, , "select * from acci_emailaddresses Limit 0")
'
'                rsSave.AddNew
'                rsSave!acci_RecID = lRecID
'                rsSave!ContactName = itmy.Text
'                rsSave!EmailAddress = "invoicing@gondwananet.com"
'                rsSave!Checked = True
'
'                rsSave.Update
'
'
'
'                If oService.Count <> 0 Then
'
'                    tmpRecID = lRecID
'
'                    For oSvCount = 1 To oService.Count
'
'                        bSkip = False
'
'                        acci_services.acciRecID = tmpRecID
'
'                        acci_services.JoiningFee = oService(oSvCount).JoiningFee
'                        acci_services.PerHour = oService(oSvCount).PerHour
'                        acci_services.PeriodFee = oService(oSvCount).PeriodFee
'                        acci_services.PerMB = oService(oSvCount).PerMBBlock
'
'                        acci_services.Activation = fAddPlan.Activation
'                        acci_services.NextCycle = DateAdd(oService(oSvCount).cycType, oService(oSvCount).cycInterval, sysNOW)
'                        acci_services.ServiceID = oService(oSvCount).ServiceID
'                        acci_services.ptRecID = oService(oSvCount).ptRecID
'                        acci_services.MBQuota = oService(oSvCount).MBQuota
'
'                        If DateDiff("s", acci_services.Activation, sysNOW) > 0 Then acci_services.Checked = True Else acci_services.Checked = False
'
'                        If oService(oSvCount).ListedOnRadius = True Then
'                            fRadius.p_SessionTimeOut = oService(oSvCount).SessionTimeout
'                            fRadius.p_IdleTimeout = oService(oSvCount).IdleTimeout
'                            fRadius.p_Sessions = oService(oSvCount).SessionsAllowed
'
'                            Screen.MousePointer = vbDefault
'                            Screen.MousePointer = vbHourglass
'
'
'                            bResult = MySQL.OpenTable(directConn, rsRadius, , "select * from radiusaccounts Limit 1")
'                            rsRadius.AddNew
'                            rsRadius!Username = itmy.SubItems(1)
'                            'rsRadius!Password = fradius.p_Password
'                            rsRadius!acci_RecID = tmpRecID
'                            acci_services.Username = itmy.SubItems(1)
'                            acci_services.Password = itmy.SubItems(2)
'                            rsRadius!SessionsAllowed = 1
'                            rsRadius!AutoActivateFlag = 0
'                            rsRadius!Activate = "12:00:00"
'                            rsRadius!Deactivate = "00:00:00"
'                            rsRadius!SessionTimeout = fRadius.p_SessionTimeOut
'                            rsRadius!IdleTimeout = fRadius.p_IdleTimeout
'                            rsRadius!Checked = True
'                            rsRadius!VirtualID = Login.lVirtualID
'                            rsRadius!ptRecID = oService(oSvCount).ptRecID
'                            acci_services.RadiusID = MySQL.SetRecID(rsRadius, "radiusaccounts", directConn)
'                            acci_services.ContactName = itmy.Text
'
'                            MySQL.Execute directConn, "UPDATE radiusaccounts SET Password=AES_ENCRYPT('" & itmy.SubItems(2) & "','" & odb.colSalts.ReturnSalt("md5Password") & "') Where RecID = " & acci_services.RadiusID
'                            'mysql.execute directConn,  "UPDATE radiusaccounts SET Password=MD5('" & fradius.p_Password & "') Where RecID = " & rsSave3!RadiusID
'
'                        End If
'
'                        FlagID = 0
'                        bSkip = False
'                        Select Case oService(oSvCount).svrCode
'                        Case "ALIAS"
'                            bSkip = True
'                        Case "ADSL", "SHDSL"
'                            bSkip = True
'                        Case "FTP"
'
'                            Screen.MousePointer = vbDefault
'                            Screen.MousePointer = vbHourglass
'                            acci_services.Username = itmy.SubItems(1)
'                            acci_services.Password = itmy.SubItems(2)
'                            acci_services.BaseURL = "/home/" & Login.sVISPDomain & "/" & itmy.SubItems(1)
'                            acci_services.DynamicField1 = fFTP.sSessions
'                            acci_services.DynamicField2 = fFTP.byteBandwidth
'                            acci_services.DynamicField3 = fFTP.byteBWUpload
'                            acci_services.ContactName = fFTP.sContactName
'                            acci_services.DynamicField4 = "21"
'                            acci_services.DynamicField5 = "21"
'
'                           ' rsSave3.Update
'                        Case "POP3"
'
'DoPOP3Again2:
'                            fPOP3.sContactName = txtAccountName.Text
'
'                            fPOP3.acci_acciRecID = acci_services.acciRecID
'                            fPOP3.acci_ptRecID = acci_services.ptRecID
'                            fPOP3.acci_Activation = acci_services.Activation
'                            fPOP3.acci_BaseURL = acci_services.BaseURL
'                            fPOP3.acci_Checked = acci_services.Checked
'                            fPOP3.acci_ContactName = itmy.Text
'                            fPOP3.acci_DomainID = acci_services.DomainID
'                            fPOP3.acci_Description = oService(oSvCount).Description
'                            fPOP3.acci_DynamicField1 = acci_services.DynamicField1
'                            fPOP3.acci_DynamicField3 = acci_services.DynamicField3
'                            fPOP3.acci_DynamicField2 = acci_services.DynamicField2
'                            fPOP3.acci_DynamicField4 = acci_services.DynamicField4
'                            fPOP3.acci_DynamicField5 = acci_services.DynamicField5
'                            fPOP3.acci_MBQuota = acci_services.MBQuota
'                            fPOP3.acci_NextCycle = acci_services.NextCycle
'                            fPOP3.acci_Password = itmy.SubItems(2)
'                            fPOP3.acci_ptRecID = acci_services.ptRecID
'                            fPOP3.acci_RadiusID = acci_services.RadiusID
'                            fPOP3.acci_RecID = lRecID
'                            fPOP3.acci_ServiceID = acci_services.ServiceID
'                            fPOP3.acci_Username = itmy.SubItems(1)
'                            fPOP3.lSubRecID = lSubRecID
'                            fPOP3.oSvrIndex = oSvCount
'
'
'                            fPOP3.acci_DynamicField1 = "Scanning"
'
'                            fPOP3.acci_DynamicField2 = oService(oSvCount).NumOf
'
'                           ' On Error Resume Next
'                            Do
'                                Err.Clear
'                                lRecID_acci_services = MySQL.GetTMPRecID("acci_services", directConn)
'
'
'                                MySQL.Execute directConn, "Insert into acci_services (RecID, RadiusID, ServiceID, ptRecID, DomainID, VirtualID, acci_RecID, SysopID, NextCycle, Checked, ContactName, Username, Password, BaseURL, DynamicField1, DynamicField2, DynamicField3, DynamicField4, DynamicField5, SubRecID, MBQuota, Activation, PeriodFee, JoiningFee, PerMB, PerHour) " + _
'                                                "VALUES (" & lRecID_acci_services & "," & fPOP3.acci_RadiusID & "," & fPOP3.acci_ServiceID & "," & fPOP3.acci_ptRecID & "," & fPOP3.acci_DomainID & "," & Login.lVirtualID & "," & fPOP3.acci_acciRecID & "," & Login.lSysopID & ",'" & Format(IIf(DateDiff("s", fPOP3.acci_Activation, sysNOW) > 0, fPOP3.acci_NextCycle, fPOP3.acci_Activation), "yyyy-mm-dd Hh:Nn:Ss") & "', " & _
'                                                IIf(fPOP3.acci_Checked = True, "-1", "0") & ",'" & MySQL.ESC(fPOP3.acci_ContactName) & "','" & MySQL.ESC(fPOP3.acci_Username) & "',AES_ENCRYPT('" & MySQL.ESC(fPOP3.acci_Password) & "','" & odb.colSalts.ReturnSalt("md5Password") & "'),'" & fPOP3.acci_BaseURL & "','" & fPOP3.acci_DynamicField1 & "','" & fPOP3.acci_DynamicField2 & _
'                                                "','" & fPOP3.acci_DynamicField3 & "','" & fPOP3.acci_DynamicField4 & "','" & fPOP3.acci_DynamicField5 & "'," & fPOP3.lSubRecID & "," & fPOP3.acci_MBQuota & ",'" & Format(fPOP3.acci_Activation, "yyyy-mm-dd Hh:Nn:Ss") & "'," & 0 & "," & 0 & "," & 0 & "," & 0 & ")"
'
'                                If Err.Number <> 0 Then
'                        '           Stop
'                                   cDebug Err.Description
'                                End If
'
'                            Loop Until Err.Number = 0
'
'
'                            Screen.MousePointer = vbDefault
'
'                            Screen.MousePointer = vbHourglass
'
'                            bSkip = True
'                        Case "DOMAIN"
'                            bSkip = True
'                        End Select
'
'                        'bResult = MySQL.OpenTable(directConn, rsSave3, , "select * from acci_services Limit 1")
'
'
'
'                        On Error Resume Next
'                        If bSkip = False Then
'
'                            If acci_services.ContactName = "" Then acci_services.ContactName = txtAccountName.Text
'
'                            Do
'                                Err.Clear
'                                lRecID_acci_services = MySQL.GetTMPRecID("acci_services", directConn)
'
'                                If oSvCount = 1 Then lSubRecID = lRecID_acci_services
'
'
'                                MySQL.Execute directConn, "Insert into acci_services (RecID, RadiusID, ServiceID, ptRecID, DomainID, VirtualID, AccI_RecID, SysopID, NextCycle, Checked, ContactName, Username, Password, BaseURL, DynamicField1, DynamicField2, DynamicField3, DynamicField4, DynamicField5, SubRecID, MBQuota, Activation, PeriodFee, JoiningFee, PerMB, PerHour) " + _
'                                                "VALUES (" & lRecID_acci_services & "," & acci_services.RadiusID & "," & acci_services.ServiceID & "," & acci_services.ptRecID & "," & acci_services.DomainID & "," & Login.lVirtualID & ",'" & lRecID & "'," & Login.lSysopID & ",'" & Format(IIf(DateDiff("s", acci_services.Activation, sysNOW) > 0, acci_services.NextCycle, acci_services.Activation), "yyyy-mm-dd Hh:Nn:Ss") & "', " & _
'                                                IIf(True = True, "-1", "0") & ",'" & MySQL.ESC(itmy.Text) & "','" & MySQL.ESC(itmy.SubItems(1)) & "',AES_ENCRYPT('" & MySQL.ESC(itmy.SubItems(2)) & "','" & odb.colSalts.ReturnSalt("md5Password") & "'),'" & acci_services.BaseURL & "','" & acci_services.DynamicField1 & "','" & acci_services.DynamicField2 & _
'                                                "','" & acci_services.DynamicField3 & "','" & acci_services.DynamicField4 & "','" & acci_services.DynamicField5 & "'," & lSubRecID & "," & acci_services.MBQuota & ",'" & Format(acci_services.Activation, "yyyy-mm-dd Hh:Nn:Ss") & "'," & acci_services.PeriodFee & "," & acci_services.JoiningFee & "," & acci_services.PerMB & "," & acci_services.PerHour & ")"
'                                If Err.Number <> 0 Then cDebug Err.Description
'
'                            Loop Until Err.Number = 0
'
'                            If bDebug = True Then On Error GoTo 0 Else On Error Resume Next
'
'                            Screen.MousePointer = vbDefault
'
'                        End If
'
'                        acci_services.JoiningFee = 0
'                        acci_services.PerHour = 0
'                        acci_services.PeriodFee = 0
'                        acci_services.PerMB = 0
'                        'acci_services.acciRecID = tmpRecID
'                        acci_services.BaseURL = ""
'                        acci_services.Checked = 0
'                        acci_services.ContactName = ""
'                        acci_services.DynamicField1 = ""
'                        acci_services.DynamicField2 = ""
'                        acci_services.DynamicField3 = ""
'                        acci_services.DynamicField4 = ""
'                        acci_services.DynamicField5 = ""
'                        'acci_services.NextCycle = Null
'                        acci_services.Password = ""
'                        acci_services.ptRecID = 0
'                        acci_services.ServiceID = 0
'                        acci_services.Username = ""
'                        acci_services.DomainID = 0
'                        acci_services.RadiusID = 0
'                        acci_services.MBQuota = 0
'
'                        Dim bRefered As Long
'
'                        If DateDiff("s", fAddPlan.Activation, sysNOW) < 0 Then
'                            mvBilling.Value = fAddPlan.Activation
'                        ElseIf bSkip = False Then
'                            Select Case False 'oService(oSvCount).BillNow
'                            Case -1
'
'                                bResult = MySQL.OpenTable(directConn, rsinvoiceout, , "select * from invoiceout Limit 1")
'
'                                rsinvoiceout.AddNew
'
'                                If lvPlans.ListItems.Count > 0 Then bRefered = vbNo
'
'                                If bRefered = 0 Then
'                                    Screen.MousePointer = vbDefault
'                                    Select Case MsgBox("Was this transaction refered by another customer? If so the Joining fee is obmitted.", vbQuestion + vbYesNo, "Was customer refered")
'                                    Case vbYes
'                                        rsinvoiceout!AmountDue = oService(oSvCount).PeriodFee
'                                        rsinvoiceout!Description = CStr(oService(oSvCount).Description) + " - Setup Free"
'                                        bRefered = vbYes
'                                    Case vbNo
'                                        rsinvoiceout!AmountDue = oService(oSvCount).PeriodFee + oService(oSvCount).JoiningFee
'                                        rsinvoiceout!Description = CStr(oService(oSvCount).Description) + " - Setup Fee Included [" + Format(oService(oSvCount).JoiningFee, "Currency") + "]"
'                                        bRefered = vbNo
'                                    End Select
'                                    Screen.MousePointer = vbHourglass
'                                Else
'                                    Select Case bRefered
'                                    Case vbYes
'                                        rsinvoiceout!AmountDue = oService(oSvCount).PeriodFee
'                                        rsinvoiceout!Description = CStr(oService(oSvCount).Description) + " - Setup Free"
'                                        bRefered = vbYes
'                                    Case vbNo
'                                        rsinvoiceout!AmountDue = oService(oSvCount).PeriodFee + oService(oSvCount).JoiningFee
'                                        rsinvoiceout!Description = CStr(oService(oSvCount).Description) + " - Setup Fee Included [" + Format(oService(oSvCount).JoiningFee, "Currency") + "]"
'                                        bRefered = vbNo
'                                    End Select
'                                End If
'
'                                rsinvoiceout!PlanServiceID = lRecID_acci_services
'                                rsinvoiceout!GSTCharged = rsinvoiceout!AmountDue * oTax(Login.TaxCode, Login.TaxCountry)
'                                rsinvoiceout!TotalDue = rsinvoiceout!AmountDue + rsinvoiceout!GSTCharged
'                                rsinvoiceout!AmountPaid = 0
'                                rsinvoiceout!SysopID = Login.lSysopID
'                                rsinvoiceout!VirtualID = Login.lVirtualID
'                                rsinvoiceout!FlagID = FlagID
'
'                                If tmpRecID <> 0 And lRecID = 0 Then
'                                    rsinvoiceout!acci_RecID = tmpRecID
'                                ElseIf tmpRecID = 0 And lRecID <> 0 Then
'                                    rsinvoiceout!acci_RecID = lRecID
'                                Else
'                                    rsinvoiceout!acci_RecID = lRecID
'                                End If
'
'                                For bx = optBillingCycle.LBound To optBillingCycle.UBound
'                                    If optBillingCycle(bx).Value = True Then
'                                        rsinvoiceout!PaymentDue = DateAdd("d", 14, sysNOW)
'                                        Exit For
'                                    End If
'                                Next
'
'                                InvRecID = MySQL.SetRecID(rsinvoiceout, "invoiceout", directConn)
'
'                            End Select
'                        End If
'
'                    Next
'                End If
'        Next ix
'    End If
'
'Screen.MousePointer = vbDefault
'Exit Sub
'
'
'
'ErrorOccur:
'Select Case oErr.chkError(directConn,Val(Err.Number), Err.Description, RoutineName, ContainerName)
'Case vbResume
'    Resume
'Case vbExit
'
'Case vbResumeNext
'    Resume Next
'End Select
'
'End Sub
