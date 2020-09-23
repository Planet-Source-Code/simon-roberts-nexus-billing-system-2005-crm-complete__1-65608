VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmPOP3Account 
   BackColor       =   &H00C00000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "POP3 Account"
   ClientHeight    =   8025
   ClientLeft      =   2610
   ClientTop       =   2625
   ClientWidth     =   8130
   ControlBox      =   0   'False
   Icon            =   "frmPOP3Account_new.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lveMail 
      Height          =   3495
      Left            =   870
      TabIndex        =   6
      Top             =   3900
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   6165
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Contact name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Username"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Password"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Domain"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6240
      TabIndex        =   8
      Top             =   7530
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   870
      TabIndex        =   7
      Top             =   7530
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00BA3F3F&
      Height          =   2985
      Left            =   870
      TabIndex        =   9
      Top             =   750
      Width           =   6945
      Begin VB.CheckBox chkSpam 
         BackColor       =   &H00BA3F3F&
         Caption         =   "Scan for spam and virus' at server"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0084E8E8&
         Height          =   255
         Left            =   210
         TabIndex        =   15
         Top             =   2550
         Value           =   1  'Checked
         Width           =   3285
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update Entry"
         Enabled         =   0   'False
         Height          =   345
         Left            =   5160
         TabIndex        =   5
         Top             =   2520
         Width           =   1515
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add e-Mail Account"
         Height          =   345
         Left            =   3570
         TabIndex        =   4
         Top             =   2520
         Width           =   1575
      End
      Begin VB.ComboBox cmbDomain 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1830
         Width           =   6495
      End
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   120
         MaxLength       =   100
         TabIndex        =   0
         Top             =   240
         Width           =   6555
      End
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1110
         Width           =   3855
      End
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   4020
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1110
         Width           =   2625
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Domain"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0084E8E8&
         Height          =   225
         Index           =   3
         Left            =   210
         TabIndex        =   14
         Top             =   2250
         Width           =   6495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0084E8E8&
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   12
         Top             =   750
         Width           =   6555
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0084E8E8&
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   11
         Top             =   1500
         Width           =   3855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0084E8E8&
         Height          =   225
         Index           =   2
         Left            =   3990
         TabIndex        =   10
         Top             =   1500
         Width           =   2625
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "POP3 Account"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0084E8E8&
      Height          =   420
      Left            =   1170
      TabIndex        =   13
      Top             =   120
      Width           =   2460
   End
   Begin VB.Image Image1 
      Height          =   1410
      Left            =   0
      Picture         =   "frmPOP3Account_new.frx":0442
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0084E8E8&
      BorderWidth     =   2
      X1              =   8130
      X2              =   60
      Y1              =   540
      Y2              =   540
   End
End
Attribute VB_Name = "frmPOP3Account"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public oCNT As clsSubscriber
Public frm As frmCustomerRec
Public PeriodFee As Currency
Public PerHour As Currency
Public PerMB As Currency
Public JoiningFee As Currency
Public DefShippingID As Long
Public Description As String
Public ptRecID As Long
Public ServiceID As Long
Public NumAdd As Long
Public SESSION As String
Public iCloseState As frm_CloseStates


Private Sub cmdAdd_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdAdd_Click"
    Const ContainerName = "frmPOP3Account"
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
    Static AmountAddedSoFar As Byte
    
    Dim itmX As ListItem
    
    If AmountAddedSoFar >= Me.NumAdd Then
        MsgBox "The maxium number of email addresses has been added!"
        Exit Sub
    Else
        AmountAddedSoFar = AmountAddedSoFar + 1
    End If
    
    If lvEmail.ListItems.Count > 0 Then
        Dim ix As Byte
        For ix = 1 To lvEmail.ListItems.Count
            If lvEmail.ListItems(ix).Text = txtField(1).Text And lvEmail.ListItems(ix).SubItems(1) = txtField(2).Text And lvEmail.ListItems(ix).SubItems(3) = cmbDomain.Text Then
                MsgBox "An Entry exactly like this is already in the list below, please revise.", vbExclamation, "Entry already exists"
                Exit Sub
            End If
        Next
    End If
    
    Dim rsCheck As adodb.Recordset
    
    If MySQL.OpenTable(directConn, rsCheck, , "select * from acci_services Where ServiceID = 4 and Username = '" & txtField(1) & "' and BaseURL = '" & cmbDomain.Text & "' Limit 1") = True Then
        If rsCheck.RecordCount > 0 Then
            MsgBox "This username for this service already exist on schemer.", vbCritical, "Username Exists"
            Exit Sub
        End If
    End If
    
    Dim lRecID_acci_services  As Long
    
    
    With oCNT.col_subServices.Add("NEW" & SESSION & oCNT.col_subServices.Count + 1, 0, Me.ptRecID, Me.ServiceID, txtField(0).Text, txtField(1).Text, txtField(2).Text, DateAdd("m", 1, sysnow), cmbDomain.Text, 0, sysnow, oCNT.fRecID, _
            IIf(chkSpam.Value = 1, "Scanning", "Standard"), "" & bNumAdd, "", "", "", True, Login.lVirtualID, sysnow, 0, 0, 0, sysnow, Me.PeriodFee, Me.PerHour, Me.PerMB, Me.JoiningFee, Login.lAgencyID, Me.DefShippingID, 0, sysnow, 1, "NEW" & SESSION & oCNT.col_subServices.Count + 1)
    
        ' Adds the Email to the Class and form listview
        Call oCNT.col_subEmails.Add("NEW" & SESSION & oCNT.col_subEmails.Count + 1, 0, oCNT.fRecID, 0, sysnow, .Username & "@" & .BaseURL, .ContactName, 0, True, "NEW" & SESSION & oCNT.col_subEmails.Count + 1)
        Set itmX = frm.lvEmail.ListItems.Add(, "NEW" & SESSION & oCNT.col_subEmails.Count, .ContactName)
        itmX.SubItems(1) = .Username & "@" & .BaseURL
        
        Set itmX = frm.lvPlans.ListItems.Add(, .Key, .ContactName)
        itmX.Checked = .Checked
        itmX.SubItems(1) = "POP3 Email Box"
        itmX.SubItems(2) = .Username
        itmX.SubItems(3) = .Password
        itmX.SubItems(4) = .BaseURL
        itmX.SubItems(5) = .DynamicField1
        itmX.SubItems(6) = .DynamicField2
        itmX.SubItems(7) = .DynamicField3
        itmX.SubItems(8) = .DynamicField4
        itmX.SubItems(9) = .DynamicField5
        
        Set itmX = lvEmail.ListItems.Add(, .Key, txtField(0))
        itmX.SubItems(1) = txtField(1).Text
        itmX.SubItems(2) = txtField(2).Text
        itmX.SubItems(3) = cmbDomain.Text
        itmX.Checked = IIf(chkSpam.Value = 1, True, False)
    
    End With
    
    Dim BillCycle As Byte
    For bx = frm.optBillingCycle.LBound To frm.optBillingCycle.UBound
        If frm.optBillingCycle(bx).Value = True Then
            BillCycle = bx
            Exit For
        End If
    Next
        
     With oCNT.col_subTrans.Add("EML" & oCNT.col_subTrans.Count + 1, 0, oCNT.fRecID, Me.PeriodFee + Me.JoiningFee, (Me.PeriodFee + Me.JoiningFee) * oTax(Login.TaxCode, Login.TaxCountry), _
            DateAdd(frm.optBillingCycle(BillCycle).Tag, Val(frm.txtBillingCycle(BillCycle).Text), sysnow), 0, "1899-12-31 12:00 AM", True, 0, (Me.PeriodFee + Me.JoiningFee) + ((Me.PeriodFee + Me.JoiningFee) * oTax(Login.TaxCode, Login.TaxCountry)), (Me.PeriodFee + Me.JoiningFee) + ((Me.PeriodFee + Me.JoiningFee) * oTax(Login.TaxCode, Login.TaxCountry)), _
            0, 0, Login.lAgencyID, Login.lVirtualID, Me.Description, 0, 0, oCNT.col_subServices.Count, 0, 0, Login.lSysopID, sysnow, 0, 0, 0, 0, sysnow, DateAdd("m", 1, sysnow), Me.ptRecID, Me.ServiceID, 0, 0, "EML" & oCNT.col_subTrans.Count + 1)
          
        Set itmX = frm.lvTransactions.ListItems.Add(, .Key, .Description)
        itmX.SubItems(1) = Format(.TotalDue, "Currency")
        itmX.SubItems(2) = Format(.GSTCharged, "Currency")
        itmX.SubItems(3) = Format(.PaymentDue, "yyyy-mm-dd ttttt")
        itmX.SubItems(4) = Format(.AmountPaid, "Currency")
        itmX.SubItems(5) = IIf(.AmountPaid + .AmountRefunded >= .TotalDue, "Yes", "No")
        itmX.Checked = .Checked
    
    End With
    
    
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

Private Sub cmdSave_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdSave_Click"
    Const ContainerName = "frmPOP3Account"
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


    iCloseState = frmCloseSave
    Unload Me
    
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

Private Sub cmdCancel_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdCancel_Click"
    Const ContainerName = "frmPOP3Account"
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


    iCloseState = frmCloseCancel
    Unload Me

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

Private Sub Command1_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Command1_Click"
    Const ContainerName = "frmPOP3Account"
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

Private Sub cmdUpdate_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdUpdate_Click"
    Const ContainerName = "frmPOP3Account"
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
    
    Dim rsCheck As adodb.Recordset
    
    If txtField(1).Tag <> txtField(1).Text Then
        If MySQL.OpenTable(directConn, rsCheck, , "select * from acci_services Where ServiceID = 4 and Username = '" & txtField(1).Text & "' and BaseURL = '" & cmbDomain.Text & "' Limit 1") = True Then
            If rsCheck.RecordCount > 0 Then
                MsgBox "Username for this service already exist on schemer.", vbCritical, "Username Exists"
                Exit Sub
            End If
        End If
    End If
    
    Select Case MsgBox("Are you sure you wish to save any changes that you have made to this email address?", vbYesNo + vbQuestion, "Save Changes")
    Case vbYes
        
        ' CHANGES
        'MySQL.Execute directConn, "Update acci_services Set ContactName='" & MySQL.ESC(txtField(0)) & "', BaseURL = '" & cmbDomain.Text & "', Username='" & MySQL.ESC(txtField(1)) & "', Password=AES_ENCRYPT('" & txtField(2) & "','" & odb.colSalts.ReturnSalt("md5Password") & "'), DynamicField1='" & IIf(chkSpam.Value = 1, "Scanning", "No Scan") & "' where RecID = " & Mid(lvEmail.SelectedItem.Key, 2)
        'MySQL.Execute directConn, "Update acci_aliases set dest='" & MySQL.ESC(txtField(1) + "@" + cmbDomain.Text) & "' where dest = '" & MySQL.ESC(lvEmail.SelectedItem.SubItems(1) + "@" + lvEmail.SelectedItem.SubItems(3)) & "'"
                
        With oCNT.col_subServices(lvEmail.SelectedItem.Key)
        
            Set itmX = lvEmail.SelectedItem
            itmX.SubItems(1) = txtField(1).Text
            itmX.SubItems(2) = txtField(2).Text
            itmX.SubItems(3) = cmbDomain.Text
            itmX.Checked = IIf(chkSpam.Value = 1, True, False)
        
            .ContactName = txtField(0).Text
            .Username = txtField(1).Text
            .Password = txtField(2).Text
            .BaseURL = cmbDomain.Text
            .DynamicField1 = IIf(itmX.Checked = True, "Scanning", "Standard")
            
            Set itmX = frm.lvPlans.ListItems(lvEmail.SelectedItem.Key)
            itmX.SubItems(1) = .Description
            itmX.SubItems(2) = .Username
            itmX.SubItems(3) = .Password
            itmX.SubItems(4) = .BaseURL
    
        End With
        
    End Select
    
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

Private Sub lvEmail_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvEmail_ItemClick"
    Const ContainerName = "frmPOP3Account"
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


    On Error Resume Next
    
    cmdUpdate.Tag = Item.Key
    txtField(0) = Item.Text
    txtField(1) = Item.SubItems(1)
    txtField(1).Tag = Item.SubItems(1)
    txtField(2) = Item.SubItems(2)
    cmbDomain.Text = Item.SubItems(3)
    cmdUpdate.Enabled = True
    
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

Private Sub txtField_GotFocus(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtField_GotFocus"
    Const ContainerName = "frmPOP3Account"
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


    txtField(Index).SelStart = 0
    txtField(Index).SelLength = Len(txtField(Index).Text)
    
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

Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtfield_KeyPress"
    Const ContainerName = "frmPOP3Account"
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


    Select Case KeyAscii
    Case Asc(" ")
        If Index = 1 Then KeyAscii = 0
    Case 13
        KeyAscii = 0
        SendKeys "{TAB}"
    End Select
            
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
Private Sub Form_Load()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Load"
    Const ContainerName = "frmPOP3Account"
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


    If SESSION = "" Then SESSION = GetSessionChar(SESSION, Me.hwnd, 14)
    
    Dim rsload As adodb.Recordset
    
    If bDebug = True Then On Error GoTo 0 Else On Error Resume Next
    
    GUI.LoadColWidths lvEmail, Me
    
    If MySQL.OpenTable(directConn, rsload, , "select count(*) as recordcount, Sum(DynamicField2) as NumAllowed from acci_services where ptRecID = " & Me.ptRecID & " and AccI_RecID = " & oCNT.fRecID) = True Then
    
        NumAdd = IIf(IsNull(rsload!NumAllowed), NumAdd, rsload!NumAllowed + NumAdd - rsload!RecordCount)
        bNumAdd = NumAdd
        
    Else
        'Stop
        bNumAdd = NumAdd
    End If
        
    txtField(0) = sContactName
    txtField(1) = sUsername
    txtField(2) = sPassword
    
    cmbDomain.AddItem Login.sVISPDomain
    
    'If Login.sVISPDomain <> "ep.net.au" Then cmbDomain.AddItem "ep.net.au"
    
                
    Dim oDM As cls_Domains
    For Each oDM In oCNT.col_Domains
        cmbDomain.AddItem oDM.Domain
    Next
    
    If cmbDomain.ListIndex = -1 Then cmbDomain.ListIndex = 0
    
    Dim itmX As ListItem
       
    Dim oEml As cls_subServices
    For Each oEml In oCNT.col_subServices
        If oEml.ServiceID = 4 Then
            Set itmX = lvEmail.ListItems.Add(, oEml.Key, oEml.ContactName)
            itmX.SubItems(1) = oEml.Username
            itmX.SubItems(2) = oEml.Password
            itmX.SubItems(3) = oEml.BaseURL
            itmX.Checked = IIf(oEml.DynamicField1 = "Scanning", True, False)
            rsload.MoveNext
        End If
    Next
        

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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_QueryUnload"
    Const ContainerName = "frmPOP3Account"
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


    GUI.SaveColWidths lvEmail, Me

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

