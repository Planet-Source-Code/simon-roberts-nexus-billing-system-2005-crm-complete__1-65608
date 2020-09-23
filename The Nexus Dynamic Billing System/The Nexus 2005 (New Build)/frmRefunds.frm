VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRefunds 
   BackColor       =   &H00A26F74&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Refund"
   ClientHeight    =   7635
   ClientLeft      =   1005
   ClientTop       =   2715
   ClientWidth     =   12195
   Icon            =   "frmRefunds.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   12195
   Begin VB.Frame Frame2 
      BackColor       =   &H00A26F74&
      Caption         =   "Items Invoiced so far"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5715
      Left            =   120
      TabIndex        =   3
      Top             =   1830
      Width           =   11985
      Begin MSComctlLib.ListView lvTransactions 
         Height          =   5295
         Left            =   90
         TabIndex        =   4
         Top             =   330
         Width           =   11805
         _ExtentX        =   20823
         _ExtentY        =   9340
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         PictureAlignment=   1
         _Version        =   393217
         ForeColor       =   16777215
         BackColor       =   10645364
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Description"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Total Due"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "GST"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Payment Due"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Amount Paid"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "All Paid"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Amount Currently Refunded"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame frame1 
      BackColor       =   &H00A26F74&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   90
      TabIndex        =   1
      Top             =   960
      Width           =   12015
      Begin VB.CommandButton cmdProcess 
         BackColor       =   &H00A26F74&
         Caption         =   "Process Refund"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   9240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   2685
      End
      Begin VB.CommandButton cmdAccount 
         BackColor       =   &H00A26F74&
         Caption         =   "Select Account Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   210
         Width           =   3285
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0084E8E8&
      BorderWidth     =   2
      Height          =   765
      Left            =   2070
      Top             =   90
      Width           =   10035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Refund"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   960
      TabIndex        =   0
      Top             =   30
      Width           =   1005
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   60
      Picture         =   "frmRefunds.frx":0BC2
      Stretch         =   -1  'True
      Top             =   30
      Width           =   840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0084E8E8&
      BorderWidth     =   2
      X1              =   -60
      X2              =   2070
      Y1              =   390
      Y2              =   390
   End
End
Attribute VB_Name = "frmRefunds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAccount_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdAccount_Click"
    Const ContainerName = "frmRefunds"
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
'**  Description:    Subroutine, Function or Property of project alpha                        **
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


    Dim fSearch As New frmSearchAccount
    Dim rsLoad As ADODB.Recordset
    Dim bResult As Boolean
    
    fSearch.Show 1
   
    If fSearch.lRecID <> 0 Then
    
        Frame1.Caption = fSearch.sAccountName
        Frame1.Tag = fSearch.lRecID
    
        bResult = MySQL.OpenTable(ADOConn, rsLoad, , "select * from invoiceout Where AccI_RecID = " & fSearch.lRecID)
    
        lvTransactions.ListItems.Clear
        
        If rsLoad.RecordCount > 0 Then
            rsLoad.MoveFirst
            While Not rsLoad.EOF And Err.Number = 0
                Set itmX = lvTransactions.ListItems.Add(, "r" & rsLoad!RecID, IIf(IsNull(rsLoad!Description), "", rsLoad!Description)) ')
                itmX.SubItems(1) = Format(rsLoad!TotalDue, "Currency")
                itmX.SubItems(2) = Format(rsLoad!GSTCharged, "Currency")
                itmX.SubItems(3) = rsLoad!PaymentDue
                itmX.SubItems(4) = Format(rsLoad!AmountPaid, "Currency")
                itmX.SubItems(5) = IIf(rsLoad!AmountPaid + rsLoad!AmountRefunded >= rsLoad!TotalDue, "Yes", "No")
                itmX.SubItems(6) = Format(IIf(IsNull(rsLoad!AmountRefunded + rsLoad!GSTRefunded), 0, rsLoad!AmountRefunded + rsLoad!GSTRefunded), "Currency")
                'mX.SubItems(7) = Format(IIf(IsNull(rsLoad!GSTRefunded), 0, rsLoad!GSTRefunded), "Currency")
                itmX.Checked = IIf(rsLoad!Checked <> 0, True, False)
                rsLoad.MoveNext
            Wend
        End If

    
    
    End If
    
        
Exit Sub



ErrorOccur:
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Sub

Private Sub cmdProcess_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdProcess_Click"
    Const ContainerName = "frmRefunds"
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
'**  Description:    Subroutine, Function or Property of project alpha                        **
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


    Dim cAmount As Currency
    Dim cAmountPaid As Currency
    Dim cRefunded As Currency
    Dim cDiff As Currency
    Dim cDiffAmount As Currency
    Dim rsTraxr As ADODB.Recordset
    Dim rsLoad As ADODB.Recordset
    
    Dim ix As Long
    Dim itmX As ListItem
    If lvTransactions.ListItems.Count > 0 Then
        For ix = 1 To lvTransactions.ListItems.Count
            Set itmX = lvTransactions.ListItems(ix)
            If itmX.Checked = True Then
                cAmount = cAmount + CCur(Mid(itmX.SubItems(1), 2))
                cAmountPaid = cAmountPaid + CCur(Mid(itmX.SubItems(4), 2))
                cRefunded = cRefunded + CCur(Mid(itmX.SubItems(6), 2))
            End If
        Next
        
        If cAmount <= 0 Then
            MsgBox "Zero Dollar Amount selected"
            Exit Sub
        End If
        
        Dim fAmount As New frmRefundAmount
        
        fAmount.cAmountDue = cAmount
        fAmount.cAmountPaid = cAmountPaid
        fAmount.cRefunded = cRefunded
        fAmount.Show 1
        
        If fAmount.cAmountRefunded = 0 Then Exit Sub
        
        'bResult = MySQL.OpenTable(ADOConn, rsTraxr, , "select * from refundtraxr Limit 1")
        'rsTraxr.AddNew
        Dim rTraxrId As Long
        
        On Error Resume Next
        Do
            Err.Clear
            rTraxrId = MySQL.GetTMPRecID("refundtraxr", ADOConn)
                                      
            ADOConn.Execute "INSERT INTO refundtraxr (RecID, RefundSerial, acci_RecID, RefundDate, Total, SysopID, VirtualID, AmountRefunded) " + _
                          "VALUES ('" & rTraxrId & "','" & Hex(rTraxrId) & "','" & Frame1.Tag & "','" & Format(sysNOW, "yyyy-mm-dd ttttt") & "','" & fAmount.cAmountRefunded & "','" & Login.lSysopID & "','" & Login.lVirtualID & "','" & fAmount.cAmountRefunded & "')"
            If Err.Number > 0 Then cDebug Err.Description
        Loop Until Err.Number = 0
        
        Dim InvInID As Long
        
        Select Case fAmount.bType
        Case 0 'Credit Account
            'bResult = MySQL.OpenTable(ADOConn, rsLoad, , "select * from invoicein Limit 1")
            On Error Resume Next
            Do
                Err.Clear
                InvInID = MySQL.GetTMPRecID("invoicein", ADOConn)
                ADOConn.Execute "INSERT INTO invoicein (RecID, acci_RecID, AmountPaid, VirtualID, Checked, AmountUsed, GSTCharged, SysopID, RefundTRXID) " + _
                              "VALUES ('" & InvInID & "','" & Frame1.Tag & "','" & fAmount.cAmountRefunded & "','" & Login.lVirtualID & "','" & -1 & "','" & 0 & "','" & 0 & "','" & Login.lSysopID & "','" & rTraxrId & "')"
            Loop Until Err.Number = 0
            
            MySQL.AddReceiptItem ADOConn, Frame1.Tag, , InvInID, , rTraxrId, , fAmount.cAmountRefunded, "Account Credit"
        Case 1 'Fiscal Amount
            
            MySQL.AddReceiptItem ADOConn, Frame1.Tag, , InvInID, , rTraxrId, , fAmount.cAmountRefunded, "Fiscal Transaction"
        End Select
        
        
        
        For ix = 1 To lvTransactions.ListItems.Count
            Set itmX = lvTransactions.ListItems(ix)
            If itmX.Checked = True Then
                Call MySQL.OpenTable(ADOConn, rsLoad, , "select acci_RecID, TotalDue, AmountRefunded, AmountPaid from invoiceout where RecID =" & Mid(itmX.Key, 2))
                cDiff = rsLoad!TotalDue - rsLoad!AmountRefunded - rsLoad!AmountPaid
                If cDiff < fAmount.cAmountRefunded And cDiff <> 0 Then
                    MySQL.Execute ADOConn, "Update invoiceout Set AmountRefunded = AmountRefunded + " & cDiff & " where RecID = " & Mid(itmX.Key, 2)
                    MySQL.Execute ADOConn, "Update invoiceout Set GSTRefunded = AmountRefunded * " & oTax(Login.TaxCode, Login.TaxCountry) & " where RecID = " & Mid(itmX.Key, 2)
                    MySQL.Execute ADOConn, "Update invoiceout Set RefundID = " & rTraxrId & " where RecID = " & Mid(itmX.Key, 2)
                    MySQL.Execute ADOConn, "INSERT INTO flags_refunds (acciServiceID, Refunded, GST, SysopID, VirtualID, acci_RecID, RefundTRXID) VALUES ('" & rsLoad!RecID & "','" & cDiff & "','" & cDiff * oTax(Login.TaxCode, Login.TaxCountry) & "','" & Login.lSysopID & "','" & Login.lVirtualID & "','" & rsLoad!acci_RecID & "','" & rTraxrId & "')"
                    fAmount.cAmountRefunded = fAmount.cAmountRefunded - cDiff
                Else
                    MySQL.Execute ADOConn, "Update invoiceout Set AmountRefunded = AmountRefunded + " & fAmount.cAmountRefunded & " where RecID = " & Mid(itmX.Key, 2)
                    MySQL.Execute ADOConn, "Update invoiceout Set GSTRefunded = AmountRefunded * " & oTax(Login.TaxCode, Login.TaxCountry) & " where RecID = " & Mid(itmX.Key, 2)
                    MySQL.Execute ADOConn, "Update invoiceout Set RefundID = " & rTraxrId & " where RecID = " & Mid(itmX.Key, 2)
                    MySQL.Execute ADOConn, "INSERT INTO flags_refunds (acciServiceID, Refunded, GST, SysopID, VirtualID, acci_RecID, RefundTRXID) VALUES ('" & rsLoad!RecID & "','" & fAmount.cAmountRefunded & "','" & fAmount.cAmountRefunded * oTax(Login.TaxCode, Login.TaxCountry) & "','" & Login.lSysopID & "','" & Login.lVirtualID & "','" & rsLoad!acci_RecID & "','" & rTraxrId & "')"
                    itmX.SubItems(6) = CCur(itmX.SubItems(6)) + (fAmount.cAmountRefunded + fAmount.cAmountRefunded * oTax(Login.TaxCode, Login.TaxCountry))
                    bDone = False
                    Exit For
                End If
                'bResult = MySQL.OpenTable(ADOConn, rsLoad, , "select AmountRefunded, GSTRefunded from invoiceout Where RecID = " & Mid(itmX.Key, 2))
                itmX.SubItems(6) = CCur(itmX.SubItems(6)) + (fAmount.cAmountRefunded + fAmount.cAmountRefunded * oTax(Login.TaxCode, Login.TaxCountry))
                
                bDone = False
                If rsLoad!TraxrID <> 0 Then
                    MySQL.Execute ADOConn, "UPDATE invoicetraxr SET AmountCredited=AmountCredited+" & fAmount.cAmountRefunded + fAmount.cAmountRefunded * oTax(Login.TaxCode, Login.TaxCountry) & " where RecID = " & rsLoad!TraxrID
                    bDone = True
                End If
                              
                rsLoad.Update
            End If
        Next
        
        If rsLoad!TraxrID <> 0 And bDone = False Then
            MySQL.Execute ADOConn, "UPDATE invoicetraxr SET AmountCredited=AmountCredited+" & fAmount.cAmountRefunded + fAmount.cAmountRefunded * oTax(Login.TaxCode, Login.TaxCountry) & " where RecID = " & rsLoad!TraxrID
        End If
    End If
    
    
Exit Sub



ErrorOccur:
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
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
    Const ContainerName = "frmRefunds"
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
'**  Description:    Subroutine, Function or Property of project alpha                        **
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


    Call GUI.LoadColWidths(lvTransactions, Me)
    
Exit Sub

    If bBigFont = True Then lvTransactions.Font.Size = 18
    

ErrorOccur:
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
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
    Const ContainerName = "frmRefunds"
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
'**  Description:    Subroutine, Function or Property of project alpha                        **
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


    Call GUI.SaveColWidths(lvTransactions, Me)
    
Exit Sub



ErrorOccur:
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Sub

Private Sub lvTransactions_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvTransactions_ColumnClick"
    Const ContainerName = "frmRefunds"
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
'**  Description:    Subroutine, Function or Property of project alpha                        **
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


    Call GUI.ColumnSort(ColumnHeader, lvTransactions)
    
Exit Sub



ErrorOccur:
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Sub
