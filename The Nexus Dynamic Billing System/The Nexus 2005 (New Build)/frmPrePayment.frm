VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmPrepayment 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Make Prepayment"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8520
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Payment Type"
      Height          =   2535
      Left            =   4590
      TabIndex        =   14
      Top             =   120
      Width           =   3825
      Begin MSComctlLib.ListView lvPaymentType 
         Height          =   2145
         Left            =   120
         TabIndex        =   15
         Tag             =   "0"
         Top             =   270
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   3784
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Description"
            Object.Width           =   4763
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Sub"
            Object.Width           =   1411
         EndProperty
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4590
      TabIndex        =   3
      Top             =   5670
      Width           =   3825
   End
   Begin VB.CommandButton cmdMakePayment 
      Caption         =   "&Make Payment"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4590
      TabIndex        =   2
      Top             =   5100
      Width           =   3825
   End
   Begin VB.Frame Frame2 
      Height          =   2265
      Left            =   4590
      TabIndex        =   1
      Top             =   2760
      Width           =   3825
      Begin VB.CheckBox chkGST 
         Caption         =   "&GST: $0.00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   12
         Top             =   810
         Visible         =   0   'False
         Width           =   3345
      End
      Begin VB.TextBox txtSum 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1590
         TabIndex        =   11
         Top             =   330
         Width           =   1995
      End
      Begin VB.Label lblTotal 
         Caption         =   "Sub: $"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   540
         TabIndex        =   16
         Top             =   1260
         Width           =   3120
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   1
         X1              =   1350
         X2              =   3690
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         Caption         =   "Total: $"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   210
         TabIndex        =   13
         Top             =   1650
         Width           =   1320
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   0
         X1              =   1350
         X2              =   3690
         Y1              =   1170
         Y2              =   1170
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Sum: $"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   525
         TabIndex        =   10
         Top             =   360
         Width           =   990
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Account"
      Height          =   5805
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   4305
      Begin VB.Frame Frame5 
         Caption         =   "Search Options"
         Height          =   1395
         Left            =   180
         TabIndex        =   5
         Top             =   210
         Width           =   3945
         Begin VB.TextBox txtSearchName 
            Height          =   285
            Left            =   180
            TabIndex        =   9
            ToolTipText     =   "Shift + Enter to search database"
            Top             =   300
            Width           =   3585
         End
         Begin VB.CheckBox chkSearchAccountNames 
            Caption         =   "Search Account Names"
            Height          =   255
            Left            =   210
            TabIndex        =   8
            Top             =   660
            Value           =   1  'Checked
            Width           =   3105
         End
         Begin VB.CheckBox chkSearchContacts 
            Caption         =   "Search All Contact Names"
            Height          =   255
            Left            =   210
            TabIndex        =   7
            Top             =   930
            Width           =   3045
         End
         Begin MSComctlLib.ProgressBar pb1 
            Height          =   105
            Left            =   60
            TabIndex        =   6
            Top             =   1200
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   185
            _Version        =   393216
            Appearance      =   0
            Max             =   6
         End
      End
      Begin MSComctlLib.ListView lvAccount 
         Height          =   3945
         Left            =   150
         TabIndex        =   4
         Top             =   1740
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   6959
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Account Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Contact Name"
            Object.Width           =   2028
         EndProperty
      End
   End
End
Attribute VB_Name = "frmPrepayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lRecID As Long
Public cTotal As Currency
Public sAccountName As String
Public cGST As Currency
Public cSub As Currency

Public iCloseState As frm_CloseStates


Private Sub cmdCancel_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdCancel_Click"
    Const ContainerName = "frmPrepayment"
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
Private Sub cmdmakePayment_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdmakePayment_Click"
    Const ContainerName = "frmPrepayment"
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


    Dim ix As Integer
    Dim bselected As Boolean
    Dim bSaveCC As Boolean
    Dim rs_creditcard As adodb.Recordset
    Dim rs_cc_Receipt As adodb.Recordset
    
    If lvAccount.ListItems.Count = 0 Then
        MsgBox "No Account name specified", vbInformation, "Search and select an Account"
        Exit Sub
    Else
        For ix = 1 To lvAccount.ListItems.Count
            If lvAccount.ListItems(ix).Checked = True Then bselected = True
        Next ix
        
        If bselected = False Then
            MsgBox "No Account name specified", vbInformation, "Search and select an Account"
            Exit Sub
        End If
    End If
    
    Dim bChecked As Boolean
    
    For ix = 1 To lvPaymentType.ListItems.Count
        If lvPaymentType.ListItems(ix).Checked = True Then
            bChecked = True
        End If
    Next
    
    If bChecked = False Then
        MsgBox "You haven't select a payment type yet!"
        Exit Sub
    End If
    
    Dim fFrmCC As New frmCreditCard
    Dim iY As Variant
    
    bselected = False
    For ix = 1 To lvPaymentType.ListItems.Count
        If lvPaymentType.ListItems(ix).Checked = True Then
            bselected = True
            If Left(lvPaymentType.ListItems(ix).Key, 1) = "c" Then
                
                '¤¤¤ Query Local DB for Previous History
                '
                ' --> HERE
                '
                '¤
                
                For iY = 1 To lvAccount.ListItems.Count
                    If lvAccount.ListItems(iY).Checked = True Then
                        If Left(lvAccount.ListItems(iY).Key, 1) = "a" Then
                            bResult = MySQL.OpenTable(directConn, rs_creditcard, , "select RecID, AccI_RecID as acci_RecID, bType, AES_DECRYPT(CardNumber,'" + odb.colSalts.ReturnSalt("CCSalt") + "" & odb.colSalts.ReturnSalt("md5Password") & "') as CardNumber, SecurityNumber, ExpiryDate, Name from creditcard where AccI_RecID = " & Mid(lvAccount.ListItems(iY).Key, 2))
                            bResult = MySQL.OpenTable(directConn, rs_cc_Receipt, , "select * from cc_Receipt limit 1")
                            Exit For
                        Else
                            bResult = MySQL.OpenTable(directConn, rs_creditcard, , "select RecID, AccI_RecID as acci_RecID, bType, AES_DECRYPT(CardNumber,'" + odb.colSalts.ReturnSalt("CCSalt") + "" & odb.colSalts.ReturnSalt("md5Password") & "') as CardNumber, SecurityNumber, ExpiryDate, Name from creditcard where AccI_RecID = " & Mid(lvAccount.ListItems(iY).Key, InStr(lvAccount.ListItems(iY).Key, "_") + 1))
                            bResult = MySQL.OpenTable(directConn, rs_cc_Receipt, , "select * from cc_Receipt limit 1")
                            Exit For
                        End If
                    End If
                Next iY
                
                If rs_creditcard.RecordCount > 0 Then
                    rs_creditcard.MoveFirst
                    While Not rs_creditcard.EOF And Err.Number = 0
                        fFrmCC.cmdName.AddItem rs_creditcard!Name
                        fFrmCC.cmdName.ItemData(fFrmCC.cmdName.ListCount - 1) = rs_creditcard!RecID
                        rs_creditcard.MoveNext
                    Wend
                    fFrmCC.cmdName.ListIndex = 0
                    rs_creditcard.Filter = ""
                Else
                    rs_creditcard.Filter = ""
                End If
                
                fFrmCC.Show 1
                If fFrmCC.iCloseState = frmCloseCancel Then Exit Sub
                bSaveCC = True
            End If
        End If
    Next ix
    
    If bselected = False Then
        MsgBox "One or more payment types are not selected.", vbInformation, "Check a payment type."
        Exit Sub
    End If
    
    If IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax(Login.TaxCode, Login.TaxCountry), 0) = 0 Then
        MsgBox "Zero value specified in sum of transaction field!", vbCritical, "Zero Value"
    End If
    
    Dim rsSave As adodb.Recordset
    Dim lacci_RecID As Variant
    
    On Error Resume Next
    Do
        Err.Clear
        lRecID = MySQL.GetTMPRecID("invoicein", directConn)
        Call MySQL.Execute(directConn, "insert into invoicein (RecID) Values ('" & lRecID & "')")
    Loop Until Err.Number = 0
    
    For ix = 1 To lvAccount.ListItems.Count
        If lvAccount.ListItems(ix).Checked = True Then
            If Left(lvAccount.ListItems(ix).Key, 1) = "a" Then
                Call MySQL.Execute(directConn, "update invoicein set acci_RecID = '" & Mid(lvAccount.ListItems(ix).Key, 2) & "' where RecID = '" & lRecID & "'")
                lacci_RecID = Mid(lvAccount.ListItems(ix).Key, 2)
            Else
                Call MySQL.Execute(directConn, "update invoicein set acci_RecID = '" & Mid(lvAccount.ListItems(ix).Key, InStr(lvAccount.ListItems(ix).Key, "_") + 1) & "' where RecID = '" & lRecID & "'")
                lacci_RecID = Mid(lvAccount.ListItems(ix).Key, InStr(lvAccount.ListItems(ix).Key, "_") + 1)
            End If
            sAccountName = lvAccount.ListItems(ix).Text
        End If
    Next ix
    
    
    Call MySQL.Execute(directConn, "update invoicein set AmountPaid = '" & Val(txtSum) & "' where RecID = '" & lRecID & "'")
    cTotal = Val(txtSum)
    Call MySQL.Execute(directConn, "update invoicein set sub = '" & lvPaymentType.Tag & "' where RecID = '" & lRecID & "'")
    cSub = lvPaymentType.Tag
    
    If chkGST.Value = 1 Then
        Call MySQL.Execute(directConn, "update invoicein set GSTCharged = '" & IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax(Login.TaxCode, Login.TaxCountry), 0) & "' where RecID = '" & lRecID & "'")
        cGST = IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax(Login.TaxCode, Login.TaxCountry), 0)
    Else
        Call MySQL.Execute(directConn, "update invoicein set GSTCharged = '0' where RecID = '" & lRecID & "'")
        cGST = 0
        'lblTotal(1).Caption = "Round: $ " & Format(Round(IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax("G","AUS0001"), 0), 2), "###,###,###,###,###.##")
    End If
    
    Call MySQL.Execute(directConn, "update invoicein set TotalPaid = '" & (Val(txtSum) + lvPaymentType.Tag + cGST) & "' where RecID = '" & lRecID & "'")
    Call MySQL.Execute(directConn, "update invoicein set VirtualID = '" & Login.lVirtualID & "' where RecID = '" & lRecID & "'")
    Call MySQL.Execute(directConn, "update invoicein set AmountUsed = '0' where RecID = '" & lRecID & "'")
    
    
    SaveFlags lRecID
    
    
    For ix = 1 To lvPaymentType.ListItems.Count
        If lvPaymentType.ListItems(ix).Checked = True Then
            If Left(lvPaymentType.ListItems(ix).Key, 1) = "c" Then
                MySQL.AddReceiptItem directConn, lacci_RecID, , lRecID, , , rsSave!TotalPaid, , "CC: " + Right(fFrmCC.screditcard, 5)
            Else
                MySQL.AddReceiptItem directConn, lacci_RecID, , lRecID, , , rsSave!TotalPaid, , lvPaymentType.ListItems(ix).Text + " Payment"
            End If
        End If
    Next

    '¤
    If bSaveCC = True Then
        rs_creditcard.Filter = "AccI_RecID = " & rsSave!acci_RecID & " AND CardNumber = '" & MySQL.NumCrypt(fFrmCC.screditcard) & "'"
        If rs_creditcard.RecordCount = -1 Or rs_creditcard.RecordCount = 0 Then
            rs_creditcard.Filter = ""
            rs_creditcard.AddNew
            rs_creditcard!acci_RecID = rsSave!acci_RecID
            rs_creditcard!SecurityNumber = fFrmCC.sSecurityNo
            rs_creditcard!Name = fFrmCC.sCardName
            rs_creditcard!ExpiryDate = fFrmCC.sExpiry
            rs_creditcard!bType = fFrmCC.bType
            rs_creditcard!bDefault = fFrmCC.bDefault
            
            MySQL.Execute directConn, "UPDATE creditcard SET CardNumber=AES_ENCRYPT('" & MySQL.NumCrypt(fFrmCC.screditcard) & "','" + odb.colSalts.ReturnSalt("CCSalt") + "" & odb.colSalts.ReturnSalt("md5Password") & "') where RecID = " & MySQL.SetRecID(rs_creditcard, "creditcard", directConn)
            
            If fFrmCC.bDefault = True Then
                MySQL.Execute directConn, "UPDATE creditcard SET bDefault=0 where acci_RecID = " & rs_creditcard!acci_RecID
                MySQL.Execute directConn, "UPDATE creditcard SET bDefault=-1 where RecID = " & rs_creditcard!RecID
            End If
            
            rs_cc_Receipt.AddNew
            rs_cc_Receipt!cc_RecID = rs_creditcard!RecID
            rs_cc_Receipt!ReceiptNumber = fFrmCC.sReceiptNo
            rs_cc_Receipt.Update
        Else
            
            rs_creditcard!bDefault = fFrmCC.bDefault
            rs_creditcard.Update
            
            rs_cc_Receipt.AddNew
            rs_cc_Receipt!cc_RecID = rs_creditcard!RecID
            rs_cc_Receipt!ReceiptNumber = fFrmCC.sReceiptNo
            rs_cc_Receipt.Update
            rs_creditcard.Filter = ""
        End If
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


Sub SaveFlags(lRecID As Variant)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "SaveFlags"
    Const ContainerName = "frmPrepayment"
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


    Dim rsSave As adodb.Recordset
    Dim bResult As Boolean
    
    bResult = MySQL.OpenTable(directConn, rsSave, , "select * from flags_invoicein Limit 1")
        
    For ix = 1 To lvPaymentType.ListItems.Count
        If lvPaymentType.ListItems(ix).Checked = True Then
            Call MySQL.Execute(directConn, "insert into flags_invoicein (InvIn_RecID, Flag) VALUES('" & lRecID & "','" & Mid(lvPaymentType.ListItems(ix).Key, 2) & "')")
        End If
    Next ix
    
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
Public Function LoadFlags(Optional ltmpRecID As Long)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "LoadFlags"
    Const ContainerName = "frmPrepayment"
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


    Dim rsload As adodb.Recordset
    Dim itmX As ListItem
    Dim bResult As Boolean
    
    bResult = MySQL.OpenTable(directConn, rsload, , "select * from paymenttype")
    
    If rsload.RecordCount > 0 Then
        rsload.MoveFirst
        While Not rsload.EOF And Err.Number = 0
            Set itmX = lvPaymentType.ListItems.Add(, IIf(rsload!CreditCard = True, "c", "a") & rsload!RecID, rsload!Description)
            itmX.SubItems(1) = IIf(rsload!sub = 0, "", Format(rsload!sub, "Currency"))
            itmX.Tag = rsload!sub
            rsload.MoveNext
        Wend
    End If

    Dim ix As Integer
    
    If ltmpRecID <> 0 Then
        bResult = MySQL.OpenTable(directConn, rsload, , "select * from Flags_invoicein Where InvIn_RecID = " & ltmpRecID)
            
        If rsload.RecordCount > 0 Then
            While Not rsload.EOF And Err.Number = 0
                For ix = 1 To lvPaymentType.ListItems.Count
                    If Val(Mid(lvPaymentType.ListItems(ix).Key, 2)) = rsload!Flag Then
                        lvPaymentType.ListItems(ix).Checked = True
                        Exit For
                    End If
                Next
                rsload.MoveNext
            Wend
            
        End If
    End If

Exit Function



ErrorOccur:
Select Case oErr.chkError(directConn, Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function
Private Sub Form_Load()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Load"
    Const ContainerName = "frmPrepayment"
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


    If lRecID <> 0 Then
    
        Call LoadInformation
        
    Else
        
        Call LoadFlags
        
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

Private Sub chkGST_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "chkGST_Click"
    Const ContainerName = "frmPrepayment"
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


    If chkGST.Value = 1 Then
        lblTotal(0).Caption = "Total: $ " & IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax(Login.TaxCode, Login.TaxCountry), 0) + Val(txtSum)
        'lblTotal(1).Caption = "Round: $ " & Round(IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax("G","AUS0001"), 0) + Val(txtSum), 2)
    Else
        lblTotal(0).Caption = "Total: $ " & Val(txtSum)
        'lblTotal(1).Caption = "Round: $ " & Round(IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax("G","AUS0001"), 0), 2)
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

Private Sub lvAccount_ItemCheck(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvAccount_ItemCheck"
    Const ContainerName = "frmPrepayment"
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


    Dim lx As Variant
    
    For lx = 1 To lvAccount.ListItems.Count
        lvAccount.ListItems(lx).Checked = False
    Next
    
    Item.Checked = True
    
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

Private Sub lvpaymenttype_ItemCheck(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvpaymenttype_ItemCheck"
    Const ContainerName = "frmPrepayment"
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


    Select Case Item.Checked
    Case True
        lvPaymentType.Tag = lvPaymentType.Tag + Item.Tag
    Case False
        lvPaymentType.Tag = lvPaymentType.Tag - Item.Tag
    End Select
    
    lblTotal(1).Caption = "Sub: " + Format(lvPaymentType.Tag, "Currency")
    
    If chkGST.Value = 1 Then
        lblTotal(0).Caption = "Total: " + Format(IIf(IsNumeric(txtSum + Chr(KeyAscii)) = True, Val(txtSum + Chr(KeyAscii)) * oTax(Login.TaxCode, Login.TaxCountry), 0) + Val(txtSum + Chr(KeyAscii)) + lvPaymentType.Tag, "Currency")
        'lblTotal(1).Caption = "Round: $ " & Format(Round(IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax("G","AUS0001"), 0) + Val(txtSum), 2), "###,###,###,###,###.##")
    Else
        lblTotal(0).Caption = "Total: " & Format(Val(txtSum + Chr(KeyAscii)) + lvPaymentType.Tag, "Currency")
        'lblTotal(1).Caption = "Round: $ " & Format(Round(IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax("G","AUS0001"), 0), 2), "###,###,###,###,###.##")
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

Private Sub txtSearchName_KeyPress(KeyAscii As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtSearchName_KeyPress"
    Const ContainerName = "frmPrepayment"
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



    Dim rsload As adodb.Recordset
    
    Select Case KeyAscii
    Case 13
        KeyAscii = 0
        pb1.Value = 0
        If Trim(txtSearchName) = "" Then Exit Sub
               
        If chkSearchAccountNames.Value <> 0 Then
            bResult = MySQL.OpenTable(directConn, rsload, , MySQL.virtualisp("select * from accountinfo Where AccountName Like '%" & txtSearchName.Text & "%'", "accountinfo"))
            If rsload.State = adStateOpen Then
                If rsload.RecordCount > 0 Then
                    rsload.MoveFirst
                    While Not rsload.EOF And Err.Number = 0
                        Set itmX = lvAccount.ListItems.Add(, "a" & rsload!RecID, rsload!AccountName)
                        rsload.MoveNext
                    Wend
                End If
            End If
        End If
        
        pb1.Value = pb1.Value + 1
        
        If chkSearchContacts.Value <> 0 Then
        
            bResult = MySQL.OpenTable(directConn, rsload, , MySQL.virtualisp("select accountinfo.AccountName, acci_addresses.Contactname, acci_addresses.AccI_RecID, acci_addresses.RecID from accountinfo, acci_addresses Where acci_addresses.AccI_RecID = accountinfo.RecID AND acci_addresses.ContactName Like '%" & txtSearchName.Text & "%'", "accountinfo"))
            
            If rsload.State = adStateOpen Then
                If rsload.RecordCount > 0 Then
                    
                    rsload.MoveFirst
                    While Not rsload.EOF And Err.Number = 0
                        Set itmX = lvAccount.ListItems.Add(, "d" & rsload!RecID & "_" & rsload!acci_RecID, rsload!AccountName)
                        itmX.SubItems(1) = rsload!ContactName
                        rsload.MoveNext
                    Wend
                End If
            End If
        End If
        pb1.Value = pb1.Value + 1
        
        If chkSearchContacts.Value <> 0 Then
        
            bResult = MySQL.OpenTable(directConn, rsload, , MySQL.virtualisp(MySQL.virtualisp("select accountinfo.AccountName, acci_emailaddresses.Contactname, acci_emailaddresses.AccI_RecID, acci_emailaddresses.RecID from accountinfo, acci_emailaddresses Where acci_emailaddresses.AccI_RecID = accountinfo.RecID AND acci_emailaddresses.ContactName Like '%" & txtSearchName.Text & "%'", "accountinfo")))
            If rsload.State = adStateOpen Then
                If rsload.RecordCount > 0 Then
                    rsload.MoveFirst
                    While Not rsload.EOF And Err.Number = 0
                        Set itmX = lvAccount.ListItems.Add(, "d" & rsload!RecID & "_" & rsload!acci_RecID, rsload!AccountName)
                        itmX.SubItems(1) = rsload!ContactName
                        rsload.MoveNext
                    Wend
                End If
            End If
        End If
        pb1.Value = pb1.Value + 1
        
        If chkSearchContacts.Value <> 0 Then
        
            bResult = MySQL.OpenTable(directConn, rsload, , "select accountinfo.AccountName, acci_ftpaccounts.Contactname, acci_ftpaccounts.AccI_RecID, acci_ftpaccounts.RecID from accountinfo, acci_ftpaccounts Where acci_ftpaccounts.AccI_RecID = accountinfo.RecID AND acci_ftpaccounts.ContactName Like '%" & txtSearchName.Text & "%'")
            If rsload.State = adStateOpen Then
                If rsload.RecordCount > 0 Then
                    rsload.MoveFirst
                    While Not rsload.EOF And Err.Number = 0
                        Set itmX = lvAccount.ListItems.Add(, "d" & rsload!RecID & "_" & rsload!acci_RecID, rsload!AccountName)
                        itmX.SubItems(1) = rsload!ContactName
                        rsload.MoveNext
                    Wend
                End If
            End If
        End If
        pb1.Value = pb1.Value + 1
        
        If chkSearchContacts.Value <> 0 Then
        
            bResult = MySQL.OpenTable(directConn, rsload, , "select accountinfo.AccountName, acci_pop3accounts.Contactname, acci_pop3accounts.AccI_RecID, acci_pop3accounts.RecID from accountinfo, acci_pop3accounts Where acci_pop3accounts.AccI_RecID = accountinfo.RecID AND acci_pop3accounts.ContactName Like '%" & txtSearchName.Text & "%'")
            
            If rsload.State = adStateOpen Then
                If rsload.RecordCount > 0 Then
                    rsload.MoveFirst
                    While Not rsload.EOF And Err.Number = 0
                        Set itmX = lvAccount.ListItems.Add(, "d" & rsload!RecID & "_" & rsload!acci_RecID, rsload!AccountName)
                        itmX.SubItems(1) = rsload!ContactName
                        rsload.MoveNext
                    Wend
                End If
            End If
            
        End If
        pb1.Value = pb1.Value + 1
        
        If chkSearchContacts.Value <> 0 Then
        
            bResult = MySQL.OpenTable(directConn, rsload, , MySQL.virtualisp("select accountinfo.AccountName, acci_phonenumbers.Contactname, acci_phonenumbers.AccI_RecID, acci_phonenumbers.RecID from accountinfo, acci_phonenumbers Where acci_phonenumbers.AccI_RecID = accountinfo.RecID AND acci_phonenumbers.ContactName Like '%" & txtSearchName.Text & "%'", "accountinfo"))
            
            If rsload.State = adStateOpen Then
                If rsload.RecordCount > 0 Then
                    rsload.MoveFirst
                    While Not rsload.EOF And Err.Number = 0
                        Set itmX = lvAccount.ListItems.Add(, "d" & rsload!RecID & "_" & rsload!acci_RecID, rsload!AccountName)
                        itmX.SubItems(1) = rsload!ContactName
                        rsload.MoveNext
                    Wend
                End If
            End If
        End If
        pb1.Value = pb1.Value + 1
       
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


Private Sub txtSum_Change()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtSum_Change"
    Const ContainerName = "frmPrepayment"
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


    
    chkGST.Caption = "GST: " & Format(Round(IIf(IsNumeric(txtSum + Chr(KeyAscii)) = True, Val(txtSum + Chr(KeyAscii)) * oTax(Login.TaxCode, Login.TaxCountry), 0), 2), "Currency")
    chkGST.Tag = IIf(IsNumeric(txtSum + Chr(KeyAscii)) = True, Val(txtSum + Chr(KeyAscii)) * oTax(Login.TaxCode, Login.TaxCountry), 0)
    
    If chkGST.Value = 1 Then
        lblTotal(0).Caption = "Total: " + Format(IIf(IsNumeric(txtSum + Chr(KeyAscii)) = True, Val(txtSum + Chr(KeyAscii)) * oTax(Login.TaxCode, Login.TaxCountry), 0) + Val(txtSum + Chr(KeyAscii)) + lvPaymentType.Tag, "Currency")
        'lblTotal(1).Caption = "Round: $ " & Format(Round(IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax("G","AUS0001"), 0) + Val(txtSum), 2), "###,###,###,###,###.##")
    Else
        lblTotal(0).Caption = "Total: " & Format(Val(txtSum + Chr(KeyAscii)) + lvPaymentType.Tag, "Currency")
        
        'lblTotal(1).Caption = "Round: $ " & Format(Round(IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax("G","AUS0001"), 0), 2), "###,###,###,###,###.##")
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

Private Sub txtSum_KeyPress(KeyAscii As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtSum_KeyPress"
    Const ContainerName = "frmPrepayment"
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


    If txtSum.Locked = True Then Exit Sub
    
    Select Case KeyAscii
    Case 8
    Case 48 To 57
        KeyAscii = KeyAscii
        chkGST.Caption = "GST: " & Format(Round(IIf(IsNumeric(txtSum + Chr(KeyAscii)) = True, Val(txtSum + Chr(KeyAscii)) * oTax(Login.TaxCode, Login.TaxCountry), 0), 2), "Currency")
        chkGST.Tag = IIf(IsNumeric(txtSum + Chr(KeyAscii)) = True, Val(txtSum + Chr(KeyAscii)) * oTax(Login.TaxCode, Login.TaxCountry), 0)
        
        If chkGST.Value = 1 Then
            lblTotal(0).Caption = "Total: " + Format(IIf(IsNumeric(txtSum + Chr(KeyAscii)) = True, Val(txtSum + Chr(KeyAscii)) * oTax(Login.TaxCode, Login.TaxCountry), 0) + Val(txtSum + Chr(KeyAscii)) + lvPaymentType.Tag, "Currency")
            'lblTotal(1).Caption = "Round: $ " & Format(Round(IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax("G","AUS0001"), 0) + Val(txtSum), 2), "###,###,###,###,###.##")
        Else
            lblTotal(0).Caption = "Total: " & Format(Val(txtSum + Chr(KeyAscii)) + lvPaymentType.Tag, "Currency")
            
            'lblTotal(1).Caption = "Round: $ " & Format(Round(IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax("G","AUS0001"), 0), 2), "###,###,###,###,###.##")
        End If
        
    Case Asc(".")
        If InStr(txtSum, ".") > 0 Then KeyAscii = 0
    Case 13
        KeyAscii = 0
        SendKeys "{TAB}"
    Case Else
        KeyAscii = 0
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



Public Function LoadInformation()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "LoadInformation"
    Const ContainerName = "frmPrepayment"
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
    
    Dim rsload As adodb.Recordset
    Dim itmX As ListItem
    
    bResult = MySQL.OpenTable(directConn, rsload, , "select accountinfo.AccountName, invoicein.* from invoicein, accountinfo Where accountinfo.RecID = AccI_RecID and invoicein.RecID = " & lRecID & " Limit 1")
    Set itmX = lvAccount.ListItems.Add(, , rsload!AccountName)
    itmX.Checked = True
    lvAccount.Enabled = False
    
    LoadFlags lRecID
    
    lvPaymentType.Enabled = False
    txtSum.Locked = True
    txtSum = "" & rsload!AmountPaid
    
    If rsload!GSTCharged <> 0 Then
        chkGST.Caption = "GST: " & Format(rsload!GSTCharged, "Currency")
        chkGST.Value = 1
        'lblTotal(1).Caption = "Round: $ " & Format(Round(IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax("G","AUS0001"), 0) + Val(txtSum), 2), "###,###,###,###,###.##")
    Else
        chkGST.Caption = "GST: " & Format(rsload!GSTCharged, "Currency")
        chkGST.Value = 0
        'lblTotal(1).Caption = "Round: $ " & Format(Round(IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax("G","AUS0001"), 0), 2), "###,###,###,###,###.##")
    End If
    
    chkGST.Enabled = False
    cmdMakePayment.Enabled = False
    cmdCancel.Caption = "&Close"
    
    lblTotal(0).Caption = "Total: " + Format(Val(txtSum) + lvPaymentType.Tag + rsload!GSTCharged, "Currency")
    lblTotal(1).Caption = "Sub: " & Format(rsload!sub, "Currency")
    
    
Exit Function



ErrorOccur:
Select Case oErr.chkError(directConn, Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function

Private Sub txtSum_LostFocus()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtSum_LostFocus"
    Const ContainerName = "frmPrepayment"
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


    
    chkGST.Caption = "GST: " & Format(Round(IIf(IsNumeric(txtSum + Chr(KeyAscii)) = True, Val(txtSum + Chr(KeyAscii)) * oTax(Login.TaxCode, Login.TaxCountry), 0), 2), "Currency")
    chkGST.Tag = IIf(IsNumeric(txtSum + Chr(KeyAscii)) = True, Val(txtSum + Chr(KeyAscii)) * oTax(Login.TaxCode, Login.TaxCountry), 0)
    
    If chkGST.Value = 1 Then
        lblTotal(0).Caption = "Total: " + Format(IIf(IsNumeric(txtSum + Chr(KeyAscii)) = True, Val(txtSum + Chr(KeyAscii)) * oTax(Login.TaxCode, Login.TaxCountry), 0) + Val(txtSum + Chr(KeyAscii)) + lvPaymentType.Tag, "Currency")
        'lblTotal(1).Caption = "Round: $ " & Format(Round(IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax("G","AUS0001"), 0) + Val(txtSum), 2), "###,###,###,###,###.##")
    Else
        lblTotal(0).Caption = "Total: " & Format(Val(txtSum + Chr(KeyAscii)) + lvPaymentType.Tag, "Currency")
        
        'lblTotal(1).Caption = "Round: $ " & Format(Round(IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax("G","AUS0001"), 0), 2), "###,###,###,###,###.##")
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
