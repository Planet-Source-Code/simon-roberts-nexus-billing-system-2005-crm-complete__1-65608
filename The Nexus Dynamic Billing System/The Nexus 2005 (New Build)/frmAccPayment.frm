VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAccPayment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Payment"
   ClientHeight    =   6480
   ClientLeft      =   1530
   ClientTop       =   2205
   ClientWidth     =   7845
   Icon            =   "frmAccPayment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdMakePayment 
      Caption         =   "&Make Payment"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3990
      TabIndex        =   18
      Top             =   5850
      Width           =   3765
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   90
      TabIndex        =   17
      Top             =   5850
      Width           =   3855
   End
   Begin VB.Frame Frame4 
      Height          =   2265
      Left            =   3990
      TabIndex        =   8
      Top             =   3480
      Width           =   3765
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
         TabIndex        =   10
         Top             =   360
         Width           =   1995
      End
      Begin VB.CheckBox chkGST 
         Caption         =   "&GST: $0.00"
         Enabled         =   0   'False
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
         TabIndex        =   9
         Top             =   810
         Visible         =   0   'False
         Width           =   3345
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
         Index           =   3
         Left            =   525
         TabIndex        =   13
         Top             =   360
         Width           =   990
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   0
         X1              =   1350
         X2              =   3630
         Y1              =   1170
         Y2              =   1170
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
         TabIndex        =   12
         Top             =   1650
         Width           =   1320
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   1
         X1              =   1350
         X2              =   3660
         Y1              =   2040
         Y2              =   2040
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
         TabIndex        =   11
         Top             =   1260
         Width           =   3120
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Payment Type"
      Height          =   2265
      Left            =   90
      TabIndex        =   6
      Top             =   3480
      Width           =   3855
      Begin MSComctlLib.ListView lvPaymentType 
         Height          =   1875
         Left            =   120
         TabIndex        =   7
         Tag             =   "0"
         Top             =   270
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   3307
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
   Begin VB.Frame Frame1 
      Caption         =   "Payments Details"
      Height          =   3345
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   7665
      Begin VB.Frame Frame2 
         Caption         =   "Payments Made"
         Height          =   2055
         Left            =   150
         TabIndex        =   4
         Top             =   1170
         Width           =   7395
         Begin MSComctlLib.ListView lvPaymentMade 
            Height          =   1605
            Left            =   120
            TabIndex        =   5
            Top             =   300
            Width           =   7125
            _ExtentX        =   12568
            _ExtentY        =   2831
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
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "When"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Amount"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "GST"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Subcharge"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Label lblPaid 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1770
         TabIndex        =   16
         Top             =   840
         Width           =   5760
      End
      Begin VB.Label lblDue 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1770
         TabIndex        =   15
         Top             =   600
         Width           =   5760
      End
      Begin VB.Label lblName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1770
         TabIndex        =   14
         Top             =   360
         Width           =   5760
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total Amount Paid:"
         Height          =   240
         Index           =   2
         Left            =   180
         TabIndex        =   3
         Top             =   870
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total Amount Due:"
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   600
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Account Name:"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   330
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmAccPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public iCloseState As frm_CloseStates

Public l_RecID As Variant

Public acci_RecID As Variant

Public s_AccountName As String
Public c_TotalDue As Currency
Public c_TotalPaid As Currency
Public c_Description As String

Public Function LoadPayments(Optional ltmpRecID As Variant)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "LoadPayments"
    Const ContainerName = "frmAccPayment"
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

    
    If bDebug = True Then On Error GoTo 0 Else On Error GoTo ErrorOccur
    
    Dim rsLoad As ADODB.Recordset
    Dim itmX As ListItem
    Dim bResult As Boolean
    
    bResult = MySQL.OpenTable(ADOConn, rsLoad, , "select * from invout_payment where InvOut_RecID = " & ltmpRecID)
    
    If rsLoad.RecordCount > 0 Then
        rsLoad.MoveFirst
        While Not rsLoad.EOF And Err.Number = 0
            Set itmX = lvPaymentMade.ListItems.Add(, "r" & rsLoad!RecID, Format(rsLoad!WhenPaid, "dd-mm-yyyy Hh:Nn:Ss"))
            itmX.SubItems(1) = IIf(rsLoad!Amount = 0, "", Format(rsLoad!Amount, "Currency"))
            itmX.SubItems(2) = IIf(rsLoad!GST = 0, "", Format(rsLoad!GST, "Currency"))
            itmX.SubItems(3) = IIf(rsLoad!Sub = 0, "", Format(rsLoad!Sub, "Currency"))
            rsLoad.MoveNext
        Wend
    End If

    If Err.Number = 0 Then Exit Function
    

Exit Function



ErrorOccur:
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function

Public Function LoadFlags(Optional ltmpRecID As Variant)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "LoadFlags"
    Const ContainerName = "frmAccPayment"
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


    If bDebug = True Then On Error GoTo 0 Else On Error GoTo ErrorOccur
    
    Dim rsLoad As ADODB.Recordset
    Dim itmX As ListItem
    Dim bResult As Boolean
    
    bResult = MySQL.OpenTable(ADOConn, rsLoad, , "select * from paymenttype")
    
    If rsLoad.RecordCount > 0 Then
        rsLoad.MoveFirst
        While Not rsLoad.EOF And Err.Number = 0
            Set itmX = lvPaymentType.ListItems.Add(, IIf(rsLoad!CreditCard = True, "c", "a") & rsLoad!RecID, rsLoad!Description)
            itmX.SubItems(1) = IIf(rsLoad!Sub = 0, "", Format(rsLoad!Sub, "Currency"))
            itmX.Tag = rsLoad!Sub
            rsLoad.MoveNext
        Wend
    End If

    Dim ix As Integer
    
    'If ltmpRecID <> 0 Then
    '    bResult = MySQL.OpenTable(adoconn, rsLoad, , "select * from Flags_invoicein Where InvIn_RecID = " & ltmpRecID)
    '
    '    If rsLoad.RecordCount > 0 Then
    '        While Not rsLoad.EOF
    '            For iX = 1 To lvpaymenttype.ListItems.Count
    '                If Val(Mid(lvpaymenttype.ListItems(iX).Key, 2)) = rsLoad!Flag Then
    '                    lvpaymenttype.ListItems(iX).Checked = True
    '                    Exit For
    '                End If
    '            Next
    '            rsLoad.MoveNext
    '        Wend
    '
    '    End If
    'End If

    If Err.Number = 0 Then Exit Function
    

Exit Function



ErrorOccur:
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function

Private Sub chkGST_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "chkGST_Click"
    Const ContainerName = "frmAccPayment"
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


    If chkGST.Value = 1 Then
        lblTotal(0).Caption = "Total: $ " & IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax(Login.TaxCode, Login.TaxCountry), 0) + Val(txtSum)
        'lblTotal(1).Caption = "Round: $ " & Round(IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax("G","AUS0001"), 0) + Val(txtSum), 2)
    Else
        lblTotal(0).Caption = "Total: $ " & Val(txtSum)
        'lblTotal(1).Caption = "Round: $ " & Round(IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax("G","AUS0001"), 0), 2)
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

Private Sub cmdCancel_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdCancel_Click"
    Const ContainerName = "frmAccPayment"
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


    Unload Me
    
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

Private Sub cmdmakePayment_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdmakePayment_Click"
    Const ContainerName = "frmAccPayment"
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


    If bDebug = True Then On Error GoTo 0 Else On Error GoTo ErrorOccur
    
    Dim ix As Integer
    Dim bselected As Boolean
    Dim bSaveCC As Boolean
    Dim bResult As Boolean
    Dim rs_creditcard As ADODB.Recordset
    Dim rs_cc_Receipt As ADODB.Recordset
    
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
                bResult = MySQL.OpenTable(ADOConn, rs_creditcard, , "select RecID, AccI_RecID as acci_RecID, bType, AES_DECRYPT(CardNumber,'" + odb.colSalts.ReturnSalt("CCSalt") + "" & odb.colSalts.ReturnSalt("md5Password") & "') as CardNumber, SecurityNumber, ExpiryDate, Name from creditcard where AccI_RecID = " & Me.acci_RecID)
                bResult = MySQL.OpenTable(ADOConn, rs_cc_Receipt, , "select * from cc_Receipt limit 1")
                                                
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
       
    If IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax(Login.TaxCode, Login.TaxCountry), 0) = 0 Then
        MsgBox "Zero value specified in sum of transaction field!", vbCritical, "Zero Value"
    Else
        If Val(txtSum) > CCur(Mid(lblDue.Caption, 2)) - CCur(Mid(lblPaid.Caption, 2)) Then
            MsgBox "Value is over the maximum amount payable for this transaction!", vbInformation, "Value specified to great"
            txtSum.Text = "" & (CCur(Mid(lblDue.Caption, 2)) - CCur(Mid(lblPaid.Caption, 2)))
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
    
    Dim rsSave As ADODB.Recordset
    Dim rsLoad As ADODB.Recordset
    Dim rsSave2 As ADODB.Recordset
    
    Dim lacci_RecID As Variant
    
    bResult = MySQL.OpenTable(ADOConn, rsLoad, , "select * from invoiceout Where RecID = " & Me.l_RecID & " Limit 1")
    bResult = MySQL.OpenTable(ADOConn, rsSave, , "select * from invout_payment Limit 1")
    
    rsSave.AddNew
    rsSave!InvOut_RecID = Me.l_RecID
    rsSave!Amount = Val(txtSum)
    rsSave!Sub = lvPaymentType.Tag
    rsSave!WhenPaid = sysNOW
    
    
    If chkGST.Value = 1 Then
        rsSave!GST = IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax(Login.TaxCode, Login.TaxCountry), 0)
        'lblTotal(1).Caption = "Round: $ " & Format(Round(IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax("G","AUS0001"), 0) + Val(txtSum), 2), "###,###,###,###,###.##")
    Else
        rsSave!GST = 0
        'lblTotal(1).Caption = "Round: $ " & Format(Round(IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax("G","AUS0001"), 0), 2), "###,###,###,###,###.##")
    End If
    
    rsSave!TotalPaid = Val(txtSum) + lvPaymentType.Tag + rsSave!GST
    rsSave!acci_RecID = rsLoad!acci_RecID
    rsSave.Update

    For ix = 1 To lvPaymentType.ListItems.Count
        If lvPaymentType.ListItems(ix).Checked = True Then
            If Left(lvPaymentType.ListItems(ix).Key, 1) = "c" Then
                MySQL.AddReceiptItem ADOConn, rsLoad!acci_RecID, rsSave!RecID, , , , rsSave!TotalPaid, , "CC: " + Right(fFrmCC.screditcard, 5)
            Else
                MySQL.AddReceiptItem ADOConn, rsLoad!acci_RecID, rsSave!RecID, , , , rsSave!TotalPaid, , lvPaymentType.ListItems(ix).Text + " Payment"
            End If
        End If
    Next
    
    'If rsLoad!AmountPaid + Val(txtSum) + rsSave!GST > rsLoad!TotalDue Then
                
        'bResult = MySQL.OpenTable(adoconn, rsSave2, , "select * from invoicein Limit 1")
        'rsSave2.AddNew
        'rsSave2!AccI_RecID = rsLoad!AccI_RecID
        'rsSave2!AmountPaid = rsLoad!AmountPaid + Val(txtSum) - rsLoad!AmountDue
        'rsSave2!GSTCharged = rsSave2!AmountPaid * oTax("G","AUS0001")
        'rsSave2!AmountUsed = 0
        'rsSave2!Sub = 0
        'rsSave2!GSTCharged = 0
        'rsSave2!TotalPaid = rsSave2!AmountPaid + rsSave2!GSTCharged
        'rsSave2.Update
        
        'rsLoad!AmountPaid = rsLoad!TotalDue
        
    'Else
    
        rsLoad!AmountPaid = rsLoad!AmountPaid + Val(txtSum) + rsSave!GST
    'End If
    
    c_TotalPaid = rsLoad!AmountPaid
    
    rsLoad!PaidWhen = sysNOW
    
    If rsLoad!TraxrID <> 0 Then
        Dim rsTraxr As ADODB.Recordset
        bResult = MySQL.OpenTable(ADOConn, rsTraxr, , "select * from invoicetraxr Where RecID = " & rsLoad!TraxrID)
        If rsTraxr.RecordCount > 0 Then
            rsTraxr!AmountPaid = rsTraxr!AmountPaid + Val(txtSum) + rsSave!GST
            If rsTraxr!AmountPaid >= rsTraxr!TotalDue Then rsTraxr!Finalised = True
            rsTraxr.Update
        End If
    End If
    
    rsLoad.Update
    
    SaveFlags Me.l_RecID
    
    '¤
    If bSaveCC = True Then
        rs_creditcard.Filter = "AccI_RecID = " & rsSave!acci_RecID & " AND CardNumber = '" & MySQL.NumCrypt(fFrmCC.screditcard) & "'"
        If rs_creditcard.RecordCount = -1 Or rs_creditcard.RecordCount = 0 Then
            
            rs_creditcard.AddNew
            rs_creditcard!acci_RecID = rsSave!acci_RecID
            rs_creditcard!SecurityNumber = fFrmCC.sSecurityNo
            rs_creditcard!ExpiryDate = fFrmCC.sExpiry
            rs_creditcard!Name = fFrmCC.sCardName
            rs_creditcard!bType = fFrmCC.bType
            rs_creditcard!bDefault = fFrmCC.bDefault
            
            MySQL.Execute ADOConn, "UPDATE creditcard SET CardNumber=AES_ENCRYPT('" & MySQL.NumCrypt(fFrmCC.screditcard) & "','" + odb.colSalts.ReturnSalt("CCSalt") + "" & odb.colSalts.ReturnSalt("md5Password") & "') where RecID = " & MySQL.SetRecID(rs_creditcard, "creditcard", ADOConn)
            If fFrmCC.bDefault = True Then
                MySQL.Execute ADOConn, "UPDATE creditcard SET bDefault=0 where acci_RecID = " & rs_creditcard!acci_RecID
                MySQL.Execute ADOConn, "UPDATE creditcard SET bDefault=-1 where RecID = " & rs_creditcard!RecID
            End If
            
            rs_cc_Receipt.AddNew
            rs_cc_Receipt!cc_RecID = rs_creditcard!RecID
            rs_cc_Receipt!ReceiptNumber = fFrmCC.sReceiptNo
            rs_cc_Receipt.Update
            
            rs_creditcard.Requery
            
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
    
    If Err.Number = 0 Then Exit Sub
    
    
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


Sub SaveFlags(lRecID As Variant)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "SaveFlags"
    Const ContainerName = "frmAccPayment"
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


    If bDebug = True Then On Error GoTo 0 Else On Error GoTo ErrorOccur
    
    Dim rsSave As ADODB.Recordset
    Dim bResult As Boolean
    Dim ix As Long
    
    bResult = MySQL.OpenTable(ADOConn, rsSave, , "select * from Flags_invoiceout Limit 0")
        
    For ix = 1 To lvPaymentType.ListItems.Count
        If lvPaymentType.ListItems(ix).Checked = True Then
            rsSave.AddNew
            rsSave!InvOut_RecID = lRecID
            rsSave!Flag = Mid(lvPaymentType.ListItems(ix).Key, 2)
            rsSave.Update
        End If
    Next ix
    
    If Err.Number = 0 Then Exit Sub
    
    
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
    Const ContainerName = "frmAccPayment"
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


    If bBigFont = True Then

        
        lvPaymentMade.Font.Size = 16
        lvPaymentType.Font.Size = 16
        
    End If

    lblName.Caption = Me.s_AccountName
    lblDue.Caption = Format(Me.c_TotalDue, "Currency")
    lblPaid.Caption = Format(Me.c_TotalPaid, "Currency")
    
    txtSum.Text = "" & Me.c_TotalDue - Me.c_TotalPaid
    
    Call GUI.LoadColWidths(lvPaymentMade, Me)
    
    Call LoadFlags
    Call LoadPayments(Me.l_RecID)
        
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_QueryUnload"
    Const ContainerName = "frmAccPayment"
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


    Call GUI.SaveColWidths(lvPaymentMade, Me)
    
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

Private Sub lvpaymenttype_ItemCheck(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvpaymenttype_ItemCheck"
    Const ContainerName = "frmAccPayment"
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


    Select Case Item.Checked
    Case True
        lvPaymentType.Tag = lvPaymentType.Tag + Item.Tag
    Case False
        lvPaymentType.Tag = lvPaymentType.Tag - Item.Tag
    End Select
    
    lblTotal(1).Caption = "Sub: " + Format(lvPaymentType.Tag, "Currency")
    
    If chkGST.Value = 1 Then
        lblTotal(0).Caption = "Total: " + Format(IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax(Login.TaxCode, Login.TaxCountry), 0) + Val(txtSum) + lvPaymentType.Tag, "Currency")
        'lblTotal(1).Caption = "Round: $ " & Format(Round(IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax("G","AUS0001"), 0) + Val(txtSum), 2), "###,###,###,###,###.##")
    Else
        lblTotal(0).Caption = "Total: " & Format(Val(txtSum) + lvPaymentType.Tag, "Currency")
        'lblTotal(1).Caption = "Round: $ " & Format(Round(IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax("G","AUS0001"), 0), 2), "###,###,###,###,###.##")
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

Private Sub txtSum_Change()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtSum_Change"
    Const ContainerName = "frmAccPayment"
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


    If bDebug = True Then On Error GoTo 0 Else On Error GoTo ErrorOccur
    
    chkGST.Caption = "GST: " & Format(Round(IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax(Login.TaxCode, Login.TaxCountry), 0), 2), "Currency")
    chkGST.Tag = IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax(Login.TaxCode, Login.TaxCountry), 0)
    
    If chkGST.Value = 1 Then
        lblTotal(0).Caption = "Total: " + Format(IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax(Login.TaxCode, Login.TaxCountry), 0) + Val(txtSum) + lvPaymentType.Tag, "Currency")
        'lblTotal(1).Caption = "Round: $ " & Format(Round(IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax("G","AUS0001"), 0) + Val(txtSum), 2), "###,###,###,###,###.##")
    Else
        lblTotal(0).Caption = "Total: " & Format(Val(txtSum) + lvPaymentType.Tag, "Currency")
        
        'lblTotal(1).Caption = "Round: $ " & Format(Round(IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax("G","AUS0001"), 0), 2), "###,###,###,###,###.##")
    End If
        
    If Err.Number = 0 Then Exit Sub
    
        
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

Private Sub txtSum_KeyPress(KeyAscii As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtSum_KeyPress"
    Const ContainerName = "frmAccPayment"
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
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Sub

Private Sub txtSum_LostFocus()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtSum_LostFocus"
    Const ContainerName = "frmAccPayment"
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



        chkGST.Caption = "GST: " & Format(Round(IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax(Login.TaxCode, Login.TaxCountry), 0), 2), "Currency")
        chkGST.Tag = IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax(Login.TaxCode, Login.TaxCountry), 0)
        
        If chkGST.Value = 1 Then
            lblTotal(0).Caption = "Total: " + Format(IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax(Login.TaxCode, Login.TaxCountry), 0) + Val(txtSum) + lvPaymentType.Tag, "Currency")
            'lblTotal(1).Caption = "Round: $ " & Format(Round(IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax("G","AUS0001"), 0) + Val(txtSum), 2), "###,###,###,###,###.##")
        Else
            lblTotal(0).Caption = "Total: " & Format(Val(txtSum) + lvPaymentType.Tag, "Currency")
            
            'lblTotal(1).Caption = "Round: $ " & Format(Round(IIf(IsNumeric(txtSum) = True, Val(txtSum) * oTax("G","AUS0001"), 0), 2), "###,###,###,###,###.##")
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

Public Function LoadAccount()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "LoadAccount"
    Const ContainerName = "frmAccPayment"
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



Exit Function



ErrorOccur:
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function
