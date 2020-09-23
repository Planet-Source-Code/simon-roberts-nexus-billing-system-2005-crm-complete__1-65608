VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmEXPa 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Single Expense #"
   ClientHeight    =   6735
   ClientLeft      =   2175
   ClientTop       =   3885
   ClientWidth     =   13065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   13065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      Caption         =   "Payments"
      Height          =   4275
      Left            =   5760
      TabIndex        =   31
      Top             =   2340
      Width           =   7185
      Begin VB.CommandButton cmdPayment 
         Caption         =   "Make a &Payment"
         Enabled         =   0   'False
         Height          =   315
         Left            =   5070
         TabIndex        =   33
         Top             =   3870
         Width           =   1935
      End
      Begin MSComctlLib.ListView lvPayments 
         Height          =   3555
         Left            =   180
         TabIndex        =   32
         Top             =   270
         Width           =   6825
         _ExtentX        =   12039
         _ExtentY        =   6271
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "!AmountPaid!^Currency^"
            Text            =   "Amount Paid"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "!RemittanceSent!^YesNo^"
            Text            =   "Remittance Sent"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "PaymentMethod^select Description as nResult from paymenttype where RecID = "
            Text            =   "Payment Method"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "!SerialNo!"
            Text            =   "Serial No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "!PaidTo!"
            Text            =   "Paid To"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Expense Header"
      Height          =   2175
      Left            =   120
      TabIndex        =   21
      Top             =   60
      Width           =   12825
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save Expense"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   10590
         TabIndex        =   37
         Top             =   270
         Width           =   2085
      End
      Begin VB.Frame Frame6 
         Caption         =   "Total Paid (Inc Tax)"
         Height          =   1065
         Index           =   1
         Left            =   10590
         TabIndex        =   35
         Top             =   990
         Width           =   2115
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Index           =   9
            Left            =   150
            Locked          =   -1  'True
            TabIndex        =   36
            Tag             =   "!AmountPaid!^Currency^"
            Text            =   "$ 0.00"
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Total Owing (Inc tax)"
         Height          =   1065
         Index           =   0
         Left            =   8400
         TabIndex        =   34
         Top             =   990
         Width           =   2115
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Index           =   8
            Left            =   150
            TabIndex        =   3
            Tag             =   "!AmountDue!^Currency^"
            Text            =   "$ 0.00"
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Maintained Last By, Last Payment Entered By"
         Height          =   1065
         Index           =   1
         Left            =   4470
         TabIndex        =   26
         Top             =   990
         Width           =   3855
         Begin VB.ComboBox cmbField 
            Height          =   315
            Index           =   3
            Left            =   1230
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Tag             =   "AssSysopID"
            Top             =   240
            Width           =   2475
         End
         Begin VB.ComboBox cmbField 
            Height          =   315
            Index           =   2
            Left            =   1230
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Tag             =   "AssVirtualID"
            Top             =   630
            Width           =   2475
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Sysop:"
            Height          =   195
            Index           =   3
            Left            =   270
            TabIndex        =   30
            Top             =   300
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Virtual ISP:"
            Height          =   195
            Index           =   2
            Left            =   270
            TabIndex        =   29
            Top             =   690
            Width           =   780
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Ownership"
         Height          =   1065
         Index           =   0
         Left            =   150
         TabIndex        =   23
         Top             =   990
         Width           =   4245
         Begin VB.ComboBox cmbField 
            Height          =   315
            Index           =   1
            Left            =   1230
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Tag             =   "VirtualID"
            Top             =   630
            Width           =   2895
         End
         Begin VB.ComboBox cmbField 
            Height          =   315
            Index           =   0
            Left            =   1230
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Tag             =   "SysopID"
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Virtual ISP:"
            Height          =   195
            Index           =   1
            Left            =   270
            TabIndex        =   25
            Top             =   690
            Width           =   780
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Sysop:"
            Height          =   195
            Index           =   0
            Left            =   270
            TabIndex        =   24
            Top             =   300
            Width           =   480
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Description"
         Height          =   765
         Left            =   150
         TabIndex        =   22
         Top             =   210
         Width           =   10365
         Begin VB.TextBox txtField 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   7
            Left            =   150
            MaxLength       =   255
            TabIndex        =   0
            Tag             =   "Description"
            Top             =   240
            Width           =   10065
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Biller Details"
      Height          =   4275
      Left            =   120
      TabIndex        =   12
      Top             =   2340
      Width           =   5565
      Begin VB.TextBox txtField 
         Height          =   315
         Index           =   10
         Left            =   1380
         MaxLength       =   255
         TabIndex        =   7
         Tag             =   "dFaxNumber"
         Top             =   1380
         Width           =   3975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Search for &Biller"
         Height          =   315
         Left            =   3420
         TabIndex        =   20
         ToolTipText     =   "Fill in the search parameters in the fields above in the Biller Details frame. Then select this button to search for a biller."
         Top             =   3870
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtField 
         Height          =   315
         Index           =   6
         Left            =   1380
         MaxLength       =   255
         TabIndex        =   11
         Tag             =   "dInvoiceNo"
         Top             =   3480
         Width           =   3975
      End
      Begin VB.TextBox txtField 
         Height          =   315
         Index           =   5
         Left            =   1380
         MaxLength       =   255
         TabIndex        =   10
         Tag             =   "dAccNo"
         Top             =   3090
         Width           =   3975
      End
      Begin VB.TextBox txtField 
         Height          =   315
         Index           =   4
         Left            =   1380
         MaxLength       =   255
         TabIndex        =   9
         Tag             =   "dEmailAddress"
         Top             =   2700
         Width           =   3975
      End
      Begin VB.TextBox txtField 
         Height          =   855
         Index           =   3
         Left            =   1380
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   8
         Tag             =   "dAddress"
         Top             =   1770
         Width           =   3975
      End
      Begin VB.TextBox txtField 
         Height          =   315
         Index           =   2
         Left            =   1380
         MaxLength       =   255
         TabIndex        =   6
         Tag             =   "dPhoneNumber"
         Top             =   990
         Width           =   3975
      End
      Begin VB.TextBox txtField 
         Height          =   315
         Index           =   1
         Left            =   1380
         MaxLength       =   255
         TabIndex        =   5
         Tag             =   "dContactname"
         Top             =   660
         Width           =   3975
      End
      Begin VB.TextBox txtField 
         Height          =   315
         Index           =   0
         Left            =   1380
         MaxLength       =   255
         TabIndex        =   4
         Tag             =   "dBodyName"
         Top             =   270
         Width           =   3975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "To Search for a biller fill in the details you want to search for and then click -->"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   300
         TabIndex        =   39
         Top             =   3840
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fax Number:"
         Height          =   195
         Index           =   7
         Left            =   150
         TabIndex        =   38
         Top             =   1440
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Invoice No.:"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   19
         Top             =   3540
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Biller Code:"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   18
         Top             =   3150
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Contact Email:"
         Height          =   195
         Index           =   4
         Left            =   150
         TabIndex        =   17
         Top             =   2760
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Postal Address:"
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   16
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Phone Number:"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   15
         Top             =   1050
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Contact Name:"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   14
         Top             =   690
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Body Name:"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   13
         Top             =   330
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmEXPa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RecID As Double
Public lv As ListView
Public CategoryID As Double

Dim Created As Boolean


Public Function LoadVISPs()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "LoadVISPs"
    Const ContainerName = "frmCustomerRec"
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



    Dim SQL As String
    
    If ViSPMAP.Count > 0 Then
        Dim iCnt As Long
        For iCnt = 1 To ViSPMAP.Count
        
            SQL = SQL + "'" & ViSPMAP(iCnt).RecIDb & "',"
            
        Next
        SQL = Left(SQL, Len(SQL) - 1)
        
        Dim rsload As adodb.Recordset
        
        Call MySQL.OpenTable(directConn, rsload, , "select RecID IN (" & SQL & "), RecID, Description from virtualisp")
        cmbField(1).Clear
        cmbField(2).Clear
        If rsload.State = adStateOpen Then
            If rsload.RecordCount > 0 Then
                Do
                    
                    cmbField(1).AddItem IIf(IsNull(rsload!Description), "(null)", rsload!Description)
                    cmbField(1).ItemData(cmbField(1).ListCount - 1) = rsload!RecID
                    cmbField(2).AddItem IIf(IsNull(rsload!Description), "(null)", rsload!Description)
                    cmbField(2).ItemData(cmbField(2).ListCount - 1) = rsload!RecID
                    
                    If rsload!RecID = Login.lVirtualID Then cmbField(1).ListIndex = cmbField(1).ListCount - 1
                    If rsload!RecID = Login.lVirtualID Then cmbField(2).ListIndex = cmbField(2).ListCount - 1
                    
                    rsload.MoveNext
                    
                Loop Until rsload.EOF Or Err.Number <> 0
            End If
        End If
    Else
    
        cmbField(1).AddItem Login.sVISPName
        cmbField(1).ItemData(cmbField(1).ListCount - 1) = Login.lVirtualID
        cmbField(1).ListIndex = 0
        cmbField(2).AddItem Login.sVISPName
        cmbField(2).ItemData(cmbField(2).ListCount - 1) = Login.lVirtualID
        cmbField(2).ListIndex = 0
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

Public Sub LoadSysops()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "LoadSysops"
    Const ContainerName = "frmCustomerRec"
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



    Dim SQL As String
    
    If ViSPMAP.Count > 0 Then
        Dim iCnt As Long
        For iCnt = 1 To ViSPMAP.Count
        
            SQL = SQL + "VirtualID = '" & ViSPMAP(iCnt).RecIDb & "' or "
            
        Next
        SQL = Left(SQL, Len(SQL) - 4)
        
        Dim rsload As adodb.Recordset
        
        Call MySQL.OpenTable(directConn, rsload, , "select RecID, Username, Firstname, Surname from sysops where " & SQL)
        
        If rsload.State = adStateOpen Then
            If rsload.RecordCount > 0 Then
                Do
                    cmbField(0).AddItem IIf(IsNull(rsload!Username), "[(null)] - ", "[" & rsload!Username & "] - ") & IIf(IsNull(rsload!Firstname), "", rsload!Firstname) & " " & IIf(IsNull(rsload!Surname), "", rsload!Surname)
                    cmbField(0).ItemData(cmbField(0).ListCount - 1) = rsload!RecID
                    
                    cmbField(3).AddItem IIf(IsNull(rsload!Username), "[(null)] - ", "[" & rsload!Username & "] - ") & IIf(IsNull(rsload!Firstname), "", rsload!Firstname) & " " & IIf(IsNull(rsload!Surname), "", rsload!Surname)
                    cmbField(3).ItemData(cmbField(3).ListCount - 1) = rsload!RecID
                    
                    If rsload!RecID = Login.lSysopID Then cmbField(0).ListIndex = cmbField(0).ListCount - 1
                    If rsload!RecID = Login.lSysopID Then cmbField(3).ListIndex = cmbField(3).ListCount - 1
                    
                    rsload.MoveNext
                Loop Until rsload.EOF Or Err.Number <> 0
            End If
        End If
    Else
    
        cmbField(0).AddItem "Sysop No: " & Login.lSysopID
        cmbField(0).ItemData(cmbField(0).ListCount - 1) = Login.lSysopID
        cmbField(0).ListIndex = 0
        cmbField(3).AddItem "Sysop No: " & Login.lSysopID
        cmbField(3).ItemData(cmbField(3).ListCount - 1) = Login.lSysopID
        cmbField(3).ListIndex = 0
        
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



Private Sub cmdPayment_Click()

    Dim fPay As New frmEXPpay
    
    fPay.AmountLeft = CCur(txtField(8).Text) - CCur(txtField(9).Text)
    fPay.InBookID = Me.RecID
    fPay.Show 1
    
    If fPay.Cancel = True Then Exit Sub
    
    Dim rsload As adodb.Recordset
    
    Call MySQL.OpenTable(directConn, rsload, , "select * from exp_flags where RecID = '" & fPay.RecID & "'")
                
    If rsload.State = adStateOpen Then
        If rsload.RecordCount > 0 Then
            
            Call MySQL.fillLV(directConn, rsload, lvPayments, True)
                            
        End If
    End If
    
    Call MySQL.Execute(directConn, "update exp_inbook set AmountPaid = AmountPaid + " & fPay.AmountPaid & " where RecID = '" & Me.RecID & "'")
    txtField(9).Text = Format(CCur(txtField(9).Text) + fPay.AmountPaid, "Currency")
    
End Sub

Private Sub cmdSave_Click()

    Dim bf As Byte
    
    For bf = txtField.LBound To txtField.UBound
        If Trim(txtField(bf).Text) = "" Then
            txtField(bf).SetFocus
            MsgBox "All Fields must be completed to create an expense in your profile."
            Exit Sub
            Exit For
        End If
    Next bf

    If Me.RecID = 0 Then
        
        On Error Resume Next
        Do
            Err.Clear
            Me.RecID = MySQL.GetTMPRecID("exp_inbook", directConn)
            MySQL.Execute directConn, "Insert into exp_inbook (RecID) Values('" & Me.RecID & "')"
        Loop Until Err.Number = 0
        
        For bf = txtField.LBound To txtField.UBound
            MySQL.Execute directConn, "update exp_inbook set `" & MySQL.fldConst(txtField(bf).Tag, 0) & "` = '" & MySQL.ESC(txtField(bf).Text) & "' where RecID = '" & Me.RecID & "'"
        Next bf
        
        For bf = cmbField.LBound To cmbField.UBound
            If cmbField(bf).ListIndex = -1 Then
                Select Case cmbField(bf).Tag
                Case "SysopID", "AssSysopID"
                    MySQL.Execute directConn, "update exp_inbook set `" & cmbField(bf).Tag & "` = '" & Login.lSysopID & "' where RecID = '" & Me.RecID & "'"
                Case "VirtualID", "AssVirtualID"
                    MySQL.Execute directConn, "update exp_inbook set `" & cmbField(bf).Tag & "` = '" & Login.lVirtualID & "' where RecID = '" & Me.RecID & "'"
                End Select
            Else
                MySQL.Execute directConn, "update exp_inbook set `" & cmbField(bf).Tag & "` = '" & cmbField(bf).ItemData(cmbField(bf).ListIndex) & "' where RecID = '" & Me.RecID & "'"
            End If
        Next bf
        
        
        MySQL.Execute directConn, "update exp_inbook set `CategoryID` = '" & Me.CategoryID & "' where RecID = '" & Me.RecID & "'"
        MySQL.Execute directConn, "update exp_inbook set `GST` = '" & (Val(txtField(8)) * (11 * oTax(Login.TaxCode, Login.TaxCountry))) - Val(txtField(8)) & "' where RecID = '" & Me.RecID & "'"
        Created = True
        
        Me.Caption = "Single Expense - [" & Me.RecID & "]"
        
        cmdPayment.Enabled = True
    End If
    
                
                
End Sub

Private Sub Form_Load()

    LoadVISPs
    LoadSysops
        
    Me.Caption = "Single Expense - [" & Me.RecID & "]"
    
    Dim rsload As adodb.Recordset
    
    If RecID <> 0 Then
        cmdSave.Enabled = False
        cmdPayment.Enabled = True
        Call MySQL.OpenTable(directConn, rsload, , "select * from exp_inbook where RecID = " & Me.RecID)
        If rsload.State = adStateOpen Then
            If rsload.RecordCount > 0 Then
                Dim bf As Byte
            
                For bf = txtField.LBound To txtField.UBound
                    txtField(bf).Text = MySQL.OP(rsload, IIf(InStr(txtField(bf).Tag, "!") > 0, txtField(bf).Tag, "!" + txtField(bf).Tag + "!"))
                    txtField(bf).Locked = True
                Next bf
                
                Dim il As Integer
                
                For bf = cmbField.LBound To cmbField.UBound
                    For il = 0 To cmbField(bf).ListCount - 1
                    
                        If rsload(txtField(bf).Tag) = cmbField(bf).ItemData(il) Then
                            cmbField(bf).ListIndex = il
                            Exit For
                        End If
                        
                    Next
                Next bf
                
                Call MySQL.OpenTable(directConn, rsload, , "select * from exp_flags where InBookID = " & Me.RecID)
                
                If rsload.State = adStateOpen Then
                    If rsload.RecordCount > 0 Then
                        
                        Call MySQL.fillLV(directConn, rsload, lvPayments)
                                        
                    End If
                End If
            End If
        End If
    End If
    
End Sub

Private Sub txtField_GotFocus(Index As Integer)
    
    txtField(Index).SelStart = 0
    txtField(Index).SelLength = Len(txtField(Index).Text)
    
End Sub

Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)

    Select Case Index
    Case 8
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
        
        ElseIf KeyAscii = 8 Then
        
        ElseIf KeyAscii = Asc(".") Then
        
            If InStr(txtField(Index).Text, ".") > 0 Then KeyAscii = 0
        
        Else
            KeyAscii = 0
        End If
        
    End Select
End Sub
