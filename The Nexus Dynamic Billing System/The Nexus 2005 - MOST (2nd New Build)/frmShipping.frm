VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmShipping 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Double click from the list view to select shipping address"
   ClientHeight    =   5715
   ClientLeft      =   1650
   ClientTop       =   1755
   ClientWidth     =   9780
   ControlBox      =   0   'False
   Icon            =   "frmShipping.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picTSContainer 
      BorderStyle     =   0  'None
      Height          =   5595
      Index           =   0
      Left            =   60
      ScaleHeight     =   5595
      ScaleWidth      =   9645
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   9645
      Begin VB.CommandButton cmdAddAddress 
         Caption         =   "&Add Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   90
         TabIndex        =   1
         Top             =   5160
         Width           =   1905
      End
      Begin MSComctlLib.ListView lvAddresses 
         Height          =   4965
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   9465
         _ExtentX        =   16695
         _ExtentY        =   8758
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HoverSelection  =   -1  'True
         PictureAlignment=   1
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
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
            Text            =   "Contact Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Street 1"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Street 2"
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Suburb"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "State"
            Object.Width           =   1589
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Postcode"
            Object.Width           =   1588
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Country"
            Object.Width           =   2540
         EndProperty
         Picture         =   "frmShipping.frx":030A
      End
   End
End
Attribute VB_Name = "frmShipping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lRecID As Long
Public ShippingID As Long
Public LinkdForm As frmCustomerRec
Public ContactName As String
Public Street1 As String
Public Street2 As String
Public Suburb As String
Public State As String
Public Postcode As String
Public Country As String



Private Sub cmdAddAddress_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdAddAddress_Click"
    Const ContainerName = "frmShipping"
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


    Dim itmX As ListItem
    Dim ffrmSnailMail As frmSnailMail
    Set ffrmSnailMail = New frmSnailMail
    
    
    ffrmSnailMail.sContactName = LinkdForm.txtAccountName.Text
    ffrmSnailMail.Show 1
    
    If ffrmSnailMail.iCloseState = frmCloseSave Then
        
        Dim ID As Long
        
        With LinkdForm.osub.colSnailMail.Add("NEW" & LinkdForm.osub.colSnailMail.Count + 1, LinkdForm.osub.fRecID, ffrmSnailMail.FlagID, _
             sysNOW, ffrmSnailMail.sContactName, ffrmSnailMail.sStreetLine1, ffrmSnailMail.sStreetLine2, ffrmSnailMail.sCountry, ffrmSnailMail.sState, _
             ffrmSnailMail.sPostcode, ffrmSnailMail.sSuburb, False, True, "NEW" & LinkdForm.osub.colSnailMail.Count + 1)
             
            Set itmX = lvAddresses.ListItems.Add(, .Key, .ContactName)
            itmX.SubItems(1) = .Street1
            itmX.SubItems(2) = .Street2
            itmX.SubItems(3) = .Suburb
            itmX.SubItems(4) = .State
            itmX.SubItems(5) = .Postcode
            itmX.SubItems(6) = .Country
                    
        
            Set itmX = LinkdForm.lvAddresses.ListItems.Add(, .Key, .ContactName)
            itmX.SubItems(1) = .Street1
            itmX.SubItems(2) = .Street2
            itmX.SubItems(3) = .Suburb
            itmX.SubItems(4) = .State
            itmX.SubItems(5) = .Postcode
            itmX.SubItems(6) = .Country
        
        End With
        
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
    Const ContainerName = "frmShipping"
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

    Dim oSNL As clsSnailMail
    
    If LinkdForm.osub.colSnailMail.Count > 0 Then
        For Each oSNL In LinkdForm.osub.colSnailMail
            Set itmX = lvAddresses.ListItems.Add(1, oSNL.Key, oSNL.ContactName) '
            itmX.SubItems(1) = oSNL.Street1
            itmX.SubItems(2) = oSNL.Street2
            itmX.SubItems(3) = oSNL.Suburb
            itmX.SubItems(4) = oSNL.State
            itmX.SubItems(5) = oSNL.Postcode
            itmX.SubItems(6) = oSNL.Country
            itmX.Checked = oSNL.Checked
        Next
    End If

    
    If bBigFont = True Then lvAddresses.Font.Size = 18
    
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

Private Sub lvAddresses_DblClick()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvAddresses_DblClick"
    Const ContainerName = "frmShipping"
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


    If lvAddresses.SelectedItem Is Nothing Then
    
    Else
        Dim oSNL As clsSnailMail
        
        For Each oSNL In LinkdForm.osub.colSnailMail
            If lvAddresses.SelectedItem.Key = oSNL.Key Then
                Me.ShippingID = oSNL.IDX
                Exit For
            End If
        Next
        
        Unload Me
        
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
