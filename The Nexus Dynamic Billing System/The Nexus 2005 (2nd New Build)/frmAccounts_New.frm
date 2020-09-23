VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAccounts 
   Caption         =   "Accounts"
   ClientHeight    =   5565
   ClientLeft      =   4275
   ClientTop       =   2520
   ClientWidth     =   8775
   Icon            =   "frmAccounts_New.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5565
   ScaleWidth      =   8775
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   3780
      Top             =   2550
   End
   Begin VB.PictureBox picAccounts 
      BorderStyle     =   0  'None
      Height          =   2835
      Index           =   1
      Left            =   630
      ScaleHeight     =   2835
      ScaleWidth      =   5295
      TabIndex        =   3
      Top             =   1620
      Width           =   5295
      Begin VB.CommandButton cmdExtra 
         Caption         =   "Make Extra Payment"
         CausesValidation=   0   'False
         Height          =   330
         Left            =   90
         TabIndex        =   6
         Top             =   2460
         Width           =   1965
      End
      Begin MSComctlLib.ListView lvExtra 
         Height          =   2355
         Left            =   60
         TabIndex        =   5
         Top             =   60
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   4154
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Account Name"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Total Paid"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Ledger Paid"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Paid When"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "GST Charged"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Amount Used"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Amount Left"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Subcharge"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.PictureBox picAccounts 
      BorderStyle     =   0  'None
      Height          =   2055
      Index           =   0
      Left            =   210
      ScaleHeight     =   2055
      ScaleWidth      =   3375
      TabIndex        =   2
      Top             =   1200
      Width           =   3375
      Begin MSComctlLib.ListView lvReceivables 
         Height          =   4005
         Left            =   60
         TabIndex        =   4
         Top             =   60
         Width           =   8235
         _ExtentX        =   14526
         _ExtentY        =   7064
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Account Name"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Amount Payable"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "GST Charged"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Total Due"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Due When"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Amount Paid"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Amount Refunded"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin MSComctlLib.TabStrip tsAccounts 
      Height          =   4065
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   7170
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      TabStyle        =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Accounts Receivable"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pre-payments"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   165
      Left            =   60
      TabIndex        =   7
      Top             =   5340
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accounts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1080
      TabIndex        =   1
      Top             =   90
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   930
      Index           =   0
      Left            =   90
      Picture         =   "frmAccounts_New.frx":0442
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   930
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   30
      X2              =   8760
      Y1              =   390
      Y2              =   390
   End
End
Attribute VB_Name = "frmAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExtra_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdExtra_Click"
    Const ContainerName = "frmAccounts"
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


    
    If bDebug = True Then On Error GoTo 0 Else On Error GoTo ErrorOccur
    
    Dim ffrmPrepaid As frmPrepayment
    Set ffrmPrepaid = New frmPrepayment
    ffrmPrepaid.Show 1
    
    If ffrmPrepaid.iCloseState = frmCloseSave Then
        Dim itmX As ListItem
                    
        Set itmX = lvExtra.ListItems.Add(, "k" & ffrmPrepaid.lRecID, ffrmPrepaid.sAccountName)
        itmX.SubItems(1) = Format(ffrmPrepaid.cTotal + ffrmPrepaid.cGST + ffrmPrepaid.cSub, "Currency")
        itmX.SubItems(2) = Format(ffrmPrepaid.cTotal, "Currency")
        itmX.SubItems(3) = Format(sysnow, "dd-mm-yyyy Hh:Nn:Ss")
        itmX.SubItems(4) = Format(ffrmPrepaid.cGST, "Currency")
        itmX.SubItems(5) = Format(0, "Currency")
        itmX.SubItems(6) = Format(IIf(IsNull(ffrmPrepaid.cTotal), 0, ffrmPrepaid.cTotal), "Currency")
        itmX.SubItems(5) = Format(ffrmPrepaid.cSub, "Currency")
    End If
    
    If Err.Number = 0 Then Exit Sub
    
    
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

Public Sub Form_Load()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Load"
    Const ContainerName = "frmAccounts"
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

           Set lvReceivables.SmallIcons = fIcon.il16x16
            Set lvReceivables.Icons = fIcon.il32x32
            Set lvExtra.SmallIcons = fIcon.il16x16
            Set lvExtra.Icons = fIcon.il32x32
            
            
    Call tsAccounts_Click
    
    LoadColumnWidths
    
    Me.Visible = True
    gSleep
    'Call loadRS
    
    cmdExtra.Enabled = Login.bMaster
    
    If bBigFont = True Then
        tsAccounts.Font.Size = 14
        lvExtra.Font.Size = 16
        lvRecievables.Font.Size = 16
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_QueryUnload"
    Const ContainerName = "frmAccounts"
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

    
    SaveColumnWidths
    
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

Private Sub Form_Resize()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Resize"
    Const ContainerName = "frmAccounts"
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


If bDebug = True Then On Error GoTo 0 Else On Error GoTo ErrorOccur

    Line1.X2 = Me.ScaleWidth + 500
    Line1.X1 = -10
    
    If Me.ScaleWidth > 2300 And Me.ScaleHeight > 900 Then
        pb1.Move 0, Me.ScaleHeight - tsAccounts.Left - pb1.Height, Me.ScaleWidth
        tsAccounts.Move tsAccounts.Left, tsAccounts.Top, Me.ScaleWidth - (tsAccounts.Left * 2), Me.ScaleHeight - (tsAccounts.Left + tsAccounts.Top + pb1.Height + tsAccounts.Left + tsAccounts.Left)
    
        Call picAccounts_Resize(tsAccounts.SelectedItem.Index - 1)
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

Private Sub lvExtra_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvExtra_ColumnClick"
    Const ContainerName = "frmAccounts"
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


    lvExtra.Sorted = True
    lvExtra.SortKey = ColumnHeader.Index - 1
    If lvExtra.SortOrder = lvwAscending Then
        lvExtra.SortOrder = lvwDescending
    Else
        lvExtra.SortOrder = lvwAscending
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

Private Sub lvExtra_DblClick()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvExtra_DblClick"
    Const ContainerName = "frmAccounts"
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


    If lvExtra.Tag = True Then
        Dim ffrmPrepaid As frmPrepayment
        Set ffrmPrepaid = New frmPrepayment
        
        ffrmPrepaid.lRecID = Val(Mid(lvExtra.SelectedItem.Key, 2))
        ffrmPrepaid.Show 1
        
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

Private Sub lvExtra_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvExtra_ItemClick"
    Const ContainerName = "frmAccounts"
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


    lvExtra.Tag = True
    
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

Private Sub lvReceivables_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvReceivables_ColumnClick"
    Const ContainerName = "frmAccounts"
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


    lvReceivables.Sorted = True
    lvReceivables.SortKey = ColumnHeader.Index - 1
    If lvReceivables.SortOrder = lvwAscending Then
        lvReceivables.SortOrder = lvwDescending
    Else
        lvReceivables.SortOrder = lvwAscending
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

Private Sub lvReceivables_DblClick()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvReceivables_DblClick"
    Const ContainerName = "frmAccounts"
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


    If lvReceivables.Tag <> "" And Login.bMaster = True Then
        Dim frmPayment As frmAccPayment
        Set frmPayment = New frmAccPayment
        
        frmPayment.l_RecID = Val(Mid(lvReceivables.SelectedItem.Key, 2))
        frmPayment.acci_RecID = lvReceivables.SelectedItem.Tag
        frmPayment.s_AccountName = lvReceivables.SelectedItem.Text
        frmPayment.c_TotalDue = CCur(Mid(lvReceivables.SelectedItem.SubItems(3), 2))
        frmPayment.c_TotalPaid = CCur(Mid(lvReceivables.SelectedItem.SubItems(5), 2))
        frmPayment.Show 1
        
        lvReceivables.SelectedItem.SubItems(5) = Format(frmPayment.c_TotalPaid, "Currency")
        
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

Private Sub lvReceivables_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvReceivables_ItemClick"
    Const ContainerName = "frmAccounts"
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


    lvReceivables.Tag = True
    
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

Private Sub picAccounts_Resize(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "picAccounts_Resize"
    Const ContainerName = "frmAccounts"
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


    picAccounts(Index).Move tsAccounts.ClientLeft, tsAccounts.ClientTop, tsAccounts.ClientWidth, tsAccounts.ClientHeight
    
    If tsAccounts.ClientHeight > 100 And tsAccounts.ClientWidth > 2300 Then
    Select Case Index
    Case 0 ' Accounts Receivable
        lvReceivables.Move lvReceivables.Left, lvReceivables.Top, picAccounts(Index).ScaleWidth - (lvReceivables.Left * 2), picAccounts(Index).ScaleHeight - (lvReceivables.Top * 2)
    Case 1 ' Extra Payment
        lvExtra.Move lvExtra.Left, lvExtra.Top, picAccounts(Index).ScaleWidth - (lvExtra.Left * 2), picAccounts(Index).ScaleHeight - (lvExtra.Top * 4) - cmdExtra.Height
        cmdExtra.Top = (lvExtra.Top * 2) + lvExtra.Height
    End Select
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

Private Sub Timer1_Timer()

    Static bSet As Boolean
    
    If Login.IconsSet = True Then
        If bSet = False Then
            
            
            'Set tvServiceTypes.ImageList = fIcon.il16x16
            'Set tvCat.ImageList = fIcon.il16x16
            Set lvReceivables.SmallIcons = fIcon.il16x16
            Set lvReceivables.Icons = fIcon.il32x32
            Set lvExtra.SmallIcons = fIcon.il16x16
            Set lvExtra.Icons = fIcon.il32x32
            
            bSet = True
        End If
    End If
    
    gSleep
    
End Sub

Private Sub tsAccounts_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "tsAccounts_Click"
    Const ContainerName = "frmAccounts"
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


    Dim X As Integer
    
    For X = picAccounts.LBound To picAccounts.UBound
        If tsAccounts.SelectedItem.Index - 1 <> X Then picAccounts(X).Visible = False
    Next
    
    If tsAccounts.SelectedItem.Index - 1 <= picAccounts.UBound Then
        picAccounts(tsAccounts.SelectedItem.Index - 1).Move tsAccounts.ClientLeft, tsAccounts.ClientTop, tsAccounts.ClientWidth, tsAccounts.ClientHeight
        picAccounts(tsAccounts.SelectedItem.Index - 1).Visible = True
        picAccounts(tsAccounts.SelectedItem.Index - 1).ZOrder 0
        Call picAccounts_Resize(tsAccounts.SelectedItem.Index - 1)
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

Public Function SaveColumnWidths()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "SaveColumnWidths"
    Const ContainerName = "frmAccounts"
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
    
    For ix = 1 To lvReceivables.ColumnHeaders.Count
        SaveSetting "projectalpha", "Accounts_frm", "lvReceivables_Width_COL" & ix, lvReceivables.ColumnHeaders(ix).Width
        SaveSetting "projectalpha", "Accounts_frm", "lvReceivables_Text_COL" & ix, lvReceivables.ColumnHeaders(ix).Text
    Next
    
    For ix = 1 To lvExtra.ColumnHeaders.Count
        SaveSetting "projectalpha", "Accounts_frm", "lvExtra_Width_COL" & ix, lvExtra.ColumnHeaders(ix).Width
        SaveSetting "projectalpha", "Accounts_frm", "lvExtra_Text_COL" & ix, lvExtra.ColumnHeaders(ix).Text
    Next
    
        
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

Public Function LoadColumnWidths()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "LoadColumnWidths"
    Const ContainerName = "frmAccounts"
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
    
    For ix = 1 To lvReceivables.ColumnHeaders.Count
        lvReceivables.ColumnHeaders(ix).Width = GetSetting("projectalpha", "Accounts_frm", "lvReceivables_Width_COL" & ix, lvReceivables.ColumnHeaders(ix).Width)
        '¤lvReceivables.ColumnHeaders(iX).Text = GetSetting("projectalpha", "Accounts_frm", "lvReceivables_Text_COL" & iX, lvReceivables.ColumnHeaders(iX).Text)
    Next
    
    For ix = 1 To lvExtra.ColumnHeaders.Count
        lvExtra.ColumnHeaders(ix).Width = GetSetting("projectalpha", "Accounts_frm", "lvExtra_Width_COL" & ix, lvExtra.ColumnHeaders(ix).Width)
        '¤lvExtra.ColumnHeaders(iX).Text = GetSetting("projectalpha", "Accounts_frm", "lvExtra_Text_COL" & iX, lvExtra.ColumnHeaders(iX).Text)
    Next
            
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



Public Function loadRS()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "loadRS"
    Const ContainerName = "frmAccounts"
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


    If bDebug = True Then On Error GoTo 0 Else On Error GoTo ErrorOccur

    Dim rsload As adodb.Recordset
    Dim iRecCount As Double
    Dim ix As Double
    Dim itmX As ListItem
   
            
    bResult = MySQL.OpenTable(directConn, rsload, , MySQL.virtualisp("select distinct  accountinfo.AccountName ,invoiceout.RecID ,invoiceout.acci_RecID ,invoiceout.AmountDue ,invoiceout.Description " + _
                                                  " ,invoiceout.GSTCharged ,invoiceout.TotalDue ,invoiceout.PaymentDue ,invoiceout.AmountPaid ,invoiceout.GSTRefunded ,invoiceout.AmountRefunded from invoiceout, accountinfo Where accountinfo.RecID = invoiceout.AccI_RecID AND (invoiceout.AmountRefunded + invoiceout.GSTRefunded + invoiceout.AmountPaid) < invoiceout.TotalDue order By accountinfo.AccountName", "invoiceout") + "")

            
          
          
    If rsload.State = adStateOpen Then
        If rsload.RecordCount > 0 Then
            If Not rsload.EOF And Not rsload.BOF Then
                rsload.MoveFirst
                While Not rsload.EOF And Err.Number = 0
                    If pb1.Value + 1 < pb1.Max Then pb1.Value = pb1.Value + 1
                    pb1.Refresh
                    'gSleep
                    Set itmX = lvReceivables.ListItems.Add(, "r" & rsload!RecID, IIf(IsNull(rsload!AccountName), "(null)", rsload!AccountName), IIf(fIcon.il16x16.ListImages.Count = 1, 1, "inv"), IIf(fIcon.il16x16.ListImages.Count = 1, 1, "inv"))
                    itmX.SubItems(6) = IIf(IsNull(rsload!Description), "", rsload!Description)
                    
                    itmX.SubItems(1) = Format(IIf(IsNull(rsload!AmountDue), 0, rsload!AmountDue), "Currency")
                    itmX.SubItems(2) = Format(IIf(IsNull(rsload!GSTCharged), 0, rsload!GSTCharged), "Currency")
                    itmX.SubItems(3) = Format(IIf(IsNull(rsload!TotalDue), 0, rsload!TotalDue), "Currency")
                    itmX.SubItems(4) = Format(IIf(IsNull(rsload!PaymentDue), #9/19/1950#, rsload!PaymentDue), "dd-mm-yyyy Hh:Nn:Ss")
                    itmX.SubItems(5) = Format(IIf(IsNull(rsload!AmountPaid), 0, rsload!AmountPaid), "Currency")
                    itmX.SubItems(7) = Format(IIf(IsNull(rsload!AmountRefunded + rsload!GSTRefunded), 0, rsload!AmountRefunded + rsload!GSTRefunded), "Currency")
                    itmX.Tag = rsload!acci_RecID
                    rsload.MoveNext
                Wend
             End If
        End If
    End If
            
    bResult = MySQL.OpenTable(directConn, rsload, , MySQL.virtualisp("select distinct accountinfo.AccountName ,invoicein.AmountPaid ,invoicein.PaidWhen ,invoicein.AmountUsed ,invoicein.Sub ,invoicein.GSTCharged ,invoicein.RecID ,invoicein.TotalPaid from invoicein, accountinfo Where accountinfo.RecID = invoicein.AccI_RecID AND AmountPaid > AmountUsed", "invoicein") + "")
    
    If rsload.State = adStateOpen Then
        If rsload.RecordCount > 0 Then
            If Not rsload.EOF And Not rsload.BOF Then
                rsload.MoveFirst
                While Not rsload.EOF And Err.Number = 0
                    
                    pb1.Refresh
                    
                    Set itmX = lvExtra.ListItems.Add(, "r" & rsload!RecID, rsload!AccountName, IIf(fIcon.il16x16.ListImages.Count = 1, 1, "vault"), IIf(fIcon.il16x16.ListImages.Count = 1, 1, "vault"))
                    itmX.SubItems(1) = Format(IIf(IsNull(rsload!TotalPaid), 0, rsload!TotalPaid), "Currency")
                    itmX.SubItems(2) = Format(IIf(IsNull(rsload!AmountPaid), 0, rsload!AmountPaid), "Currency")
                    itmX.SubItems(3) = Format(IIf(IsNull(rsload!PaidWhen), #9/19/1950#, rsload!PaidWhen), "dd-mm-yyyy Hh:Nn:Ss")
                    itmX.SubItems(4) = Format(IIf(IsNull(rsload!GSTCharged), 0, rsload!GSTCharged), "Currency")
                    itmX.SubItems(5) = Format(IIf(IsNull(rsload!AmountUsed), 0, rsload!AmountUsed), "Currency")
                    itmX.SubItems(6) = Format(IIf(IsNull(rsload!AmountPaid), 0, rsload!AmountPaid) - IIf(IsNull(rsload!AmountUsed), 0, rsload!AmountUsed), "Currency")
                    itmX.SubItems(7) = Format(IIf(IsNull(rsload!SUB), 0, rsload!SUB), "Currency")
                    rsload.MoveNext
                    pb1.Value = pb1.Value + 1
                    
                Wend
             End If
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
