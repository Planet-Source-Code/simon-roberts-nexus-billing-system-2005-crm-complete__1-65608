VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAlias 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "eMail Alias'"
   ClientHeight    =   7725
   ClientLeft      =   4290
   ClientTop       =   3480
   ClientWidth     =   9270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAlias.frx":0000
   ScaleHeight     =   7725
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
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
      Left            =   120
      TabIndex        =   7
      Top             =   7080
      Width           =   1605
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   7590
      TabIndex        =   6
      Top             =   7080
      Width           =   1485
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0066B0DD&
      Caption         =   "Redirected Email Addresses"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5235
      Left            =   90
      TabIndex        =   4
      Top             =   1710
      Width           =   8985
      Begin VB.Frame Frame2 
         BackColor       =   &H0066B0DD&
         Caption         =   "Details"
         Height          =   1125
         Left            =   150
         TabIndex        =   8
         Top             =   270
         Width           =   8655
         Begin VB.CommandButton cmdAddAlias 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   8130
            TabIndex        =   13
            Top             =   210
            Width           =   435
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
            ForeColor       =   &H00BA3F3F&
            Height          =   405
            Left            =   3360
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   210
            Width           =   4725
         End
         Begin VB.TextBox txtLocal 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   120
            TabIndex        =   11
            Top             =   210
            Width           =   3195
         End
         Begin VB.OptionButton optType 
            Caption         =   "Site"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   660
            Value           =   -1  'True
            Width           =   3780
         End
         Begin VB.OptionButton optType 
            Caption         =   "System"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   5850
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   690
            Width           =   2700
         End
      End
      Begin MSComctlLib.ListView lvAlias 
         Height          =   3555
         Left            =   150
         TabIndex        =   5
         Top             =   1530
         Width           =   8685
         _ExtentX        =   15319
         _ExtentY        =   6271
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Local"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Domain"
            Object.Width           =   6350
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.ComboBox cmbeMail 
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
      Left            =   3060
      TabIndex        =   2
      Text            =   "cmbeMail"
      Top             =   900
      Width           =   6045
   End
   Begin VB.TextBox txtfield 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   3060
      MaxLength       =   50
      TabIndex        =   1
      Top             =   150
      Width           =   6015
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "e-Mail Alias"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0066B0DD&
      Height          =   435
      Left            =   90
      TabIndex        =   14
      Top             =   1080
      Width           =   1995
   End
   Begin VB.Image Image1 
      Height          =   1050
      Left            =   60
      Picture         =   "frmAlias.frx":B757A
      Stretch         =   -1  'True
      Top             =   60
      Width           =   1230
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "e-Mail Address to redirect to"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0066B0DD&
      Height          =   285
      Index           =   1
      Left            =   4470
      TabIndex        =   3
      Top             =   1320
      Width           =   3180
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0066B0DD&
      Height          =   285
      Index           =   0
      Left            =   5085
      TabIndex        =   0
      Top             =   540
      Width           =   1875
   End
   Begin VB.Menu mnuPopups 
      Caption         =   "Popups"
      Visible         =   0   'False
      Begin VB.Menu mnuPopups_Delete 
         Caption         =   "Delete Alias"
      End
   End
End
Attribute VB_Name = "frmAlias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public oCNT As frmCustomerRec
Public BaseURL As String
Public ContactName As String
Public iCloseState As frm_CloseStates
Public Session As String

Private Sub cmdAddAlias_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdAddAlias_Click"
    Const ContainerName = "frmAlias"
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


    If Trim(txtLocal) = "" Then
        MsgBox "You must specify a local name for the alias!"
        txtLocal.SetFocus
        Exit Sub
    End If
    
    Dim imtX As ListItem
    
    Set itmx = lvAlias.ListItems.Add(, , txtLocal)
    itmx.SubItems(1) = cmbDomain.Text
    Dim bx As Byte
    For bx = optType.LBound To optType.UBound
     
        If optType(bx).Value = True Then itmx.SubItems(2) = optType(bx).Caption
    
    Next
    
    Call oCNT.osub.col_subAliases.Add("NEW" & Session & oCNT.osub.col_subAliases.Count + 1, oCNT.osub.fRecID, _
         itmx.SubItems(2), itmx.Text + "@" + cmbDomain.Text, cmbeMail.Text, True, "NEW" & Session & oCNT.osub.col_subAliases.Count + 1)
    
    cmdSave.Enabled = True
    
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

Private Sub cmdSave_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdSave_Click"
    Const ContainerName = "frmAlias"
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

    Dim itmx As ListItem
    If txtField(0).Text = "" Then
        MsgBox "You must specify the customers Name!"
        txtField(0).SetFocus
        Exit Sub
    End If
    
    If cmbeMail.ListIndex = -1 Then
        MsgBox "You must specify an email address to redirect to"
        cmbeMail.SetFocus
        Exit Sub
    End If
    
    If lvAlias.ListItems.Count = 0 Then
        MsgBox "No Redirected email addresses specified"
        Exit Sub
    End If
    
    Me.BaseURL = cmbeMail.Text
    Me.ContactName = txtField(0).Text
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


Private Sub cmdClose_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdClose_Click"
    Const ContainerName = "frmAlias"
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


    iCloseState = frmCloseCancel
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

Private Sub Form_Load()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Load"
    Const ContainerName = "frmAlias"
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

    If Session = "" Then Session = GetSessionChar(Session, Me.hwnd, 14)

If bDebug = True Then On Error GoTo 0 Else On Error GoTo ErrorOccur

    Dim rsLoad As ADODB.Recordset
    
    cmbDomain.AddItem Login.sVISPDomain
    'If Login.sVISPDomain <> "ep.net.au" Then cmbDomain.AddItem "ep.net.au"
    
    txtField(0).Text = sContactName
    
    If oCNT.osub.col_subEmails.Count > 0 Then
        Dim oSVR As cls_subServices
        For Each oSVR In oCNT.osub.col_subServices
            If oSVR.ServiceID = 4 Then
                cmbeMail.AddItem oSVR.Username & "@" & IIf(oSVR.BaseURL = "", "ep.net.au", oSVR.BaseURL)
            End If
        Next
    
        If cmbeMail.ListIndex = -1 Then
            If cmbeMail.ListCount > 0 Then
                cmbeMail.ListIndex = 0
            End If
        End If
    End If
    
    Dim oDM As cls_Domains
    
    For Each oDM In oCNT.osub.col_Domains
        cmbDomain.AddItem oDM.Domain
    Next
    
    Dim bx As Integer
    For bx = 0 To cmbDomain.ListCount - 1
        If cmbDomain.List(bx) = sDomain Then
            cmbDomain.ListIndex = bx
            Exit For
        End If
    Next
    
    If cmbDomain.ListIndex = -1 Then cmbDomain.ListIndex = 0
    
    Dim oAls As cls_subAliases
    For Each oAls In oCNT.osub.col_subAliases
        With lvAlias.ListItems.Add(, oAls.Key, Left(oAls.eMail, InStr(oAls.eMail, "@") - 1))
            .SubItems(1) = Mid(oAls.eMail, InStr(oAls.eMail, "@") + 1)
            .SubItems(2) = oAls.ftype
        End With
    Next
        
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_QueryUnload"
    Const ContainerName = "frmAlias"
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


    
    sBaseURL = cmbeMail.Text
    sContactName = txtField(0).Text
    
    On Error Resume Next
    
    DF1 = lvAlias.ListItems(1).Text & "@" & lvAlias.ListItems(1).SubItems(1)
    DF2 = lvAlias.ListItems(2).Text & "@" & lvAlias.ListItems(2).SubItems(1)
    DF3 = lvAlias.ListItems(3).Text & "@" & lvAlias.ListItems(3).SubItems(1)
    DF4 = lvAlias.ListItems(4).Text & "@" & lvAlias.ListItems(4).SubItems(1)
    DF5 = lvAlias.ListItems(5).Text & "@" & lvAlias.ListItems(5).SubItems(1)

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

Private Sub lvAlias_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvAlias_MouseDown"
    Const ContainerName = "frmAlias"
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


    
    Select Case Button
    Case vbRightButton
            
        PopupMenu mnuPopups
        
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

Private Sub mnuPopups_Delete_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuPopups_Delete_Click"
    Const ContainerName = "frmAlias"
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

    If Not lvAlias.SelectedItem Is Nothing Then
        Select Case MsgBox("Are you sure you wish to delete " & lvAlias.SelectedItem.Text & "@" & lvAlias.SelectedItem.SubItems(1) & " from the aliases list?", vbYesNo + vbQuestion, "Delete Alias")
        Case vbYes
            If lvAlias.SelectedItem.Key = "" Then
                lvAlias.ListItems.Remove lvAlias.SelectedItem.Index
            Else
                MySQL.Execute ADOConn, "Delete from acci_aliases where RecID = " & Mid(lvAlias.SelectedItem.Key, 2)
                lvAlias.ListItems.Remove lvAlias.SelectedItem.Index
            End If
        End Select
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
