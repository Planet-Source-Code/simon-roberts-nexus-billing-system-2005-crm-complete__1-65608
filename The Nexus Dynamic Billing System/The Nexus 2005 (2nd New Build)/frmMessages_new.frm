VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmMessages 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Auto Message Editor"
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13905
   Icon            =   "frmMessages_new.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   13905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMessage 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   4605
      Index           =   1
      Left            =   4200
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Top             =   2700
      Width           =   7275
   End
   Begin VB.TextBox txtMessage 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   4605
      Index           =   0
      Left            =   3780
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   2490
      Width           =   7275
   End
   Begin MSComctlLib.TabStrip ts 
      Height          =   7665
      Left            =   3660
      TabIndex        =   6
      Top             =   1800
      Width           =   10185
      _ExtentX        =   17965
      _ExtentY        =   13520
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Text"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "HTML"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilMessageIcons 
      Left            =   2910
      Top             =   2010
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessages_new.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessages_new.frx":0D1C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fields"
      Height          =   855
      Left            =   3660
      TabIndex        =   2
      Top             =   900
      Width           =   10185
      Begin VB.CommandButton Command1 
         Caption         =   "+"
         Height          =   375
         Left            =   9660
         TabIndex        =   4
         Top             =   270
         Width           =   375
      End
      Begin VB.ComboBox cmbFields 
         Height          =   360
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   300
         Width           =   9435
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Messages"
      Height          =   8535
      Left            =   120
      TabIndex        =   0
      Top             =   900
      Width           =   3465
      Begin MSComctlLib.ListView lvMessages 
         Height          =   8085
         Left            =   150
         TabIndex        =   1
         Top             =   300
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   14261
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ilMessageIcons"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   14
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Messages"
            Object.Width           =   5292
         EndProperty
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto Response Message Center"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1050
      TabIndex        =   5
      Top             =   120
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   90
      Picture         =   "frmMessages_new.frx":116E
      Stretch         =   -1  'True
      Top             =   -30
      Width           =   870
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   30
      X2              =   13950
      Y1              =   450
      Y2              =   420
   End
End
Attribute VB_Name = "frmMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bChanged As Boolean
Dim bGotFocus As Byte

Private Sub SaveChanges()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "SaveChanges"
    Const ContainerName = "frmMessages"
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


    Dim bResult As Boolean
    Dim rsload As adodb.Recordset
    
    If Not lvMessages.SelectedItem Is Nothing Then
        bResult = MySQL.OpenTable(directConn, rsload, , "select RecID, MessageDraft, HTMLDraft from autoMessages where RecID = " & Mid(lvMessages.SelectedItem.Key, 2))
    
        If rsload.RecordCount > 0 Then
            rsload!MessageDraft = txtMessage(0).Text
            rsload!HTMLDraft = txtMessage(1).Text
            rsload.Update
        End If
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

Private Sub Command1_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Command1_Click"
    Const ContainerName = "frmMessages"
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


    If cmbFields.ListIndex > -1 Then
        If txtMessage(bGotFocus).SelStart = 0 Then txtMessage(bGotFocus).SelStart = 1
        If Len(txtMessage(bGotFocus)) > 0 And txtMessage(bGotFocus).SelStart < Len(txtMessage(bGotFocus)) Then
            txtMessage(bGotFocus).Text = Left(txtMessage(bGotFocus), txtMessage(bGotFocus).SelStart) + cmbFields.List(cmbFields.ListIndex) + Mid(txtMessage(bGotFocus), txtMessage(bGotFocus).SelStart + 1)
            txtMessage(bGotFocus).SelStart = txtMessage(bGotFocus).SelStart + Len(cmbFields.List(cmbFields.ListIndex))
        ElseIf txtMessage(bGotFocus).SelStart = Len(txtMessage(bGotFocus)) Then
            txtMessage(bGotFocus).Text = txtMessage(bGotFocus).Text + cmbFields.List(cmbFields.ListIndex)
            txtMessage(bGotFocus).SelStart = Len(txtMessage(bGotFocus))
        Else
            txtMessage(bGotFocus).Text = cmbFields.List(cmbFields.ListIndex)
            txtMessage(bGotFocus).SelStart = Len(txtMessage(bGotFocus))
        End If
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

Private Sub Form_Load()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Load"
    Const ContainerName = "frmMessages"
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

   
    Call ts_Click
    
    Dim rsload As adodb.Recordset
    Dim bResult As Boolean
    Dim itmX As ListItem
    
    bResult = MySQL.OpenTable(directConn, rsload, , MySQL.virtualisp("select Description, RecID, MSGType from autoMessages"))
    
    If rsload.RecordCount > 0 Then
        While Not rsload.EOF And Err.Number = 0
            Set itmX = lvMessages.ListItems.Add(, "k" & rsload!RecID, IIf(IsNull(rsload!Description), "", rsload!Description))
            itmX.Tag = rsload!msgType
            Select Case rsload!msgType
            Case 0, 1
                itmX.SmallIcon = 1
            Case 2
                itmX.SmallIcon = 2
            End Select
            rsload.MoveNext
        Wend
    
    
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
    Const ContainerName = "frmMessages"
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


    If bChanged = True Then SaveChanges
    
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

Private Sub lvMessages_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvMessages_ItemClick"
    Const ContainerName = "frmMessages"
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


    If bChanged = True Then SaveChanges
        
    cmbFields.Clear
    
    Select Case Item.Tag
    Case 0, 1
        cmbFields.AddItem "{Date}"
        cmbFields.AddItem "{AccountName}"
        cmbFields.AddItem "{ContactName}"
        cmbFields.AddItem "{EmailAddress}"
        cmbFields.AddItem "{Username}"
        cmbFields.AddItem "{Password}"
        cmbFields.AddItem "{sfCycle_Upload}"
        cmbFields.AddItem "{sfCycle_Download}"
        cmbFields.AddItem "{sfCycle_Mins}"
        cmbFields.AddItem "{sfCycle_HrsMins}"
        cmbFields.AddItem "{NextCycle}"
        cmbFields.AddItem "{DOB}"
        cmbFields.AddItem "{Classification}"
        cmbFields.AddItem "{Realm}"
        cmbFields.AddItem "{ActivationDate}"
        cmbFields.AddItem "{ExpiryDate}"
        cmbFields.AddItem "{Activation}"
        cmbFields.AddItem "{Deactivation}"
    Case 2
        cmbFields.AddItem "{Date}"
        cmbFields.AddItem "{AccountName}"
        cmbFields.AddItem "{ContactName}"
        cmbFields.AddItem "{EmailAddress}"
        cmbFields.AddItem "{TotalDue}"
        cmbFields.AddItem "{AmountPaid}"
        cmbFields.AddItem "{TEXT_Invoice_Table}"
        cmbFields.AddItem "{HTML_Invoice_Table}"
        cmbFields.AddItem "{FirstAddress}"
        cmbFields.AddItem "{NextCycle}"
    End Select
    
    Dim bResult As Boolean
    Dim rsload As adodb.Recordset
    
    bResult = MySQL.OpenTable(directConn, rsload, , "select MessageDraft, HTMLDraft from autoMessages where RecID = " & Mid(Item.Key, 2))
    
    If rsload.RecordCount > 0 Then
        txtMessage(0).Text = IIf(IsNull(rsload!MessageDraft), "", rsload!MessageDraft)
        txtMessage(1).Text = IIf(IsNull(rsload!HTMLDraft), "", rsload!HTMLDraft)
        bChanged = False
    End If
    
    'cmdSave.Enabled = True
    
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

Private Sub ts_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "ts_Click"
    Const ContainerName = "frmMessages"
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


    Dim bx As Byte
    
    For bx = 0 To 1
        txtMessage(bx).Visible = False
    Next
        
    txtMessage(ts.SelectedItem.Index - 1).Move ts.ClientLeft, ts.ClientTop, ts.ClientWidth, ts.ClientHeight
    txtMessage(ts.SelectedItem.Index - 1).Visible = True
    
        
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

Private Sub txtMessage_Change(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtMessage_Change"
    Const ContainerName = "frmMessages"
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

    
    bChanged = True
    
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

Private Sub txtMessage_GotFocus(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtMessage_GotFocus"
    Const ContainerName = "frmMessages"
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


    bGotFocus = Index
    
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
