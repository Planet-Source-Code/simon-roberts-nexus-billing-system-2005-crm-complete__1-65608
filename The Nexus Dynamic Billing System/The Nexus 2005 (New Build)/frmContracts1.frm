VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmContracts1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configure setup contracts that are available"
   ClientHeight    =   8505
   ClientLeft      =   3540
   ClientTop       =   2985
   ClientWidth     =   10545
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmContracts1.frx":0000
   ScaleHeight     =   567
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   703
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame7 
      BackColor       =   &H00A3A3FE&
      Caption         =   "Contract Description"
      Height          =   1965
      Left            =   4800
      TabIndex        =   3
      Top             =   60
      Width           =   5655
      Begin VB.CommandButton cmdContract 
         BackColor       =   &H00A3A3FE&
         Caption         =   "Update Table"
         Height          =   345
         Index           =   0
         Left            =   4050
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1530
         Width           =   1515
      End
      Begin VB.TextBox txtFee 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0DFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   1260
         TabIndex        =   9
         Tag             =   "FeePerHour"
         Top             =   1500
         Width           =   1395
      End
      Begin VB.TextBox txtFee 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0DFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   1260
         TabIndex        =   8
         Tag             =   "FeePerBlock"
         Top             =   1050
         Width           =   1395
      End
      Begin VB.TextBox txtFee 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0DFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   1260
         TabIndex        =   7
         Tag             =   "JoiningFee"
         Top             =   600
         Width           =   1395
      End
      Begin VB.TextBox txtFee 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0DFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   3990
         TabIndex        =   6
         Tag             =   "PeriodFee"
         Top             =   1050
         Width           =   1545
      End
      Begin VB.TextBox txtFee 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0DFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   3990
         TabIndex        =   5
         Tag             =   "Termination"
         Top             =   600
         Width           =   1545
      End
      Begin VB.TextBox txtField 
         BackColor       =   &H00E0DFFF&
         Height          =   285
         Index           =   0
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   4
         Tag             =   "Description"
         Top             =   240
         Width           =   4275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Extra Per Hour:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Extra Per MB:"
         Height          =   195
         Index           =   5
         Left            =   210
         TabIndex        =   15
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Joining Fee:"
         Height          =   195
         Index           =   4
         Left            =   330
         TabIndex        =   14
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Billing Cycle Fee:"
         Height          =   195
         Index           =   3
         Left            =   2730
         TabIndex        =   13
         Top             =   1140
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Termination Fee:"
         Height          =   195
         Index           =   2
         Left            =   2730
         TabIndex        =   12
         Top             =   690
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   11
         Top             =   300
         Width           =   840
      End
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   315
      Left            =   5040
      TabIndex        =   2
      Top             =   8070
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load the checked setup contracts as support by my Network"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   8040
      Width           =   4845
   End
   Begin MSComctlLib.ListView lvContracts 
      Height          =   5775
      Left            =   90
      TabIndex        =   0
      Tag             =   "0"
      Top             =   2100
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   10186
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   7697007
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "Description"
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "NoPeriods"
         Text            =   "Interval Length (ttl)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "TypePeriods"
         Text            =   "Interval Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "Termination"
         Text            =   "Termination Fee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "JoiningFee"
         Text            =   "Joining Fee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "PeriodFee"
         Text            =   "Billing Cycle Fee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "FeePerBlock"
         Text            =   "Fee Per MB Block"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "FeePerHour"
         Text            =   "Per Extra Hour"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Tag             =   "RecID"
         Text            =   "Contract ID"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmContracts1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ptRecID As Variant
Public PeriodFee As Single

Private Sub cmdContract_Click(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdContract_Click"
    Const ContainerName = "frmContracts1"
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


    If Not lvContracts.SelectedItem Is Nothing Then
        lvContracts.SelectedItem.SubItems(3) = txtFee(6)
        lvContracts.SelectedItem.SubItems(5) = txtFee(7)
        lvContracts.SelectedItem.SubItems(4) = txtFee(8)
        lvContracts.SelectedItem.SubItems(6) = txtFee(9)
        lvContracts.SelectedItem.SubItems(7) = txtFee(10)
        'lvContracts.selectedItem.SubItems(8) = lvContracts.Tag
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

Private Sub Command1_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Command1_Click"
    Const ContainerName = "frmContracts1"
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


    Dim li As Long
    
    
    pb.Value = 0
    pb.Max = lvContracts.ListItems.Count + 1
    
    lvContracts.Enabled = False
    Command1.Enabled = False
    frmAgent.oChar.Speak "Creating Contract and Setup Fee Matrix"
    frmAgent.oChar.Play "Processing"
    
    
    Dim rsLoad As ADODB.Recordset
    
    Call MySQL.OpenTable(ADOConn, rsLoad, , "select TemplateID from plantypes where RecID = " & ptRecID & "")

    
    Dim itmy As ListItem
    
    Call MySQL.Execute(ADOConn, "delete from contractsruntime where ptRecID = '" & ptRecID & "'")
    
    For li = 1 To pb.Max - 1
        Set itmy = lvContracts.ListItems(li)
        If itmy.Checked = True Then
                            
            Call MySQL.Execute(ADOConn, "INSERT INTO contractsruntime (ptRecID, TemplateID, ContractID, VirtualID, Termination, JoiningFee, PeriodFee, FeePerBlock, FeePerHour) " + _
                                        "VALUES ('" & ptRecID & "','" & rsLoad!TemplateID & "','" & Mid(itmy.Key, 2) & "','" & Login.lVirtualID & "','" & itmy.SubItems(3) & "','" & itmy.SubItems(4) & "','" & itmy.SubItems(5) & "','" & itmy.SubItems(6) & "','" & itmy.SubItems(7) & "')")
            
        End If
        pb.Value = li
    Next li
    
    frmAgent.oChar.Stop
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
    Const ContainerName = "frmContracts1"
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


    Call GUI.LoadColWidths(lvContracts, Me)
    
        Dim rsLoad As ADODB.Recordset
        
        Call MySQL.OpenTable(ADOConn, rsLoad, , "select TemplateID from plantypes where RecID = " & ptRecID & "")
        
        Call MySQL.OpenTable(ADOConn, rsLoad, , "select * from contracttemplates where ptRecID = " & rsLoad!TemplateID & " and bDeleted = 0")
        Dim itmX As ListItem
        Dim bx As Byte
        lvContracts.ListItems.Clear
        If rsLoad.State = adStateOpen Then
            If rsLoad.RecordCount > 0 Then
                While Not rsLoad.EOF And Err.Number = 0
                    Set itmX = lvContracts.ListItems.Add(, "c" & rsLoad!RecID, IIf(IsNull(rsLoad!Description), "(null)", rsLoad!Description))
                    For bx = 2 To lvContracts.ColumnHeaders.Count
                        If lvContracts.ColumnHeaders(bx).Tag = "PeriodFee" Then
                            
                            itmX.SubItems(bx - 1) = 0
                        
                        Else
                        
                            itmX.SubItems(bx - 1) = IIf(IsNull(rsLoad(lvContracts.ColumnHeaders(bx).Tag)), "0", rsLoad(lvContracts.ColumnHeaders(bx).Tag))
                        End If
                    Next
                    itmX.Checked = True
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_QueryUnload"
    Const ContainerName = "frmContracts1"
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


    Call GUI.SaveColWidths(lvContracts, Me)
    
    
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

Private Sub lvContracts_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvContracts_ItemClick"
    Const ContainerName = "frmContracts1"
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


    txtField(0) = Item.Text
    txtFee(6) = Item.SubItems(3)
    txtFee(7) = Item.SubItems(5)
    txtFee(8) = Item.SubItems(4)
    txtFee(9) = Item.SubItems(6)
    txtFee(10) = Item.SubItems(7)
    
    lvContracts.Tag = Item.SubItems(8)
    
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
