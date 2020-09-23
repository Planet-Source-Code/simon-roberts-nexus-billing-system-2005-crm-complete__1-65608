VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImpContracts 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Import Contracts from pre-existing framework"
   ClientHeight    =   7635
   ClientLeft      =   2775
   ClientTop       =   4620
   ClientWidth     =   14520
   Icon            =   "frmImpContracts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   509
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   968
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar pb 
      Height          =   165
      Left            =   90
      TabIndex        =   5
      Top             =   7380
      Width           =   14325
      _ExtentX        =   25268
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Finished"
      Height          =   585
      Left            =   13110
      TabIndex        =   4
      Top             =   5520
      Width           =   1305
   End
   Begin VB.CommandButton cmbImport 
      Caption         =   "Import Checked Contracts"
      Height          =   1125
      Left            =   13110
      TabIndex        =   3
      Top             =   6180
      Width           =   1305
   End
   Begin MSComctlLib.TreeView tvServiceTypes 
      Height          =   7185
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   12674
      _Version        =   393217
      Indentation     =   441
      LabelEdit       =   1
      Style           =   7
      HotTracking     =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lvAccounts 
      Height          =   3795
      Left            =   4140
      TabIndex        =   1
      Top             =   120
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   6694
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "Description<No Format"
         Text            =   "Description"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "VendorID<select vname as nResult from vendors where RecID = "
         Text            =   "Vendor"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "VendorPartID<No Format"
         Text            =   "Vendor Code"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "SubPartID<No Format"
         Text            =   "Sub Part Number"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "PeriodFee<Currency"
         Text            =   "Monthly Fee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "MBPerPeriod<###,###,###,###,### MB's"
         Text            =   "Montly Data"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "HoursPerPeriod<###,###,###,###,### Hrs"
         Text            =   "Monthly Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "SessionTimeout<No Format"
         Text            =   "Session Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Tag             =   "IdleTimeout<No Format"
         Text            =   "Idle"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvContracts 
      Height          =   3285
      Left            =   4140
      TabIndex        =   2
      Tag             =   "0"
      Top             =   4020
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   5794
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
Attribute VB_Name = "frmImpContracts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ptRecID As Variant


Private Sub cmbImport_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmbImport_Click"
    Const ContainerName = "frmImpContracts"
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



    pb.Value = 0
    pb.Max = lvContracts.ListItems.Count + 1
    
    If pb.Max = 1 Then Exit Sub
    
    Dim il As Long
    
    Dim SQL As String
    Dim itmX As ListItem
    Dim itmy As ListItem
    Dim rsScan As ADODB.Recordset
    Dim rsload As ADODB.Recordset
    
    frmAgent.oChar.Play "Writing"
    
    lvContracts.Enabled = False
    lvAccounts.Enabled = False
    tvServiceTypes.Enabled = False
    
    For il = 1 To pb.Max - 1
        Set itmy = lvContracts.ListItems(il)
        If itmy.Checked = True Then
                    
        
        
      On Error Resume Next
      Do
          Err.Clear
          lvContracts.Tag = MySQL.GetTMPRecID("contracttemplates", ADOConn)
          Call MySQL.Execute(ADOConn, "INSERT into contracttemplates (RecID, ptRecID) VALUES ('" & lvContracts.Tag & "','" & Me.ptRecID & "')")
      Loop Until Err.Number = 0
      

    Call MySQL.OpenTable(ADOConn, rsload, , "select PeriodFee from plantemplates where RecID = '" & ptRecID & "'")
    
        SQL = "Update contracttemplates set "
        SQL = SQL + "`" & lvContracts.ColumnHeaders(1).Tag & "` = '" & itmy.Text & "', "
        SQL = SQL + "`" & lvContracts.ColumnHeaders(2).Tag & "` = '" & itmy.SubItems(1) & "', "
        SQL = SQL + "`" & lvContracts.ColumnHeaders(3).Tag & "` = '" & itmy.SubItems(2) & "', "
        SQL = SQL + "`" & lvContracts.ColumnHeaders(4).Tag & "` = '" & itmy.SubItems(3) & "', "
        SQL = SQL + "`" & lvContracts.ColumnHeaders(5).Tag & "` = '" & itmy.SubItems(4) & "', "
        SQL = SQL + "`" & lvContracts.ColumnHeaders(6).Tag & "` = '0', "
        SQL = SQL + "`" & lvContracts.ColumnHeaders(7).Tag & "` = '" & itmy.SubItems(6) & "' "
        SQL = SQL + " where RecID = " & lvContracts.Tag
        
        MySQL.Execute ADOConn, SQL
    
        
        Call MySQL.OpenTable(ADOConn, rsScan, , "select * from flags_tempextras where ContractID = '" & Mid(itmy.Key, 2) & "'")
                    
        If rsScan.State = adStateOpen Then
            If rsScan.RecordCount > 0 Then
                While Not rsScan.EOF And Err.Number = 0
                    On Error GoTo 0
                    
                    Call MySQL.Execute(ADOConn, "INSERT INTO flags_tempextras (PlanType, NumberOf, Checked, ContractID) VALUES ('" & Val(rsScan!PlanType) & "','" & Val(rsScan!NumberOf) & "','" & Val(rsScan!Checked) & "','" & Val(lvContracts.Tag) & "')")
                    
                    rsScan.MoveNext
                    
                Wend
            End If
        End If
                    
        End If
        pb.Value = il
    Next il
    
    frmAgent.oChar.StopAll
    
    pb.Value = pb.Max
    
    lvContracts.Enabled = True
    lvAccounts.Enabled = True
    tvServiceTypes.Enabled = True
    
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
    Const ContainerName = "frmImpContracts"
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

Private Sub Form_Load()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Load"
    Const ContainerName = "frmImpContracts"
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


    PopulateList
    Call GUI.LoadColWidths(lvAccounts, Me)
    Call GUI.LoadColWidths(lvContracts, Me)
    
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

Public Function PopulateList()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "PopulateList"
    Const ContainerName = "frmImpContracts"
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


    Dim bResult As Boolean
    Dim rsload As ADODB.Recordset
    Dim NodX As Node
    Dim NodeX As Node
    
    bResult = MySQL.OpenTable(ADOConn, rsload, "servicetypes")
    
    tvServiceTypes.NodeS.Clear
    
    If rsload.RecordCount > 0 Then
        While Not rsload.EOF And Err.Number = 0
            If Not IsNull(rsload!SubofRecID) Then
                If rsload!SubofRecID <> 0 Then
                    Set NodeX = tvServiceTypes.NodeS("k" & rsload!SubofRecID)
                    Set NodX = tvServiceTypes.NodeS.Add(NodeX.Key, tvwChild, "k" & rsload!RecID, IIf(IsNull(rsload!Description), "Description Not set", rsload!Description))
                    NodX.Tag = rsload!ServiceKey
                    NodeX.Expanded = True
                Else
                    Set NodX = tvServiceTypes.NodeS.Add(, , "k" & rsload!RecID, IIf(IsNull(rsload!Description), "Description Not set", rsload!Description))
                    NodX.Tag = rsload!ServiceKey
                End If
            Else
                Set NodX = tvServiceTypes.NodeS.Add(, , "k" & rsload!RecID, IIf(IsNull(rsload!Description), "Description Not set", rsload!Description))
                NodX.Tag = rsload!ServiceKey
            End If
            rsload.MoveNext
        Wend

    End If
        
   'Exit Function
    
    
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_QueryUnload"
    Const ContainerName = "frmImpContracts"
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


  Call GUI.SaveColWidths(lvAccounts, Me)
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

Private Sub lvAccounts_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvAccounts_ItemClick"
    Const ContainerName = "frmImpContracts"
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


        Dim rsload As ADODB.Recordset
        
        
        Call MySQL.OpenTable(ADOConn, rsload, , "select * from contracttemplates where ptRecID = " & Mid(Item.Key, 2) & " and bDeleted = 0")
        Dim itmX As ListItem
        Dim bx As Byte
        lvContracts.ListItems.Clear
        If rsload.State = adStateOpen Then
            If rsload.RecordCount > 0 Then
                While Not rsload.EOF And Err.Number = 0
                    Set itmX = lvContracts.ListItems.Add(, "c" & rsload!RecID, IIf(IsNull(rsload!Description), "(null)", rsload!Description))
                    For bx = 2 To lvContracts.ColumnHeaders.Count
                        itmX.SubItems(bx - 1) = IIf(IsNull(rsload(lvContracts.ColumnHeaders(bx).Tag)), "0", rsload(lvContracts.ColumnHeaders(bx).Tag))
                    Next
                    itmX.Checked = True
                    rsload.MoveNext
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

Private Sub tvservicetypes_NodeClick(ByVal Node As MSComctlLib.Node)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "tvservicetypes_NodeClick"
    Const ContainerName = "frmImpContracts"
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


    On Error Resume Next
    Dim bResult As Boolean
    
'    Frame1(0).Caption = "Services and Plans: " & Node.Text
    
 '   Command1.Enabled = False
    Dim rsPlanTemp As ADODB.Recordset
    
    bResult = MySQL.OpenTable(ADOConn, rsPlanTemp, , "select * from plantemplates Where ServiceID = " & Mid(Node.Key, 2))
    
    'rsPlanTemp.Filter
    
    
    Dim bNumeric As Boolean
    Dim rsload As ADODB.Recordset
    Dim SQL As String
    
    
    lvAccounts.ListItems.Clear
    If rsPlanTemp.RecordCount > 0 Then
        While Not rsPlanTemp.EOF And Err.Number = 0
            
            'cmbRollover.AddItem rsPlanTemp!Description
            'cmbRollover.ItemData(cmbRollover.ListCount - 1) = rsPlanTemp!RecID
            
            Set itmX = lvAccounts.ListItems.Add(, "r" & rsPlanTemp!RecID, "")
            For X = 1 To lvAccounts.ColumnHeaders.Count
                If lvAccounts.ColumnHeaders(X).Tag <> "" Then
                    If IsNull(MySQL.OP(rsPlanTemp, Left(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") - 1))) Then
                    
                    Else
                    
                        Select Case Mid(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") + 1, Len(lvAccounts.ColumnHeaders(X).Tag) - InStr(lvAccounts.ColumnHeaders(X).Tag, "^"))
                        Case "No Format"
                            If X = 1 Then
                                itmX.Text = MySQL.OP(rsPlanTemp, Left(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") - 1))
                                If InStr(itmX.Text, "-1") > 0 Then itmX.Text = sSTR.ReplaceString(itmX.Text, "-1", "Unlimited")
                            Else
                                
                                itmX.SubItems(X - 1) = MySQL.OP(rsPlanTemp, Left(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") - 1))
                                If InStr(itmX.SubItems(X - 1), "-1") > 0 Then itmX.SubItems(X - 1) = sSTR.ReplaceString(itmX.SubItems(X - 1), "-1", "Unlimited")
                            End If
                        Case Else
                            
                                If InStr(LCase(Mid(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") + 1)), "select") > 0 Then
                                    sResult = MySQL.fldType(rsPlanTemp(Left(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") - 1)).Type, bNumeric)
                                    Select Case bNumeric
                                    Case True
                                        If Not IsNull(rsPlanTemp(Left(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") - 1))) Then
                                            SQL = LCase(Mid(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") + 1)) & "'" & Val(rsPlanTemp(Left(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") - 1))) & "'"
                                        Else
                                            SQL = LCase(Mid(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") + 1)) & "'0'"
                                        End If
                                    Case False
                                        If Not IsNull(rsPlanTemp(Left(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") - 1))) Then
                                            SQL = LCase(Mid(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") + 1)) & "'" & MySQL.ESC(rsPlanTemp(Left(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") - 1))) & "'"
                                        Else
                                            SQL = LCase(Mid(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") + 1)) & "''"
                                        End If
                                    End Select
                                    Call MySQL.OpenTable(ADOConn, rsload, , SQL)
                                    If rsload.RecordCount > 0 Then
                                       sResult = MySQL.fldType(rsload("nResult").Type, bNumeric)
                                       Select Case bNumeric
                                       Case True
                                            
                                            If X = 1 Then
                                                itmX.Text = "" & Val(IIf(IsNull(rsload("nResult")), 0, rsload("nResult")))
                                            Else
                                                itmX.SubItems(X - 1) = "" & Val(IIf(IsNull(rsload("nResult")), 0, rsload("nResult")))
                                            End If
                                            
                                       Case False
    
                                            If X = 1 Then
                                                itmX.Text = "" & IIf(IsNull(rsload("nResult")), 0, rsload("nResult"))
                                            Else
                                                itmX.SubItems(X - 1) = "" & IIf(IsNull(rsload("nResult")), "", rsload("nResult"))
                                            End If
                                                                               End Select
                                       'itmX.SubItems(X-1) = Format(MySQL.OP(rsPlanTemp, Left(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") - 1)), Mid(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") + 1, Len(lvAccounts.ColumnHeaders(X).Tag) - InStr(lvAccounts.ColumnHeaders(X).Tag, "^")))
                                    End If
                                Else
                                    
                                    If X = 1 Then
                                        itmX.Text = "" & Format(MySQL.OP(rsPlanTemp, Left(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") - 1)), Mid(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") + 1, Len(lvAccounts.ColumnHeaders(X).Tag) - InStr(lvAccounts.ColumnHeaders(X).Tag, "^")))
                                        If InStr(itmX.Text, "-1") > 0 Then itmX.Text = sSTR.ReplaceString(itmX.Text, "-1", "Unlimited")
                                    Else
                                        itmX.SubItems(X - 1) = "" & Format(MySQL.OP(rsPlanTemp, Left(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") - 1)), Mid(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") + 1, Len(lvAccounts.ColumnHeaders(X).Tag) - InStr(lvAccounts.ColumnHeaders(X).Tag, "^")))
                                        If InStr(itmX.SubItems(X - 1), "-1") > 0 Then itmX.SubItems(X - 1) = sSTR.ReplaceString(itmX.SubItems(X - 1), "-1", "Unlimited")
                                    End If
                                            
                                End If
                            
                        End Select
                    End If
                End If
                gSleep
            Next X
            rsPlanTemp.MoveNext
            If Err.Number <> 0 Then Exit Sub
        Wend
        
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
