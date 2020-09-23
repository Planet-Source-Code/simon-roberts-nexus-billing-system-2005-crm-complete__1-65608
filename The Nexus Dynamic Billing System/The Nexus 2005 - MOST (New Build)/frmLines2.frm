VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLines 
   Caption         =   "Line Usage and History"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   9075
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   1530
      TabIndex        =   9
      Top             =   180
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   345
      Left            =   60
      TabIndex        =   8
      Top             =   150
      Width           =   1425
   End
   Begin VB.PictureBox picTSContainer 
      BorderStyle     =   0  'None
      Height          =   5415
      Index           =   1
      Left            =   90
      ScaleHeight     =   5415
      ScaleWidth      =   8835
      TabIndex        =   6
      Top             =   1080
      Width           =   8835
      Begin MSComctlLib.ListView lvPools 
         Height          =   5985
         Left            =   30
         TabIndex        =   7
         Top             =   60
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   10557
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Pool Name"
            Object.Width           =   3526
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Number of Connections"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   1000
      Left            =   1950
      Top             =   90
   End
   Begin VB.TextBox txtRefreshMin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   7710
      TabIndex        =   4
      Text            =   "10"
      Top             =   60
      Width           =   1005
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   450
      Left            =   8716
      TabIndex        =   3
      Top             =   60
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   794
      _Version        =   393216
      Value           =   10
      BuddyControl    =   "txtRefreshMin"
      BuddyDispid     =   196612
      OrigLeft        =   8730
      OrigTop         =   120
      OrigRight       =   8970
      OrigBottom      =   495
      Max             =   20
      Min             =   1
      SyncBuddy       =   -1  'True
      Wrap            =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.PictureBox picTSContainer 
      BorderStyle     =   0  'None
      Height          =   5415
      Index           =   0
      Left            =   120
      ScaleHeight     =   5415
      ScaleWidth      =   8835
      TabIndex        =   1
      Top             =   1050
      Width           =   8835
      Begin MSComctlLib.ListView lvLine 
         Height          =   5985
         Left            =   30
         TabIndex        =   2
         Top             =   60
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   10557
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Username"
            Object.Width           =   3526
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Upload"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Download"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Logged On"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "How Long"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Radius Pool"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin MSComctlLib.TabStrip ts 
      Height          =   5925
      Left            =   60
      TabIndex        =   0
      Top             =   660
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   10451
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Line Usage"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Radius Pools"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Number of Minutes to Refresh:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3780
      TabIndex        =   5
      Top             =   120
      Width           =   3825
   End
End
Attribute VB_Name = "frmLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oConn As adodb.Connection


'Dim wrkODBC As DAO.Workspace

Dim daoConnected As Boolean

Public lCount As Long


Private Sub cmdRefresh_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdRefresh_Click"
    Const ContainerName = "frmLines"
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


    lCount = 0
    
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
    Const ContainerName = "frmLines"
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


    Dim rsLog As adodb.Recordset
    Dim rsAlive As adodb.Recordset
    Dim itmX As ListItem
    
    'If daoConnected = False Then
    '    Set wrkODBC = CreateWorkspace("The Nexus_linerefresh", UID, PWD, dbUseODBC)
   '
    '    Call GUI.LoadColWidths(lvLine, Me)
   '
   '     Call MySQL.DAOConn("Radius", sServer, sUID, sPWD, oDAOConn, wrkODBC)
   ' End If
    
    Exit Sub
    
    
    Call MySQL.Connection("Radius", sServer, sUID, sPWD, oConn)
    
    
    
    If Login.bTestBench = True Then Exit Sub
    
    br = MySQL.OpenTable(oConn, rsLog, , "select RadAcctID as RecID, User-Name, SessionStart from RadiusLogTwo Where `Acct-Status-Type` = 'Start'")
    rsLog.CursorLocation = adUseServer
    
    lvLine.ListItems.Clear
    If br = True Then
        If rsLog.RecordCount > 0 Then
            While Not rsLog.EOF And Err.Number = 0
                Set itmX = lvLine.ListItems.Add(, , rsLog("User-Name"))
                itmX.SubItems(3) = Format(rsLog!SessionStart, "dd-mm-yyyy Hh:Nn:Ss")
                
                br = MySQL.OpenTable(oConn, rsAlive, , "select SessionStart, timestamp, `Acct-Input-Octets` , `Acct-Output-Octets` from RadiusLogTwo Where `Acct-Status-Type` Like 'Alive' and `User-Name` = '" & rsLog("User-Name") & "'")
                If rsAlive.RecordCount > 0 Then
                    itmX.SubItems(1) = Oct(IIf(IsNull(rsAlive("Acct-Input-Octets")), 0, rsAlive("Acct-Input-Octets"))) & " bytes"
                    itmX.SubItems(2) = Oct(IIf(IsNull(rsAlive("Acct-Output-Octets")), 0, rsAlive("Acct-Output-Octets"))) & " bytes"
                    itmX.SubItems(4) = DateDiff("n", rsLog!SessionStart, rsAlive!SessionStart) & " mins"
                End If
                
                rsLog.MoveNext
            Wend
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

Private Sub Form_Resize()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Resize"
    Const ContainerName = "frmLines"
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


    If Me.ScaleHeight < 800 Or Me.ScaleWidth < 1500 Then Exit Sub
    If Me.WindowState <> vbMinimized Then
        txtRefreshMin.Move Me.ScaleWidth - txtRefreshMin.Width - UpDown1.Width - 60
        UpDown1.Move txtRefreshMin.Left + txtRefreshMin.Width
        Label1.Move txtRefreshMin.Left - Label1.Width - 60
        ts.Move 60, ts.Top, Me.ScaleWidth - 120, Me.ScaleHeight - 120 - ts.Top
        Call ts_Click
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

Private Sub mnuMinimise_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuMinimise_Click"
    Const ContainerName = "frmLines"
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


    Me.WindowState = vbMinimized
        
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

Private Sub Form_Unload(Cancel As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Unload"
    Const ContainerName = "frmLines"
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


    Call GUI.SaveColWidths(lvLine, Me)
    'Cancel = True
    
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

Private Sub lvLine_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvLine_ColumnClick"
    Const ContainerName = "frmLines"
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


    Call GUI.ColumnSort(ColumnHeader, lvLine)
    
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

Private Sub picTSContainer_Resize(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "picTSContainer_Resize"
    Const ContainerName = "frmLines"
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


    If picTSContainer(Index).ScaleHeight < 120 Or picTSContainer(Index).ScaleWidth < 1500 Then Exit Sub
    
    Select Case Index
    Case 0
        lvLine.Move 60, 60, picTSContainer(Index).ScaleWidth - 120, picTSContainer(Index).ScaleHeight - 120
    Case 1
        lvPools.Move 60, 60, picTSContainer(Index).ScaleWidth - 120, picTSContainer(Index).ScaleHeight - 120
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

Private Sub tmrRefresh_Timer()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "tmrRefresh_Timer"
    Const ContainerName = "frmLines"
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


    
    'Exit Sub
    
    If bDebug = True Then On Error GoTo 0 Else On Error GoTo ErrorOccur
       
    lCount = lCount + 1
    
    Dim rsLog As adodb.Recordset
    'Dim rsAlive As DAO.Recordset
    Dim rsPlans As adodb.Recordset
    Dim rsPools As adodb.Recordset
    Dim rsRadiusAcc As adodb.Recordset
    
    
    
    Dim itmX As ListItem
    
    If lCount / 60 >= Val(txtRefreshMin.Text) Then
        lCount = 0
    Else
        Me.Caption = "Line Usgage and History (" & Val(txtRefreshMin.Text) * 60 - lCount & " seconds till refresh)"
    End If
    
    If lCount = 1 Then
        lvLine.ListItems.Clear
        lvPools.ListItems.Clear

        'If Login.bTestBench = True Then Exit Sub
        
        Screen.MousePointer = vbArrowHourglass
        Me.Caption = "(Connecting to MySQL Server)"
            
        If daoConnected = False Then
        
            'wrkODBC.Close
            'Set wrkODBC = CreateWorkspace("The Nexus_linerefresh", UID, PWD, dbUseODBC)
            'Call MySQL.DAOConn("Radius", sServer, sUID, sPWD, oDAOConn, wrkODBC)
            
            daoConnected = True
        End If
        
        frmMDIMain.Caption = "(Requerying)"
        gSleep
redoagain:
        Me.Caption = "(Requerying Radius Database)"
        'br = MySQL.OpenTable(directConn, rsLog, , "select `User-Name`, SessionStart, timestamp, `Acct-Input-Octets` , `Acct-Output-Octets` from radius.RadiusLogTwo Where `Acct-Status-Type` Like 'Start'")
        gSleep
        bResult = MySQL.OpenTable(directConn, rsPools, , "select * from radiuspools")
        gSleep
        bResult = MySQL.OpenTable(directConn, rsPlans, , "select * from plantypes")
        gSleep
        bResult = MySQL.OpenTable(directConn, rsRadiusAcc, , "select radiusaccounts.* from projectalpha.radiusaccounts, projectalpha.accountinfo, projectalpha.Flags, projectalpha.plantypes Where projectalpha.radiusaccounts.acci_RecID = projectalpha.accountinfo.RecID and projectalpha.accountinfo.Cancelled = 0")
        gSleep
        
        While Not rsPools.EOF And Err.Number = 0
            Set itmX = lvPools.ListItems.Add(, "k" & rsPools!Description, rsPools!Description)
            rsPools.MoveNext
        Wend
        
        GoTo runAliveTest
        
        If br = True Then
            If Not rsLog Is Nothing Then
            If rsLog.RecordCount > 0 Then
                While Not rsLog.EOF And Err.Number = 0
                    Set itmX = lvLine.ListItems.Add(, , rsLog("User-Name"))
                    itmX.SubItems(3) = Format(rsLog!SessionStart, "dd-mm-yyyy Hh:Nn:Ss")
                    gSleep
                    'br = MySQL.OpenTable(oDAOConn, rsAlive, , "select SessionStart, timestamp, `Acct-Input-Octets` , `Acct-Output-Octets` from RadiusLogTwo Where `Acct-Status-Type` = 'Alive' and `User-Name` = '" & rsLog("User-Name") & "'")
                    gSleep
                    'If rsAlive.RecordCount > 0 Then
                        itmX.SubItems(1) = Oct(IIf(IsNull(Oct(rsLog("Acct-Output-Octets"))), 0, Oct(rsLog("Acct-Output-Octets")))) & " bytes"
                        itmX.SubItems(2) = Oct(IIf(IsNull(Oct(rsLog("Acct-Input-Octets"))), 0, Oct(rsLog("Acct-Input-Octets")))) & " bytes"
                        itmX.SubItems(4) = DateDiff("n", rsLog!SessionStart, rsLog!TimeStamp) & " mins"
                    'End If
                    
                    If rsLog("User-Name") <> "" Then
                        On Error Resume Next
                        rsRadiusAcc.Filter = "Username Like '" & rsLog("User-Name") & "'"
                        gSleep
                        rsPlans.Filter = "RecID = " & rsRadiusAcc!ptRecID
                        gSleep
                        rsPools.Filter = "RecID = " & rsPlans!RadiusID
                        gSleep
                        itmX.SubItems(5) = rsPools!Description
                        gSleep
                        lvPools.ListItems("k" & rsPools!Description).SubItems(1) = Val(lvPools.ListItems("k" & rsPools!Description).SubItems(1)) + 1
                        gSleep
                    End If
                    rsLog.MoveNext
                Wend
            End If
            End If
        End If
        
runAliveTest:
        On Error GoTo 0
        Dim oMaxs As adodb.Recordset
        
        br = MySQL.OpenTable(directConn, oMaxs, , "select max(`Acct-Input-Octets`) as MaxInput , max(`Acct-Output-Octets`) as MaxOutput from radius.RadiusLogTwo Where `Acct-Status-Type` = 'Alive' or `Acct-Status-Type` = 'Start'")
        br = MySQL.OpenTable(directConn, rsLog, , "select `User-Name`, SessionStart, timestamp, `Acct-Input-Octets`, `Acct-Output-Octets` from radius.RadiusLogTwo Where `Acct-Status-Type` = 'Alive' or `Acct-Status-Type` = 'Start'")
        If br = True Then
            If Not rsLog Is Nothing Then
            If rsLog.RecordCount > 0 Then
                ProgressBar1.Value = 0
                ProgressBar1.Max = rsLog.RecordCount
                
                While Not rsLog.EOF And Err.Number = 0
                    Set itmX = lvLine.ListItems.Add(, , rsLog("User-Name"))
                    itmX.SubItems(3) = Format(rsLog!SessionStart, "dd-mm-yyyy Hh:Nn:Ss")
                    gSleep
                    'br = MySQL.OpenTable(oDAOConn, rsAlive, , "select SessionStart, timestamp, `Acct-Input-Octets` , `Acct-Output-Octets` from RadiusLogTwo Where `Acct-Status-Type` = 'Alive' and `User-Name` = '" & rsLog("User-Name") & "'")
                    gSleep
                    'If rsAlive.RecordCount > 0 Then
                        itmX.SubItems(1) = IIf(IsNull(rsLog("Acct-Output-Octets")), String(Len("" & Oct(oMaxs!MaxOutput)) - Len("" & Oct(rsLog("Acct-Output-Octets"))), "0") & "0", String(Len("" & Oct(oMaxs!MaxOutput)) - Len("" & Oct(rsLog("Acct-Output-Octets"))), "0") + "" & Oct(rsLog("Acct-Output-Octets"))) & " bytes"
                        itmX.SubItems(2) = IIf(IsNull(rsLog("Acct-Input-Octets")), String(Len("" & Oct(oMaxs!MaxInput)) - Len("" & Oct(rsLog("Acct-Input-Octets"))), "0") & "0", String(Len("" & Oct(oMaxs!MaxInput)) - Len("" & Oct(rsLog("Acct-Input-Octets"))), "0") + "" & Oct(rsLog("Acct-Input-Octets"))) & " bytes"
                        itmX.SubItems(4) = DateDiff("n", rsLog!SessionStart, IIf(rsLog!TimeStamp = "" Or IsNull(rsLog!TimeStamp), sysnow, rsLog!TimeStamp)) & " mins"
                    'End If
                    
                    If rsLog("User-Name") <> "" Then
                        On Error Resume Next
                        rsRadiusAcc.Filter = "Username Like '" & rsLog("User-Name") & "'"
                        gSleep
                        rsPlans.Filter = "RecID = " & rsRadiusAcc!ptRecID
                        gSleep
                        rsPools.Filter = "RecID = " & rsPlans!RadiusID
                        gSleep
                        itmX.SubItems(5) = rsPools!Description
                        gSleep
                        lvPools.ListItems("k" & rsPools!Description).SubItems(1) = Val(lvPools.ListItems("k" & rsPools!Description).SubItems(1)) + 1
                        gSleep
                    End If
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    rsLog.MoveNext
                Wend
            End If
            End If
        End If
        
        Screen.MousePointer = vbDefault
        Me.Caption = "Line Usgage and History"
        frmMDIMain.Caption = "The Nexus"
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

Private Sub ts_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "ts_Click"
    Const ContainerName = "frmLines"
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
    
    For X = picTSContainer.LBound To picTSContainer.UBound
        If ts.SelectedItem.Index - 1 <> X Then picTSContainer(X).Visible = False
    Next
    
    If ts.SelectedItem.Index - 1 <= picTSContainer.UBound Then
        picTSContainer(ts.SelectedItem.Index - 1).Move ts.ClientLeft, ts.ClientTop, ts.ClientWidth, ts.ClientHeight
        picTSContainer(ts.SelectedItem.Index - 1).Visible = True
        picTSContainer(ts.SelectedItem.Index - 1).ZOrder 0
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

