VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmPools 
   Caption         =   "Radius Pools"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9660
   ControlBox      =   0   'False
   Icon            =   "frmPools_new.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1230
      TabIndex        =   3
      Top             =   4950
      Width           =   945
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   315
      Left            =   2190
      TabIndex        =   4
      Top             =   4950
      Width           =   1065
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Server"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   4950
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3945
      Left            =   3420
      ScaleHeight     =   3945
      ScaleWidth      =   6045
      TabIndex        =   20
      Top             =   1260
      Width           =   6045
      Begin VB.TextBox txtPoolDesc 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1620
         TabIndex        =   6
         Tag             =   "0"
         Top             =   120
         Width           =   4365
      End
      Begin VB.Frame Frame4 
         Caption         =   "Radius User List"
         Height          =   1635
         Left            =   90
         TabIndex        =   27
         Top             =   2250
         Width           =   5925
         Begin VB.TextBox txtRadiusField 
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   11
            Left            =   2070
            TabIndex        =   13
            Top             =   240
            Width           =   2805
         End
         Begin VB.TextBox txtRadiusField 
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   10
            Left            =   2070
            TabIndex        =   15
            Text            =   "/"
            Top             =   570
            Width           =   3705
         End
         Begin VB.TextBox txtRadiusField 
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   9
            Left            =   2070
            TabIndex        =   16
            Top             =   900
            Width           =   3705
         End
         Begin VB.TextBox txtRadiusField 
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   8
            Left            =   2070
            TabIndex        =   17
            Top             =   1230
            Width           =   1425
         End
         Begin VB.TextBox txtRadiusField 
            BorderStyle     =   0  'None
            Height          =   270
            IMEMode         =   3  'DISABLE
            Index           =   7
            Left            =   4560
            PasswordChar    =   "#"
            TabIndex        =   18
            Top             =   1230
            Width           =   1215
         End
         Begin VB.TextBox txtRadiusField 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   6
            Left            =   4920
            TabIndex        =   14
            Text            =   "21"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "FTP Server:"
            Height          =   240
            Index           =   0
            Left            =   570
            TabIndex        =   32
            Top             =   240
            Width           =   1350
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Target Directory:"
            Height          =   240
            Left            =   570
            TabIndex        =   31
            Top             =   570
            Width           =   1350
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Upload Filename:"
            Height          =   240
            Left            =   570
            TabIndex        =   30
            Top             =   900
            Width           =   1350
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Username:"
            Height          =   240
            Left            =   570
            TabIndex        =   29
            Top             =   1260
            Width           =   1350
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Password:"
            Height          =   240
            Left            =   3660
            TabIndex        =   28
            Top             =   1260
            Width           =   765
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Radius Log File"
         Height          =   1725
         Left            =   60
         TabIndex        =   21
         Top             =   510
         Width           =   5955
         Begin VB.TextBox txtRadiusField 
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   0
            Left            =   2100
            TabIndex        =   7
            Top             =   270
            Width           =   2805
         End
         Begin VB.TextBox txtRadiusField 
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   1
            Left            =   2100
            TabIndex        =   9
            Text            =   "/"
            Top             =   600
            Width           =   3705
         End
         Begin VB.TextBox txtRadiusField 
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   2
            Left            =   2100
            TabIndex        =   10
            Top             =   930
            Width           =   3705
         End
         Begin VB.TextBox txtRadiusField 
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   3
            Left            =   2100
            TabIndex        =   11
            Top             =   1260
            Width           =   1425
         End
         Begin VB.TextBox txtRadiusField 
            BorderStyle     =   0  'None
            Height          =   270
            IMEMode         =   3  'DISABLE
            Index           =   4
            Left            =   4590
            PasswordChar    =   "#"
            TabIndex        =   12
            Top             =   1260
            Width           =   1215
         End
         Begin VB.TextBox txtRadiusField 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   5
            Left            =   4950
            TabIndex        =   8
            Text            =   "21"
            Top             =   270
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "FTP Server:"
            Height          =   240
            Index           =   1
            Left            =   600
            TabIndex        =   26
            Top             =   270
            Width           =   1350
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Target Directory:"
            Height          =   240
            Left            =   600
            TabIndex        =   25
            Top             =   600
            Width           =   1350
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Target Filename:"
            Height          =   240
            Left            =   600
            TabIndex        =   24
            Top             =   930
            Width           =   1350
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Username:"
            Height          =   240
            Left            =   600
            TabIndex        =   23
            Top             =   1290
            Width           =   1350
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Password:"
            Height          =   240
            Left            =   3690
            TabIndex        =   22
            Top             =   1290
            Width           =   765
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Pool Description:"
         Height          =   240
         Left            =   90
         TabIndex        =   33
         Top             =   150
         Width           =   1275
      End
   End
   Begin MSComctlLib.TabStrip ts 
      Height          =   4395
      Left            =   3360
      TabIndex        =   19
      Top             =   870
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   7752
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Settings"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Servers"
      Height          =   4095
      Left            =   90
      TabIndex        =   5
      Top             =   750
      Width           =   3165
      Begin MSComctlLib.ListView lvServers 
         Height          =   3705
         Left            =   120
         TabIndex        =   1
         Top             =   270
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   6535
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         Sorted          =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Servers"
            Object.Width           =   4762
         EndProperty
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Radius Servers"
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
      Index           =   0
      Left            =   1020
      TabIndex        =   0
      Top             =   60
      Width           =   1635
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   30
      Picture         =   "frmPools_new.frx":0442
      Stretch         =   -1  'True
      Top             =   -30
      Width           =   930
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   11820
      Y1              =   420
      Y2              =   420
   End
End
Attribute VB_Name = "frmPools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdAdd_Click"
    Const ContainerName = "frmPools"
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

 
    txtPoolDesc.Tag = 0
    txtPoolDesc.Text = ""
    txtRadiusField(0).Text = ""
    txtRadiusField(1).Text = ""
    txtRadiusField(2).Text = ""
    txtRadiusField(3).Text = ""
    txtRadiusField(4).Text = ""
    txtRadiusField(5).Text = ""
    txtRadiusField(6).Text = ""
    txtRadiusField(7).Text = ""
    txtRadiusField(8).Text = ""
    txtRadiusField(9).Text = ""
    txtRadiusField(10).Text = ""
    txtRadiusField(11).Text = ""
    
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

Private Sub cmdClose_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdClose_Click"
    Const ContainerName = "frmPools"
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

Private Sub cmdSave_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdSave_Click"
    Const ContainerName = "frmPools"
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
    Select Case txtPoolDesc.Tag
    Case 0
        bResult = MySQL.OpenTable(directConn, rsSave, , "select * from radiuspools lIMIT 1")
        rsSave.AddNew
    Case Else
        bResult = MySQL.OpenTable(directConn, rsSave, , "select * from radiuspools where RecID = " & txtPoolDesc.Tag)
    End Select
    
    rsSave!Description = txtPoolDesc.Text
    rsSave!rlFTPServer = txtRadiusField(0).Text
    rsSave!rlTagetDir = txtRadiusField(1).Text
    rsSave!rlFilename = txtRadiusField(2).Text
    rsSave!rlUsername = txtRadiusField(3).Text
    'rsSave!rlPassword = txtRadiusField(4).Text
    rsSave!rlPort = txtRadiusField(5).Text
    rsSave!ufPort = txtRadiusField(6).Text
    'rsSave!ufPassword = txtRadiusField(7).Text
    rsSave!ufUsername = txtRadiusField(8).Text
    rsSave!ufFilename = txtRadiusField(9).Text
    rsSave!ufTagetDir = txtRadiusField(10).Text
    rsSave!ufFTPServer = txtRadiusField(11).Text
    
    Dim itmX As ListItem
    Select Case txtPoolDesc.Tag
    Case 0
        txtPoolDesc.Tag = MySQL.SetRecID(rsSave, "radiuspools", directConn)
        Set itmX = lvServers.ListItems.Add(, "k" & txtPoolDesc.Tag, rsSave!Description)
    Case Else
        lvServers.SelectedItem.Text = rsSave!Description
    End Select

    rsSave.Update
    
    MySQL.Execute directConn, "UPDATE radiuspools SET ufPassword=AES_ENCRYPT('" + txtRadiusField(7).Text + "','" & odb.colSalts.ReturnSalt("md5Password") & "'), rlPassword=AES_ENCRYPT('" + txtRadiusField(4).Text + "','" & odb.colSalts.ReturnSalt("md5Password") & "') Where RecID = " & txtPoolDesc.Tag
    
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
    Const ContainerName = "frmPools"
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


    LoadServers
    
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

Public Sub LoadServers()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "LoadServers"
    Const ContainerName = "frmPools"
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
    Dim bResult As Boolean
    Dim itmX As ListItem
    
    bResult = MySQL.OpenTable(directConn, rsload, , "select * from radiuspools")
    
    If rsload.RecordCount > 0 Then
        Do
            Set itmX = lvServers.ListItems.Add(, "k" & rsload!RecID, rsload!Description)
            rsload.MoveNext
        Loop Until rsload.EOF Or Err.Number <> 0
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

Private Sub lvServers_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvServers_ItemClick"
    Const ContainerName = "frmPools"
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
    Dim bResult As Boolean

    bResult = MySQL.OpenTable(directConn, rsload, , "select Description, rlFTPServer, rlTagetDir, rlFilename, rlUsername, AES_DECRYPT(rlPassword,'" & odb.colSalts.ReturnSalt("md5Password") & "') as rlPassword, rlPort, ufFTPServer, ufTagetDir, ufFilename, ufUsername, AES_DECRYPT(ufPassword,'" & odb.colSalts.ReturnSalt("md5Password") & "') as ufPassword, ufPort from radiuspools where RecID = " & Mid(Item.Key, 2))
    
    txtPoolDesc.Tag = Mid(Item.Key, 2)
    txtPoolDesc.Text = rsload!Description
    txtRadiusField(0).Text = IIf(IsNull(rsload!rlFTPServer), "", rsload!rlFTPServer) '
    txtRadiusField(1).Text = IIf(IsNull(rsload!rlTagetDir), "", rsload!rlTagetDir) 'rsLoad!rlTagetDir
    txtRadiusField(2).Text = IIf(IsNull(rsload!rlFilename), "", rsload!rlFilename) 'rsLoad!rlFilename
    txtRadiusField(3).Text = IIf(IsNull(rsload!rlUsername), "", rsload!rlUsername) '
    txtRadiusField(4).Text = IIf(IsNull(rsload!rlPassword), "", rsload!rlPassword) '
    txtRadiusField(5).Text = IIf(IsNull(rsload!rlPort), "", rsload!rlPort) '
    txtRadiusField(6).Text = IIf(IsNull(rsload!ufPort), "", rsload!ufPort) '
    txtRadiusField(7).Text = IIf(IsNull(rsload!ufPassword), "", rsload!ufPassword) '
    txtRadiusField(8).Text = IIf(IsNull(rsload!ufUsername), "", rsload!ufUsername) '
    txtRadiusField(9).Text = IIf(IsNull(rsload!ufFilename), "", rsload!ufFilename) '
    txtRadiusField(10).Text = IIf(IsNull(rsload!ufTagetDir), "", rsload!ufTagetDir) '
    txtRadiusField(11).Text = IIf(IsNull(rsload!ufFTPServer), "", rsload!ufFTPServer) '

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
