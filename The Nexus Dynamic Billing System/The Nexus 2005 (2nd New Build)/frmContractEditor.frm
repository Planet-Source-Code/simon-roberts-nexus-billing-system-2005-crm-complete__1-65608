VERSION 5.00
Begin VB.Form frmContractEditor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Change Contract Info"
   ClientHeight    =   1995
   ClientLeft      =   6060
   ClientTop       =   3660
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame7 
      BackColor       =   &H00A3A3FE&
      Caption         =   "Contract Description"
      Height          =   1965
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.TextBox txtField 
         BackColor       =   &H00E0DFFF&
         Height          =   285
         Index           =   0
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   7
         Tag             =   "Description"
         Top             =   240
         Width           =   4275
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
         TabIndex        =   6
         Tag             =   "Termination"
         Top             =   600
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
         Index           =   7
         Left            =   3990
         TabIndex        =   5
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
         Index           =   8
         Left            =   1260
         TabIndex        =   4
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
         Index           =   9
         Left            =   1260
         TabIndex        =   3
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
         Index           =   10
         Left            =   1260
         TabIndex        =   2
         Tag             =   "FeePerHour"
         Top             =   1500
         Width           =   1395
      End
      Begin VB.CommandButton cmdContract 
         BackColor       =   &H00A3A3FE&
         Caption         =   "Update Table"
         Height          =   345
         Index           =   0
         Left            =   4050
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1530
         Width           =   1515
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   13
         Top             =   300
         Width           =   840
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
         Caption         =   "Billing Cycle Fee:"
         Height          =   195
         Index           =   3
         Left            =   2730
         TabIndex        =   11
         Top             =   1140
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Joining Fee:"
         Height          =   195
         Index           =   4
         Left            =   330
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Extra Per MB:"
         Height          =   195
         Index           =   5
         Left            =   210
         TabIndex        =   9
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Extra Per Hour:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmContractEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RecID As Variant
Public Parnt As frmAccountTypes


Private Sub cmdContract_Click(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdContract_Click"
    Const ContainerName = "frmContractEditor"
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


    If Not Parnt.lvContracts.SelectedItem Is Nothing Then
        Parnt.lvContracts.SelectedItem.SubItems(3) = txtFee(6)
        Parnt.lvContracts.SelectedItem.SubItems(5) = txtFee(7)
        Parnt.lvContracts.SelectedItem.SubItems(4) = txtFee(8)
        Parnt.lvContracts.SelectedItem.SubItems(6) = txtFee(9)
        Parnt.lvContracts.SelectedItem.SubItems(7) = txtFee(10)
        'lvContracts.selectedItem.SubItems(8) = lvContracts.Tag
        Call MySQL.Execute(directConn, "update contractsruntime set Termination = '" & txtFee(6) & "', PeriodFee = '" & txtFee(7) & "', JoiningFee = '" & txtFee(8) & "', FeePerBlock = '" & txtFee(9) & "', FeePerHour = '" & txtFee(10) & "' where RecID = " & RecID)
        Unload Me
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

Private Sub txtFee_DblClick(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtFee_DblClick"
    Const ContainerName = "frmContractEditor"
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


    frmGSTCalc.Show 1
    txtFee(Index) = "" & frmGSTCalc.cAmount
    
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

Private Sub txtFee_KeyPress(Index As Integer, KeyAscii As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtFee_KeyPress"
    Const ContainerName = "frmContractEditor"
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



    Select Case KeyAscii
    Case 8
    Case 48 To 57
    Case Asc(".")
        If InStr(txtFee(Index), ".") > 0 Then KeyAscii = 0
    Case Else
        KeyAscii = 0
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

