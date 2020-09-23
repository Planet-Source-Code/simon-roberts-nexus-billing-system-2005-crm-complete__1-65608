VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFTPAccount 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FTP Account"
   ClientHeight    =   7620
   ClientLeft      =   5580
   ClientTop       =   4755
   ClientWidth     =   4875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   3660
      TabIndex        =   5
      Top             =   7230
      Width           =   1125
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   315
      Left            =   150
      TabIndex        =   4
      Top             =   7230
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      Height          =   5985
      Left            =   150
      TabIndex        =   6
      Top             =   1140
      Width           =   4635
      Begin VB.Frame Frame2 
         Caption         =   "Upload Speed"
         Height          =   1335
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   4410
         Width           =   4395
         Begin VB.OptionButton optBandwidth 
            Caption         =   "2048KBit"
            Height          =   285
            Index           =   11
            Left            =   2220
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   900
            Width           =   2000
         End
         Begin VB.OptionButton optBandwidth 
            Caption         =   "1024Kbit"
            Height          =   285
            Index           =   10
            Left            =   210
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   900
            Value           =   -1  'True
            Width           =   2000
         End
         Begin VB.OptionButton optBandwidth 
            Caption         =   "512Kbit"
            Height          =   285
            Index           =   9
            Left            =   2220
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   600
            Width           =   2000
         End
         Begin VB.OptionButton optBandwidth 
            Caption         =   "256Kbit"
            Height          =   285
            Index           =   8
            Left            =   210
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   600
            Width           =   2000
         End
         Begin VB.OptionButton optBandwidth 
            Caption         =   "128Kbit"
            Height          =   285
            Index           =   7
            Left            =   2220
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   300
            Width           =   2000
         End
         Begin VB.OptionButton optBandwidth 
            Caption         =   "64Kbit"
            Height          =   285
            Index           =   6
            Left            =   210
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   300
            Width           =   2000
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Download Speed"
         Height          =   1335
         Index           =   0
         Left            =   150
         TabIndex        =   15
         Top             =   2850
         Width           =   4395
         Begin VB.OptionButton optBandwidth 
            Caption         =   "64Kbit"
            Height          =   285
            Index           =   0
            Left            =   210
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   300
            Width           =   2000
         End
         Begin VB.OptionButton optBandwidth 
            Caption         =   "128Kbit"
            Height          =   285
            Index           =   1
            Left            =   2220
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   300
            Width           =   2000
         End
         Begin VB.OptionButton optBandwidth 
            Caption         =   "256Kbit"
            Height          =   285
            Index           =   2
            Left            =   210
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   600
            Width           =   2000
         End
         Begin VB.OptionButton optBandwidth 
            Caption         =   "512Kbit"
            Height          =   285
            Index           =   3
            Left            =   2220
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   600
            Width           =   2000
         End
         Begin VB.OptionButton optBandwidth 
            Caption         =   "1024Kbit"
            Height          =   285
            Index           =   4
            Left            =   210
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   900
            Value           =   -1  'True
            Width           =   2000
         End
         Begin VB.OptionButton optBandwidth 
            Caption         =   "2048KBit"
            Height          =   285
            Index           =   5
            Left            =   2220
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   900
            Width           =   2000
         End
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   270
         Left            =   2115
         TabIndex        =   14
         Top             =   2250
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   476
         _Version        =   393216
         Value           =   4
         BuddyControl    =   "txtField(4)"
         BuddyDispid     =   196614
         BuddyIndex      =   4
         OrigLeft        =   2370
         OrigTop         =   1860
         OrigRight       =   2610
         OrigBottom      =   2175
         Max             =   30
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   4
         Left            =   120
         MaxLength       =   50
         TabIndex        =   12
         Text            =   "4"
         Top             =   2250
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   3
         Left            =   120
         MaxLength       =   255
         TabIndex        =   3
         Top             =   1590
         Width           =   4395
      End
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   120
         MaxLength       =   100
         TabIndex        =   0
         Top             =   270
         Width           =   4395
      End
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   120
         MaxLength       =   50
         TabIndex        =   1
         Top             =   900
         Width           =   2235
      End
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2370
         MaxLength       =   50
         TabIndex        =   2
         Top             =   900
         Width           =   2145
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Sessions"
         Height          =   225
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   2520
         Width           =   2235
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Default Directory"
         Height          =   225
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1890
         Width           =   4395
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Contact Name"
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   570
         Width           =   4395
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Username"
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   2235
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Password"
         Height          =   225
         Index           =   2
         Left            =   2370
         TabIndex        =   7
         Top             =   1200
         Width           =   2145
      End
   End
   Begin VB.Image Image1 
      Height          =   1080
      Left            =   30
      Picture         =   "frmFTPAccount_new.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FTP Account"
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
      Left            =   990
      TabIndex        =   10
      Top             =   150
      Width           =   1350
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   4980
      X2              =   60
      Y1              =   510
      Y2              =   510
   End
End
Attribute VB_Name = "frmFTPAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sSessions As String
Public byteBandwidth As String
Public byteBWUpload As String
Public sUsername As String
Public sPassword As String
Public sContactName As String
Public sBaseDIR As String
Public lRecID As Long

Public iCloseState As frm_CloseStates


Private Sub cmdSave_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdSave_Click"
    Const ContainerName = "frmFTPAccount"
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
    Dim sSQL As String
    
    If InStr(txtField(3).Text, "/home/" & Login.sVISPDomain & "/") = 0 Then
        If Login.bMaster = False Then
            MsgBox "Directory Path is not valid", vbCritical, "Invalide Default Directory"
            txtField(3).Text = "/home/" & Login.sVISPDomain & "/" & txtField(1).Text
            Exit Sub
        End If
    End If
    
    If lRecID <> 0 Then
        bResult = MySQL.OpenTable(directConn, rsload, , "select * from acci_services where RecID = " & lRecID)
        If rsload!Username <> txtField(1) Then bCheckUsername = True
    End If
    
    If lRecID = 0 Then bCheckUsername = True
    
    Select Case bCheckUsername
    Case True
        
        bResult = MySQL.OpenTable(directConn, rsload, , "select plantypes.RecID from plantypes, servicetypes where servicetypes.ServiceKey = 'FTP' And plantypes.ServiceID = servicetypes.RecID")
        If rsload.RecordCount > 0 Then
            
            Do
                If Len(SQL) > 0 Then sSQL = sSQL + " OR "
                sSQL = sSQL + "ptRecID = " & rsload!RecID
                rsload.MoveNext
            Loop Until rsload.EOF Or Err.Number <> 0
            
            bResult = MySQL.OpenTable(directConn, rsload, , "select * from acci_services Where (" & sSQL & ") and Username Like '" & MySQL.ESC(txtField(1)) & "'")
            
            If rsload.RecordCount > 0 Then
                MsgBox "A username exist in the system for that service already", vbCritical, "Username Exists"
                Exit Sub
            End If
        End If
    
    End Select
    
    iCloseState = frmCloseSave
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

Private Sub cmdCancel_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdCancel_Click"
    Const ContainerName = "frmFTPAccount"
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


    iCloseState = frmCloseCancel
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

Private Sub txtfield_Change(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtfield_Change"
    Const ContainerName = "frmFTPAccount"
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


    Select Case Index
    Case 1
        txtField(3).Text = "/home/" & Login.sVISPDomain & "/" & txtField(1).Text
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

Private Sub txtField_GotFocus(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtfield_GotFocus"
    Const ContainerName = "frmFTPAccount"
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


    txtField(Index).SelStart = 0
    txtField(Index).SelLength = Len(txtField(Index).Text)
    
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

Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtfield_KeyPress"
    Const ContainerName = "frmFTPAccount"
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


    Select Case Index
    Case 1
        Select Case KeyAscii
        Case Asc(" ")
            KeyAscii = 0
        End Select
    Case 3
        Select Case KeyAscii
        Case Asc("\")
            KeyAscii = Asc("/")
        End Select
    Case 4
        Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
            KeyAscii = 0
            SendKeys "{TAB}"
        Case Else
            KeyAscii = 0
        End Select
    Case Else
        Select Case KeyAscii
        Case 13
            KeyAscii = 0
            SendKeys "{TAB}"
        End Select
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
Private Sub Form_Load()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Load"
    Const ContainerName = "frmFTPAccount"
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


    txtField(0) = sContactName
    txtField(1) = sUsername
    txtField(2) = sPassword
    txtField(3) = sBaseDIR
    txtField(4) = IIf(sSessions = "", "4", sSessions)
    
    Dim bx As Byte
    
    For bx = 0 To 5
        If optBandwidth(bx).Caption = byteBandwidth Then
            optBandwidth(bx).Value = True
            Exit For
        End If
    Next
    
    For bx = 6 To 11
        If optBandwidth(bx).Caption = byteBWUpload Then
            optBandwidth(bx).Value = True
            Exit For
        End If
    Next
    
            
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
    Const ContainerName = "frmFTPAccount"
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



    sContactName = txtField(0)
    sUsername = txtField(1)
    sPassword = txtField(2)
    sBaseDIR = txtField(3)
    sSessions = txtField(4)
    
    Dim bx As Byte
    
    For bx = 0 To 5
        If optBandwidth(bx).Value = True Then
            byteBandwidth = optBandwidth(bx).Caption
            Exit For
        End If
    Next
    
    For bx = 6 To 11
        If optBandwidth(bx).Value = True Then
            byteBWUpload = optBandwidth(bx).Caption
            Exit For
        End If
    Next
    
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
