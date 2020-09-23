VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmUsername 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Suggested Log In Usernames"
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      Height          =   255
      Left            =   90
      TabIndex        =   2
      Top             =   8910
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3450
      Top             =   3690
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   135
      Left            =   90
      TabIndex        =   1
      Top             =   9240
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.ListView lvUsername 
      Height          =   8745
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   15425
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   14737632
      BackColor       =   12582912
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Username"
         Object.Width           =   6526
      EndProperty
   End
End
Attribute VB_Name = "frmUsername"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sUsername As String

Private Sub Command1_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Command1_Click"
    Const ContainerName = "frmUsername"
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


    Timer1.Enabled = True
    
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

Private Sub lvUsername_DblClick()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvUsername_DblClick"
    Const ContainerName = "frmUsername"
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


    If lvUsername.SelectedItem Is Nothing Then
    
    Else
        Me.sUsername = lvUsername.SelectedItem.Text
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

Private Sub Timer1_Timer()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Timer1_Timer"
    Const ContainerName = "frmUsername"
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


    On Error Resume Next
    
    If Dir(IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\") + "ENGLISH.TXT", vbNormal) = "" Then
    
        MsgBox "ENGLISH.TXT was not found in the application directory!"
        Timer1.Enabled = False
        Exit Sub
    Else
        
        Dim ilen As Double
        ilen = FileLen(IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\") + "ENGLISH.TXT")
        pb.Value = 0
        pb.Max = 200
        
        Open IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\") + "ENGLISH.TXT" For Input As #22
        Dim sUser As String
        
        Randomize Now / Rnd * 54646
        
        lvUsername.ListItems.Clear
        Do
            
            NumbertoSkip = Round(Rnd * 3000) + 500
            Skipped = 0
            Do
                If EOF(22) Then Exit Do
                Line Input #22, sUser
                Skipped = Skipped + 1
            Loop Until NumbertoSkip = Skipped Or Err.Number <> 0
            
            If Len(sUser) <= 64 Then
                lvUsername.ListItems.Add , , sUser
                lvUsername.Refresh
            End If
            PopulateResellers = PopulateResellers + 1: pb.Value = pb.Value + 1
            gSleep
            If lvUsername.ListItems.Count = 200 Then Exit Do
            On Error Resume Next
        Loop Until EOF(22) Or Err.Number <> 0
        
        Timer1.Enabled = False
        Close #22
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
