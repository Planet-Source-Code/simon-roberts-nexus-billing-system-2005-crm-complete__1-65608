VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3990
   ClientLeft      =   5595
   ClientTop       =   5130
   ClientWidth     =   7470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   266
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   498
   Begin VB.Timer timeout 
      Interval        =   10000
      Left            =   4710
      Top             =   2100
   End
   Begin VB.PictureBox picBg 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   1230
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   103
      TabIndex        =   3
      Top             =   1770
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer tmrDBMapper 
      Enabled         =   0   'False
      Interval        =   66
      Left            =   3150
      Top             =   1890
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3150
      Top             =   2550
   End
   Begin VB.Timer tmrMoveForm 
      Interval        =   1
      Left            =   5940
      Top             =   4020
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   4020
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Timer tmrUnload 
      Interval        =   2000
      Left            =   6540
      Top             =   3990
   End
   Begin VB.Label lblAction 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004FD2F9&
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   7035
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   4620
      Width           =   765
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public bFinished As Boolean

Function GetWindowsDir() As String
    
    '*[ Error Checking Variables ]**********************************************************************************
    Const RoutineName = "frmSplash"
    Const ContainerName = "GetWindowsDir"
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
    
    Dim Temp As String
    Dim Ret As Long
    Const MAX_LENGTH = 145

    Temp = String$(MAX_LENGTH, 0)
    Ret = GetWindowsDirectory(Temp, MAX_LENGTH)
    Temp = Left$(Temp, Ret)
    If Temp <> "" And Right$(Temp, 1) <> "\" Then
        GetWindowsDir = Temp & "\"
    Else
        GetWindowsDir = Temp
    End If
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



Private Sub Form_Load()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Load"
    Const ContainerName = "frmSplash"
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


    lblVersion.Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision

    
    
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

Private Sub Form_Paint()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Paint"
    Const ContainerName = "frmSplash"
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


    Static bAgent As Boolean
    
    If bAgent = False Then
        
        DirName = GetWindowsDir()

        'frmAgent.oChar.Speak "Welcome to project alpha, this intuative wan client allows you to interact with the carriers service."
        bAgent = True
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

Private Sub timeout_Timer()

    Static lastlblAction As String
    Static LastCount As Long
    
    If lastlblAction = lblAction.Caption Then
        LastCount = LastCount + 10
        
        If LastCount > 50 Then
            
                Msg = "There appears no action on intial login. If the login screen hasn't appeared then it is possible that this is due to no direct TCP/IP SSL Connection to the server."
                Msg = Msg + ". This can occur from TCP/IP Faults and Firewalls." & vbCrLf
                Msg = Msg + vbCrLf + "Also it means that you may not be able to access this node of the project alpha's server farm or log onto this server at this time."
                Msg = Msg + "Please be advised to contact your system administrator reguarding this and allow for time to check permissions on firewalls."
                Msg = Msg + vbCrLf & vbCrLf & "You will have to allow port 3306 on your hardware firewall for MySQL and install the latest MyODBC Drivers."
                Msg = Msg + "These are available from http://www.mysql.com/, under 'binary downloads'."
                Msg = Msg + vbCrLf & vbCrLf & "Would you like to end the application now or would you prefer to wait for a futher minute?"
                
                Select Case MsgBox(Msg, vbCritical + vbYesNo, "Timeout")
                Case vbYes
                    End
                Case vbNo
                    LastCount = 0
                End Select
        End If
    Else
        lastlblAction = lblAction.Caption
    End If
    
End Sub

Private Sub Timer1_Timer()


    Static lPC(5, 5) As Long ' Previous Colour
    Static sX As Single
    Static sY As Single
    Static bStage As Byte
    
    Dim X As Single
    Dim Y As Single
    
    Select Case bStage
    Case 0
        
        Randomize Now / Rnd ^ Rnd
        
        sX = Round(Me.ScaleWidth * Rnd)
        sY = Round(Me.ScaleHeight * Rnd)
        
        For X = 5 To 1 Step -1
            For Y = 5 To 1 Step -1
                lPC(X, Y) = Me.Point((sX + 2) - X, (sY + 2) - Y)
            Next Y
        Next X
        bStage = 1
    Case 1
        
        Me.PSet (sX, sY), RGB(250, 250, 250)
        bStage = 2
    
    Case 2
        
        Me.PSet (sX - 1, sY), RGB(250, 250, 250)
        Me.PSet (sX, sY - 1), RGB(250, 250, 250)
        bStage = 3
        
    Case 3
    
        Me.PSet (sX + 1, sY), RGB(250, 250, 250)
        Me.PSet (sX, sY + 1), RGB(250, 250, 250)
    
        bStage = 4
        
    Case 4
    
        Me.PSet (sX - 2, sY), RGB(250, 250, 250)
        Me.PSet (sX, sY - 2), RGB(250, 250, 250)
        Me.PSet (sX + 2, sY), RGB(250, 250, 250)
        Me.PSet (sX, sY + 2), RGB(250, 250, 250)

        bStage = 5
        
    Case 5

        Me.PSet (sX - 1, sY - 1), RGB(250, 250, 250)
        Me.PSet (sX + 1, sY - 1), RGB(250, 250, 250)
        Me.PSet (sX - 1, sY + 1), RGB(250, 250, 250)
        Me.PSet (sX + 1, sY - 1), RGB(250, 250, 250)
    
        bStage = 6
        
    Case 6 To 10
    
        bStage = bStage + 1
        
        
    Case 11
    
        For X = 5 To 1 Step -1
            For Y = 5 To 1 Step -1
                Me.PSet ((sX + 2) - X, (sY + 2) - Y), lPC(Y, X)
            Next Y
        Next X

        bStage = 0
        
    End Select
    
    gSleep

    
End Sub

Private Sub tmrDBMapper_Timer()

    Static rstSchema As ADODB.Recordset
    Static iCount As Integer
    
    If rstSchema Is Nothing Then
    
        Set rstSchema = ADOConn.OpenSchema(adSchemaTables)
        Call odb.colDBObjects.Clear
        odb.colDBObjects.dbTables = rstSchema.RecordCount
        pb1.Max = pb1.Max + rstSchema.RecordCount
    End If
    
    If rstSchema.EOF = True Then
        tmrDBMapper.Enabled = False
        tmrDBMapper.Interval = 0
        rstSchema.Close
        Exit Sub
    End If
    
    Dim X As Long
    Static rsDesc As ADODB.Recordset
    Static rsload As ADODB.Recordset
    
    If rsDesc Is Nothing Then
        On Error GoTo ProfBuild
        bResult = MySQL.OpenTable(ADOConn, rsload, , "Select * from " & rstSchema!TABLE_NAME & " Limit 1,1")
        bResult = MySQL.OpenTable(ADOConn, rsDesc, , "describe " & rstSchema!TABLE_NAME & "")
        If Err.Number > 0 Then GoTo ProfBuild
        iCount = 1
        tmrDBMapper.Interval = 150
    Else
        iCount = iCount + 1
        If Not tmrDBMapper.Interval = 7 Then tmrDBMapper.Interval = 7
        If iCount > rsload.Fields.Count Then
            rstSchema.MoveNext
            Set rsDesc = Nothing
            Set rsload = Nothing
            If pb1.Value + 1 < pb1.Max Then pb1.Value = pb1.Value + 1
            Exit Sub
        End If
    End If
    
          
    Select Case LCase(Left(rstSchema!TABLE_NAME, 3))

    
    Case Else
     
        If rsload.State = adStateOpen Then
                
           lblAction.Caption = "Building Database Profile [`" & "projectalpha" & "`.`" & Left(rstSchema!TABLE_NAME, 3) + Right(rstSchema!TABLE_NAME, 2) & "`.`" & rsload.Fields(iCount - 1).Name & "`]"
           lblAction.Refresh
           
           If rsDesc.State = adStateOpen Then
               rsDesc.Filter = "Field = '" & rsload.Fields(iCount - 1).Name & "'"
               Call odb.colDBObjects.Add("f_" & X & "_" & rstSchema!TABLE_NAME, Login.lLevel, IIf(rsDesc!Null = "YES", True, False), IIf(rsDesc!Key = "PRI" And rsDesc!Extra = "auto_increment", True, False), IIf(IsNull(rsDesc!Extra), "", rsDesc!Extra), IIf(IsNull(rsDesc!Default), "", rsDesc!Default), IIf(IsNull(rsDesc!Key), "", rsDesc!Key), "projectalpha", rstSchema!TABLE_NAME, rsload.Fields(iCount - 1).Name, 0, rsload.Fields(iCount - 1).DefinedSize, rsload.Fields(iCount - 1).NumericScale, rsload.Fields(iCount - 1).Precision, rsload.Fields(iCount - 1).Status, rsload.Fields(iCount - 1).Type, MySQL.fldType(rsload.Fields(iCount - 1).Type), rsload.Fields(iCount - 1).Attributes)
           Else
               Call odb.colDBObjects.Add("f_" & X & "_" & rstSchema!TABLE_NAME, Login.lLevel, True, IIf(rstSchema!fIELD_NAME = "RecID", True, False), "", _
                                         "NULL", "", "projectalpha", rstSchema!TABLE_NAME, rsload.Fields(iCount - 1).Name, 0, rsload.Fields(iCount - 1).DefinedSize, rsload.Fields(iCount - 1).NumericScale, rsload.Fields(iCount - 1).Precision, rsload.Fields(iCount - 1).Status, rsload.Fields(iCount - 1).Type, MySQL.fldType(rsload.Fields(iCount - 1).Type), rsload.Fields(iCount - 1).Attributes)
           End If
           gSleep
        End If
             
    End Select
    
    gSleep
    
    Exit Sub
                 
ProfBuild:

If Err.Number <> 0 Then

     Select Case Err.Number
     Case -2147467259 ' Axxess is Denied
        Err.Clear
        
        Call odb.colDBObjects.Add("f_" & X & "_" & rstSchema!TABLE_NAME, 0, False, False, "", _
                                        "", "", "projectalpha", rstSchema!TABLE_NAME, "Access Denied", 0, 0, 0, 0, 0, 0, "Access Denied", 0)
        
     Case Else
        
        Call odb.colDBObjects.Add("f_" & X & "_" & rstSchema!TABLE_NAME, 0, False, False, "", _
                                        "", "", "projectalpha", rstSchema!TABLE_NAME, "Error " & Err.Number, 0, 0, 0, 0, 0, 0, Err.Description, 0)
        Err.Clear
        
        
     End Select

Else


    Call odb.colDBObjects.Add("f_" & X & "_" & rstSchema!TABLE_NAME, 0, False, False, "", _
         "", "", "projectalpha", rstSchema!TABLE_NAME, "General Error", 0, 0, 0, 0, 0, 0, "Could Not Select Table", 0)
                                
End If
    
    
        
End Sub

Private Sub tmrMoveForm_Timer()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "tmrMoveForm_Timer"
    Const ContainerName = "frmSplash"
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


    If Me.Height < 5320 Then
        Me.Height = Me.Height + IIf(Me.Height < 4600, 5, IIf(Me.Height < 5000, 10, 15))
    Else
        tmrMoveForm.Enabled = False
        RunOpenSequence pb1, lblAction
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

Public Sub tmrUnload_Timer()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "tmrUnload_Timer"
    Const ContainerName = "frmSplash"
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


    
    If bFinished = True Then Unload Me
    
    
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

