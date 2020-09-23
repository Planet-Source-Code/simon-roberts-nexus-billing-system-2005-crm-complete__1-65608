VERSION 5.00
Object = "{5B7759CE-C04E-4C5D-993B-8297E30D9065}#1.0#0"; "chilkatftp.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUpgrade 
   BackColor       =   &H00D39969&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Project Alpha Deployment"
   ClientHeight    =   1485
   ClientLeft      =   4365
   ClientTop       =   4725
   ClientWidth     =   7410
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUpgrade2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   1230
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Timer tmrUp 
      Interval        =   3096
      Left            =   8970
      Top             =   180
   End
   Begin VB.Label lblPer 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   90
      TabIndex        =   2
      Top             =   930
      Width           =   7335
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   -300
      Picture         =   "frmUpgrade2.frx":0442
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   1890
   End
   Begin CHILKATFTPLibCtl.ChilkatFTP cFTP 
      Left            =   8970
      OleObjectBlob   =   "frmUpgrade2.frx":0C52
      Top             =   1380
   End
   Begin VB.Label ls 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   825
      Left            =   1050
      TabIndex        =   0
      Top             =   90
      Width           =   6315
   End
End
Attribute VB_Name = "frmUpgrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Steps = 7

Dim oConn As ADODB.Connection

Dim rsload As ADODB.Recordset
Dim iRevision As Integer
Dim bStep As Byte

Private Sub cFTP_GetProgress(ByVal pctDone As Long)

    lblPer.Caption = Round(lblPer.Tag * (pctDone / 100)) & " bytes of " & lblPer.Tag & " downloaded (" & pctDone & "%)"
    
    Me.Refresh
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    pb.Value = pctDone
    
End Sub

Private Sub Form_Load()

    Me.Show
    
    iRevision = GetSetting("projectalpha", "Main", "Revision", -1)
    
    If iRevision = -1 Then
            
        ls.Caption = "Querying for latest Microsoft Installer File"
        
    Else
    
        ls.Caption = "Revision " & iRevision & " detected, verifying..."
        
    End If
    
    pb.Max = Steps
    
End Sub

Private Sub lblStats_Click()

End Sub

Private Sub ls_Change()

    ls.Refresh
    
End Sub



Private Sub tmrUp_Timer()

    
    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "NumCrypt"
    Const ContainerName = "clsMySQL"
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
    
    tmrUp.Enabled = False
    
    Select Case bStep
    
    Case 1 ' Connect to DB and get upgrade record.
        
        ls.Caption = "Connecting to Project Server"
        
        If MySQL.Connection(, , , , oConn) = True Then
       
            ls.Caption = "Connecting to Project Server"
            If iRevision = -1 Then
            
                
                ls.Caption = "Querying for latest Microsoft Installer File"
                bResult = MySQL.OpenTable(oConn, rsload, , "Select RecID, decode(Password,'eafa2804a87078afc643f8148dd8ec78') as Password, Port, Username, Server, Filename, Version, revision, MSI, remotedir from upgrade Where revision > " & iRevision & " and MSI <> 0 Order By revision DESC Limit 1")
                
            Else
            
                ls.Caption = "Querying for latest build"
                bResult = MySQL.OpenTable(oConn, rsload, , "Select RecID, decode(Password,'eafa2804a87078afc643f8148dd8ec78') as Password, Port, Username, Server, Filename, Version, revision, MSI, remotedir from upgrade Where revision > " & iRevision & " Order By revision DESC Limit 1")
                
            End If
            If rsload.RecordCount >= 1 Then
                
                'fUpgrade.sPassWord = rsload!Password
                'fUpgrade.sPort = rsload!Port
                'fUpgrade.sUserName = rsload!Username
                ls.Caption = "Preparing to connect to " & rsload!server & vbCrLf & "Clearing filename " & rsload!FileName
                
                If Dir(IIf(Right(App.path, 1) = "\", App.path, App.path & "\") & "deploy\" & rsload!FileName, vbNormal) <> "" Then
                    ls.Caption = "Removing previous attempt to upgrade/install"
                    Kill IIf(Right(App.path, 1) = "\", App.path, App.path & "\") & "deploy\" & rsload!FileName
                End If
                
                If Dir(IIf(Right(App.path, 1) = "\", App.path, App.path + "\") & "deploy", vbDirectory) = "" Then MkDir App.path & "\deploy"
                
                'fUpgrade.sServer = rsload!server
                'fUpgrade.rsload!Filename = rsload!Filename
                'fUpgrade.bMSI = rsload!MSI
                'fUpgrade.sRemoteDir = rsload!remotedir
                'fUpgrade.Show
            Else
                ls.Caption = "No new version of project alpha to download"
                bStep = 255
            End If
        Else
            ls.Caption = "ERROR: Unable to connect to MySQL server"
            MsgBox "Unable to connect to MySQL server, check that you have the MyODBC Driver installed and that your internet connection is working or your firewall is not preventing access to the Internet.", vbCritical, "Error connecting to Server"
            bStep = 255
        End If
        
    Case 2
        ls.Caption = "Connecting to " & rsload!server
        
        cFTP.UseIEProxy = False
        cFTP.HostName = rsload!server
        cFTP.Username = rsload!Username
        cFTP.Password = rsload!Password
        cFTP.Port = rsload!Port
        cFTP.Connect
        
        Dim bCnt As Byte
        While cFTP.IsConnected <> 1
            ls.Caption = ls.Caption + "."
            bCnt = bCnt + 1
            If bCnt = 10 Then
                ls.Caption = Left(ls.Caption, Len(ls.Caption) - bCnt)
                bCnt = 0
            End If
            DoEvents
        Wend
        
    Case 3
        
        ls.Caption = "Locating file on remote server and commencing download."
        
        Call cFTP.ChangeRemoteDir(rsload!remotedir)
        Dim ftpdir As String
        ftpdir = cFTP.GetCurrentDirListing(rsload!FileName)
        lblPer.Tag = cFTP.GetSize(0)

        If ftpdir = "" Then
            ls.Caption = "There was an error finding the file on the server contact your support officer or local representative."
            bStep = 255
            Exit Sub
        End If
    
        pb.Value = 0
        pb.Max = 100
        
        cFTP.GetFile rsload!FileName, IIf(Right(App.path, 1) = "\", App.path, App.path + "\") & "deploy\" + rsload!FileName
    
    Case 4
    
        ls.Caption = "Waiting for download to complete."
        While FileLen(IIf(Right(App.path, 1) = "\", App.path, App.path + "\") & "deploy\" + rsload!FileName) <> lblPer.Tag
            DoEvents
        Wend
        
    Case 5
    
        ls.Caption = "Disconnecting from local regional server."
        cFTP.Disconnect
    
    Case 6
        Select Case Val(rsload!MSI)
        Case 0
        
            ls.Caption = "Unpacking upgraded EXE's and Resources."
            On Error Resume Next
            Do
                Err.Clear
                If Dir(IIf(Right(App.path, 1) = "\", App.path, App.path + "\") & "projectalpha.exe", vbNormal) <> "" Then Kill IIf(Right(App.path, 1) = "\", App.path, App.path + "\") & "projectalpha.exe"
                If Err.Number <> 0 Then
                    MsgBox "You must shut down project alpha before upgrading can continue!", vbCritical, "Shut down projectalpha"
                End If
            Loop Until Err.Number = 0 Or Err.Number = 53
            
            VFunzip1.AlwaysOverwrite = True
            VFunzip1.DestinationDir = IIf(Right(App.path, 1) = "\", App.path, App.path + "\")
            VFunzip1.toDir = IIf(Right(App.path, 1) = "\", App.path, App.path + "\")
            VFunzip1.FileName = IIf(Right(App.path, 1) = "\", App.path, App.path + "\") & "deploy\" & rsload!FileName
            VFunzip1.UnzipAll
            
            bStep = 255
            
            'If Dir(IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\") & "projectalpha.exe", vbNormal) <> "" Then Shell IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\") & "projectalpha.exe", vbNormalFocus
            
        Case Else
          
        End Select
    
    Case 7
        
        
        Select Case Val(rsload!MSI)
        Case 0
        
        Case Else
        
             Dim path As String, extension As String
    
            path = IIf(Right(App.path, 1) = "\", App.path, App.path + "\")
    
            SaveSetting "projectalpha", "db", "ConnectionString", ""
    
            If Dir(IIf(Right(App.path, 1) = "\", App.path, App.path & "\") & "deploy\" & rsload!FileName, vbNormal) <> "" Then
            
                ShellExecute Me.hwnd, vbNullString, IIf(Right(App.path, 1) = "\", App.path, App.path & "\") & "deploy\" & rsload!FileName, vbNullString, "C:\", SW_SHOWNORMAL
            
                End
                
            End If
        
        End Select
        
        bStep = 255
        
    Case 255
        tmrUp.Enabled = False
        pb.Value = pb.Max
        
        ShellLauncher
        Exit Sub
        
    End Select
    
    DoEvents
    Me.Refresh
        
    tmrUp.Enabled = True
    
    If bStep < 255 Then
        bStep = bStep + 1
        If Not pb.Max = Steps Then pb.Max = Steps
        pb.Value = bStep
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
