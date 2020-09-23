VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   7500
   ClientLeft      =   9300
   ClientTop       =   6315
   ClientWidth     =   9840
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   9840
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   8610
      TabIndex        =   2
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   315
      Left            =   1710
      TabIndex        =   3
      Top             =   7020
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Save && Close"
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Top             =   7020
      Width           =   1575
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   6330
      Index           =   1
      Left            =   180
      ScaleHeight     =   6330
      ScaleWidth      =   9435
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   9435
      Begin VB.PictureBox Picture1 
         Height          =   4365
         Left            =   60
         ScaleHeight     =   4305
         ScaleWidth      =   9045
         TabIndex        =   20
         Top             =   30
         Width           =   9105
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H00A9C9AE&
            ForeColor       =   &H80000008&
            Height          =   4335
            Left            =   0
            ScaleHeight     =   4305
            ScaleWidth      =   9045
            TabIndex        =   21
            Top             =   0
            Width           =   9075
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   4335
         Left            =   9150
         TabIndex        =   19
         Top             =   60
         Width           =   255
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   60
         TabIndex        =   18
         Top             =   4380
         Width           =   9105
      End
      Begin VB.Frame Frame2 
         Caption         =   "eMail Gateway"
         Height          =   1635
         Left            =   60
         TabIndex        =   9
         Top             =   4680
         Width           =   9315
         Begin VB.CheckBox chkManual 
            Alignment       =   1  'Right Justify
            Caption         =   "Manual Override"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   6510
            TabIndex        =   17
            Top             =   180
            Width           =   2685
         End
         Begin VB.TextBox txtSMTP 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   270
            Index           =   3
            Left            =   1770
            TabIndex        =   12
            Top             =   870
            Width           =   3675
         End
         Begin VB.TextBox txtSMTP 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   270
            Index           =   2
            Left            =   1770
            TabIndex        =   14
            Text            =   "support@no1.com.au"
            Top             =   1200
            Width           =   3675
         End
         Begin VB.TextBox txtSMTP 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   270
            Index           =   0
            Left            =   1770
            TabIndex        =   10
            Top             =   240
            Width           =   2805
         End
         Begin VB.TextBox txtSMTP 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   270
            Index           =   1
            Left            =   4620
            TabIndex        =   11
            Text            =   "25"
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "SMTP Domain:"
            Height          =   240
            Index           =   6
            Left            =   240
            TabIndex        =   16
            Top             =   870
            Width           =   1350
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Reply Address:"
            Height          =   240
            Index           =   2
            Left            =   210
            TabIndex        =   15
            Top             =   1230
            Width           =   1350
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "SMTP Server:"
            Height          =   240
            Index           =   1
            Left            =   240
            TabIndex        =   13
            Top             =   270
            Width           =   1350
         End
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   6330
      Index           =   0
      Left            =   210
      ScaleHeight     =   6330
      ScaleWidth      =   9405
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   510
      Visible         =   0   'False
      Width           =   9405
      Begin VB.Frame Frame3 
         Caption         =   "Sysop Manifest"
         Height          =   6195
         Left            =   60
         TabIndex        =   6
         Top             =   30
         Width           =   9315
         Begin VB.CommandButton cmdAddSysop 
            Caption         =   "Add Sysop"
            Enabled         =   0   'False
            Height          =   285
            Left            =   4200
            TabIndex        =   7
            Top             =   5820
            Visible         =   0   'False
            Width           =   1275
         End
         Begin MSComctlLib.ImageList ilSysops 
            Left            =   150
            Top             =   5580
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   7
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOptions.frx":030A
                  Key             =   "k100"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOptions.frx":0624
                  Key             =   "k080_099"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOptions.frx":093E
                  Key             =   "k060_079"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOptions.frx":1218
                  Key             =   "k040_059"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOptions.frx":166A
                  Key             =   "k001_020"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOptions.frx":1ABC
                  Key             =   "k020_039"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOptions.frx":1F0E
                  Key             =   "k000"
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView lvSysops 
            Height          =   5745
            Left            =   150
            TabIndex        =   8
            Top             =   300
            Width           =   9045
            _ExtentX        =   15954
            _ExtentY        =   10134
            View            =   3
            Arrange         =   1
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393217
            Icons           =   "ilSysops"
            SmallIcons      =   "ilSysops"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Level"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Username"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Description"
               Object.Width           =   8819
            EndProperty
         End
      End
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   6795
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   11986
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "System Operators"
            Key             =   "Group1"
            Object.ToolTipText     =   "Set Options for your System Operators"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Primary Scheduler"
            Key             =   "Group2"
            Object.ToolTipText     =   "Set Options for Radius And eMail"
            ImageVarType    =   2
         EndProperty
      EndProperty
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
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iTasks As Integer

Public frmMDI As frmMDIMain

Private Sub chkManual_Click(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "chkManual_Click"
    Const ContainerName = "frmOptions"
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


    Dim bx As Byte
    
    Select Case Index
    Case 2
    Case 1
    Case 0
        txtSMTP(0).Enabled = IIf(chkManual(Index) = 1, True, False)
        txtSMTP(1).Enabled = IIf(chkManual(Index) = 1, True, False)
        txtSMTP(2).Enabled = IIf(chkManual(Index) = 1, True, False)
        txtSMTP(3).Enabled = IIf(chkManual(Index) = 1, True, False)
    End Select
    
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

Private Sub cmdAddSysop_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdAddSysop_Click"
    Const ContainerName = "frmOptions"
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


    Dim ffrmSysop As frmSysopDetails
    Set ffrmSysop = New frmSysopDetails
    ffrmSysop.Show 1
    
    If ffrmSysop.iCloseState = frmCloseSave Then
        Dim itmX As ListItem
        Set itmX = lvSysops.ListItems.Add(, , String(3 - Len("" & ffrmSysop.byLevel), "0") + "" & ffrmSysop.byLevel)
        itmX.SubItems(1) = ffrmSysop.sUsername
        itmX.SubItems(2) = ffrmSysop.sDescription
        itmX.Tag = ffrmSysop.sPassword
        
        Dim imgX As Byte
        Dim imgMin As Byte
        Dim imgMax As Byte
    
        For imgX = 1 To ilSysops.ListImages.Count
            If InStr(ilSysops.ListImages(imgX).Key, "_") > 0 Then
                imgMin = CByte(Mid(ilSysops.ListImages(imgX).Key, 2, 3))
                imgMax = CByte(Mid(ilSysops.ListImages(imgX).Key, 6, 3))
            Else
                imgMin = CByte(Mid(ilSysops.ListImages(imgX).Key, 2, 3))
                imgMax = imgMin
            End If
            If ffrmSysop.byLevel >= imgMin And ffrmSysop.byLevel <= imgMax Then
                itmX.SmallIcon = imgX
            End If
        Next imgX
        
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

Private Sub cmdApply_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdApply_Click"
    Const ContainerName = "frmOptions"
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

    
    'reg.iRadiusHistory = sldDBMins(0).Value
    'reg.iUpkeep = sldDBMins(1).Value
    'reg.iUnpaid = sldDBMins(2).Value
    'reg.iSendPO = sldDBMins(3).Value
    'reg.iUpdate = sldDBMins(4).Value

    reg.smtpServer = txtSMTP(0).Text
    reg.smtpPort = txtSMTP(1).Text
    reg.smtpDomain = txtSMTP(3).Text
    reg.ReplyAddress = txtSMTP(2).Text
    
    Call Savesysops
        
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


Public Sub Savesysops()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Savesysops"
    Const ContainerName = "frmOptions"
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


    Dim itmX As ListItem
    Dim sa As Integer
    Dim rsSave As ADODB.Recordset
    
    If lvSysops.ListItems.Count > 0 Then
    
        For sa = 1 To lvSysops.ListItems.Count
            Set itmX = lvSysops.ListItems(sa)
            If itmX.Key = "" Then
                
                'Call MySQL.OpenTable(ADOConn, rsSave, , "select * from sysops Limit 0")
                
                
                'rsSave.AddNew
                'rsSave!SecurityLevel = CByte(itmX.Text)
                'rsSave!User-Name = itmX.SubItems(1)
                'rsSave!Description = itmX.SubItems(2)
                'rsSave!Password = itmX.Tag
                'rsSave!Checked = itmX.Checked
                'rsSave.Update
                                
                MySQL.Execute ADOConn, "INSERT INTO sysops (Username, Password, Description, Checked, SecurityLevel, VirtualID, bVISP)" + _
                                "VALUES ('" + itmX.SubItems(1) + "',ENCODE('" + itmX.Tag + "','" + odb.colSalts.ReturnSalt(PWSalt) + "'), '" + sSTR.ReplaceString(itmX.SubItems(2), "'", "\'") + "',-1," & CByte(itmX.Text) & "," & Login.lVirtualID & ",0)"
                                
            ElseIf Left(itmX.Key, 1) = "e" Then
            
                Call MySQL.OpenTable(ADOConn, rsSave, , "select * from sysops where RecID = " & Mid(itmX.Key, 2) & " Limit 1")
                
                ADOConn.Execute "UPDATE sysops SET SecurityLevel = '" & CByte(itmX.Text) & "' where RecID = " & Mid(itmX.Key, 2)
                ADOConn.Execute "UPDATE sysops SET Username = '" & MySQL.ESC(itmX.SubItems(1)) & "' where RecID = " & Mid(itmX.Key, 2)
                ADOConn.Execute "UPDATE sysops SET Description = '" & MySQL.ESC(itmX.SubItems(2)) & "' where RecID = " & Mid(itmX.Key, 2)
                ADOConn.Execute "UPDATE sysops SET Checked = '" & IIf(itmX.Checked = True, "-1", "0") & "' where RecID = " & Mid(itmX.Key, 2)
                               
                ADOConn.Execute "UPDATE sysops SET Password = encode('" & itmX.Tag & "','" + odb.colSalts.ReturnSalt(PWSalt) + "') where RecID = " & Mid(itmX.Key, 2)
                
            End If
        Next
    
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
Private Sub cmdCancel_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdCancel_Click"
    Const ContainerName = "frmOptions"
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

Private Sub cmdOk_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdOK_Click"
    Const ContainerName = "frmOptions"
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

    
    Call cmdApply_Click
    
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_KeyDown"
    Const ContainerName = "frmOptions"
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


    

    Dim I As Integer
    'handle ctrl+tab to move to the next tab
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        I = tbsOptions.SelectedItem.Index
        If I = tbsOptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(I + 1)
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

Private Sub Form_Load()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Load"
    Const ContainerName = "frmOptions"
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

    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
    Call tbsOptions_Click

    
    Call LoadInformation
    
    
    txtSMTP(0).Text = reg.smtpServer
    txtSMTP(1).Text = reg.smtpPort
    txtSMTP(2).Text = reg.ReplyAddress
    txtSMTP(3).Text = reg.smtpDomain
        
Exit Sub



ErrorOccur:
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

Exit Sub

    If frmMDI.lvSchedule.ListItems.Count > 0 Then
        Dim iLn As Byte
        
        Dim lx As Long
        Dim itmX As ListItem
        
        For lx = 1 To frmMDI.lvSchedule.ListItems.Count
            Set itmX = frmMDI.lvSchedule.ListItems(lx)
            
            Select Case Left(itmX.Key, 1)
            Case "u", "d"
            
            Case Else
                iTasks = iTasks + 1
                Select Case iTasks
                Case 1
                    'frmSched(iTasks - 1).Caption = itmX.SubItems(2)
                    'frmSched(iTasks - 1).Move 120, 120, frmSched(iTasks - 1).Width, frmSched(iTasks - 1).Height
                    'frmSched(iTasks - 1).Tag = itmX.Key
                Case Else
                    'Load frmSched(iTasks - 1)

                    'sldDBMins(iTasks - 1).Value = Val(itmX.SubItems(1))
                    
                    'iLn = iLn + 1
                    'If iLn = 7 Then
                    '    frmSched(iTasks - 1).Move frmSched(iTasks - 2).Left + frmSched(iTasks - 1).Width + 120, 120, frmSched(iTasks - 1).Width, frmSched(iTasks - 1).Height
                    '    iLn = 0
                    'Else
                    '    frmSched(iTasks - 1).Move frmSched(iTasks - 2).Left, 120 + frmSched(iTasks - 2).Top + frmSched(iTasks - 2).Height, frmSched(iTasks - 1).Width, frmSched(iTasks - 1).Height
                    'End If
                    'frmSched(iTasks - 1).Caption = itmX.SubItems(2)
                    'frmSched(iTasks - 1).Tag = itmX.Key
                    'frmSched(iTasks - 1).Visible = True
                    'lblDBMins(iTasks - 1).Visible = True
                    'opHorz(iTasks - 1).Visible = True
                    'optVert(iTasks - 1).Visible = True
                    'sldDBMins(iTasks - 1).Visible = True
                End Select
            End Select
        Next
    Else
        'Unload frmSched(0)
    
    End If
    
End Sub

Private Sub Label6_Click(Index As Integer)
End Sub

Private Sub lvsysops_DblClick()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvsysops_DblClick"
    Const ContainerName = "frmOptions"
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


    Exit Sub
    
    If lvSysops.Tag <> "" Then
    
        Dim ffrmSysop As frmSysopDetails
        Set ffrmSysop = New frmSysopDetails
        Dim itmX As ListItem
        Set itmX = lvSysops.SelectedItem
        
        ffrmSysop.sUsername = itmX.SubItems(1)
        ffrmSysop.sPassword = IIf(IsNull(itmX.Tag), "binary101", itmX.Tag)
        ffrmSysop.sDescription = itmX.SubItems(2)
        ffrmSysop.byLevel = itmX.Text
        ffrmSysop.Show 1
        
        If ffrmSysop.iCloseState = frmCloseSave Then
            itmX.Text = String(3 - Len("" & ffrmSysop.byLevel), "0") + "" & ffrmSysop.byLevel
            itmX.SubItems(1) = ffrmSysop.sUsername
            itmX.SubItems(2) = ffrmSysop.sDescription
            itmX.Tag = ffrmSysop.sPassword
            If itmX.Key <> "" Then itmX.Key = "e" + Mid(itmX.Key, 2)
        End If
    
        Dim imgX As Byte
        Dim imgMin As Byte
        Dim imgMax As Byte
    
        For imgX = 1 To ilSysops.ListImages.Count
            If InStr(ilSysops.ListImages(imgX).Key, "_") > 0 Then
                imgMin = CByte(Mid(ilSysops.ListImages(imgX).Key, 2, 3))
                imgMax = CByte(Mid(ilSysops.ListImages(imgX).Key, 6, 3))
            Else
                imgMin = CByte(Mid(ilSysops.ListImages(imgX).Key, 2, 3))
                imgMax = imgMin
            End If
            If ffrmSysop.byLevel >= imgMin And ffrmSysop.byLevel <= imgMax Then
                itmX.SmallIcon = imgX
            End If
        Next imgX
    
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

Private Sub lvsysops_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvsysops_ItemClick"
    Const ContainerName = "frmOptions"
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


    lvSysops.Tag = True
    
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

Private Sub sldDBMins_Click(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "sldDBMins_Click"
    Const ContainerName = "frmOptions"
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


    'lblDBMins(Index).Caption = "" & sldDBMins(Index).Value
    
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

Private Sub tbsOptions_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "tbsOptions_Click"
    Const ContainerName = "frmOptions"
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

    
    Dim I As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For I = 0 To picOptions.UBound
        If I = tbsOptions.SelectedItem.Index - 1 Then
            picOptions(I).Move tbsOptions.clientLeft, tbsOptions.clientTop, tbsOptions.clientWidth, tbsOptions.clientHeight
            picOptions(I).ZOrder 0
            picOptions(I).Visible = True
            picOptions(I).Enabled = True
        Else
            picOptions(I).Visible = False
            picOptions(I).Enabled = False
        End If
    Next
    
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

Public Function LoadInformation()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "LoadInformation"
    Const ContainerName = "frmOptions"
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


    Dim rsLoad As ADODB.Recordset
    Dim itmX As ListItem
    Dim bResult As Boolean
    
    bResult = MySQL.OpenTable(ADOConn, rsLoad, , MySQL.virtualisp("select distinct sysops.Checked ,sysops.SecurityLevel ,sysops.RecID ,sysops.Username ,sysops.Description from sysops ", "sysops", False, Login.bMaster))
    lvSysops.ListItems.Clear
    If rsLoad.BOF And rsLoad.EOF Then
    
    Else
    
        If rsLoad.RecordCount > 0 Then
            rsLoad.MoveFirst
            While Not rsLoad.EOF And Err.Number = 0
                If rsLoad!RecID <> 1 Then
                
                    Set itmX = lvSysops.ListItems.Add(, "r" & rsLoad!RecID, String(3 - Len("" & rsLoad!SecurityLevel), "0") + "" & rsLoad!SecurityLevel)
                    itmX.SubItems(1) = rsLoad!Username
                    itmX.SubItems(2) = rsLoad!Description
                    'itmX.Tag = rsload!Password
                    itmX.Checked = IIf(rsLoad!Checked <> 0, True, False)
                    
                    Dim imgX As Byte
                    Dim imgMin As Byte
                    Dim imgMax As Byte
                
                    For imgX = 1 To ilSysops.ListImages.Count
                        If InStr(ilSysops.ListImages(imgX).Key, "_") > 0 Then
                            imgMin = CByte(Mid(ilSysops.ListImages(imgX).Key, 2, 3))
                            imgMax = CByte(Mid(ilSysops.ListImages(imgX).Key, 6, 3))
                        Else
                            imgMin = CByte(Mid(ilSysops.ListImages(imgX).Key, 2, 3))
                            imgMax = imgMin
                        End If
                        If CByte(rsLoad!SecurityLevel) >= imgMin And CByte(rsLoad!SecurityLevel) <= imgMax Then
                            itmX.SmallIcon = imgX
                        End If
                    Next imgX
                
                End If
                
                rsLoad.MoveNext
            Wend
        End If
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

