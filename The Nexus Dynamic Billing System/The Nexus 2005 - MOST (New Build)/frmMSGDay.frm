VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMSGDay 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sysops Configuration and Message of the Day"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Message of the Day"
      Height          =   6195
      Left            =   5730
      TabIndex        =   3
      Top             =   120
      Width           =   5625
      Begin VB.CommandButton Command1 
         Caption         =   "Apply to Checked Sysops"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   5790
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Apply to All"
         Height          =   285
         Index           =   1
         Left            =   3660
         TabIndex        =   5
         Top             =   5790
         Width           =   1845
      End
      Begin VB.TextBox txtMSG 
         Height          =   5325
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   330
         Width           =   5415
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Sysop Manifest"
      Height          =   6195
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   5595
      Begin VB.CheckBox chkSel 
         Caption         =   "Invert Selection"
         Height          =   375
         Index           =   1
         Left            =   3630
         TabIndex        =   8
         Top             =   270
         Width           =   1845
      End
      Begin VB.CheckBox chkSel 
         Caption         =   "Select All"
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   300
         Width           =   1395
      End
      Begin VB.CommandButton cmdAddSysop 
         Caption         =   "Add Sysop"
         Height          =   285
         Left            =   4200
         TabIndex        =   1
         Top             =   5790
         Width           =   1275
      End
      Begin MSComctlLib.ImageList ilSysops 
         Left            =   480
         Top             =   5610
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
               Picture         =   "frmMSGDay.frx":0000
               Key             =   "k100"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMSGDay.frx":031A
               Key             =   "k080_099"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMSGDay.frx":0634
               Key             =   "k060_079"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMSGDay.frx":0F0E
               Key             =   "k040_059"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMSGDay.frx":1360
               Key             =   "k001_020"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMSGDay.frx":17B2
               Key             =   "k020_039"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMSGDay.frx":1C04
               Key             =   "k000"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvSysops 
         Height          =   4995
         Left            =   150
         TabIndex        =   2
         Top             =   690
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   8811
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
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
Attribute VB_Name = "frmMSgDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkSel_Click(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "chkSel_Click"
    Const ContainerName = "frmMSgDay"
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

Dim ix As Long

    Select Case Index
    Case 0
        chkSel(1).Value = 0
        
        For ix = 1 To lvSysops.ListItems.Count
            lvSysops.ListItems(ix).Checked = -chkSel(0).Value
        Next
    Case 1
        
        For ix = 1 To lvSysops.ListItems.Count
            lvSysops.ListItems(ix).Checked = Not lvSysops.ListItems(ix).Checked
        Next
    
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
    Const ContainerName = "frmMSgDay"
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

Public Sub Savesysops()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Savesysops"
    Const ContainerName = "frmMSgDay"
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
                                "VALUES ('" + itmX.SubItems(1) + "',AES_ENCRYPT('" + itmX.Tag + "','" + odb.colSalts.ReturnSalt(PWSalt) + "'), '" + sSTR.ReplaceString(itmX.SubItems(2), "'", "\'") + "',-1," & CByte(itmX.Text) & "," & Login.lVirtualID & ",0)"
                                
            ElseIf Left(itmX.Key, 1) = "e" Then
            
                Call MySQL.OpenTable(ADOConn, rsSave, , "select * from sysops where RecID = " & Mid(itmX.Key, 2) & " Limit 1")
                
                rsSave!SecurityLevel = CByte(itmX.Text)
                rsSave!Username = itmX.SubItems(1)
                rsSave!Description = itmX.SubItems(2)
                'rsSave!Checked = itmX.Checked
                rsSave.Update
                                
                ADOConn.Execute "UPDATE sysops SET Password = AES_ENCRYPT('" & itmX.Tag & "','" + odb.colSalts.ReturnSalt(PWSalt) + "') where RecID = " & Mid(itmX.Key, 2)
                
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

Private Sub Command1_Click(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Command1_Click"
    Const ContainerName = "frmMSgDay"
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


    Dim lx As Long
        
    Select Case Index
    Case 0
        For lx = 1 To lvSysops.ListItems.Count
            If lvSysops.ListItems(lx).Checked = True Then
                MySQL.Execute ADOConn, "UPDATE sysops SET msg = '" & MySQL.ESC(txtMSG.Text) & "' where RecID = " & Mid(lvSysops.ListItems(lx).Key, 2)
            End If
        Next lx
    Case 1
        Select Case MsgBox("Are you sure you wish to apply this a message of the day to all of the sysops?", vbQuestion + vbYesNo, "Apply to all")
        Case vbYes
            For lx = 1 To lvSysops.ListItems.Count
                lvSysops.ListItems(lx).Checked = True
                MySQL.Execute ADOConn, "UPDATE sysops SET msg = '" & MySQL.ESC(txtMSG.Text) & "' where RecID = " & Mid(lvSysops.ListItems(lx).Key, 2)
            Next lx
        End Select
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

Private Sub Form_Load()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Load"
    Const ContainerName = "frmMSgDay"
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
    
    bResult = MySQL.OpenTable(ADOConn, rsLoad, , MySQL.virtualisp("select sysops.Checked, sysops.SecurityLevel, sysops.RecID, sysops.Username, sysops.description, decode(Password,'" + odb.colSalts.ReturnSalt(PWSalt) + "') as Password from sysops", "sysops"))
    
    If rsLoad.BOF And rsLoad.EOF Then
    
    Else
    
        If rsLoad.RecordCount > 0 Then
            rsLoad.MoveFirst
            While Not rsLoad.EOF And Err.Number = 0
                Set itmX = lvSysops.ListItems.Add(, "r" & rsLoad!RecID, String(3 - Len("" & rsLoad!SecurityLevel), "0") + "" & rsLoad!SecurityLevel)
                itmX.SubItems(1) = rsLoad!Username
                itmX.SubItems(2) = rsLoad!Description
                itmX.Tag = rsLoad!Password
                'itmX.Checked = IIf(rsLoad!Checked <> 0, True, False)
                
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

Private Sub Form_Unload(Cancel As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Unload"
    Const ContainerName = "frmMSgDay"
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


    Savesysops
    
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
    Const ContainerName = "frmMSgDay"
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


    Dim rsMSG As ADODB.Recordset
    
    If MySQL.OpenTable(ADOConn, rsMSG, , "select msg from sysops where RecID = " + Mid(Item.Key, 2)) = True Then
        If Not rsMSG.EOF And Not rsMSG.BOF Then
            txtMSG = IIf(IsNull(rsMSG!Msg), "", rsMSG!Msg)
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
