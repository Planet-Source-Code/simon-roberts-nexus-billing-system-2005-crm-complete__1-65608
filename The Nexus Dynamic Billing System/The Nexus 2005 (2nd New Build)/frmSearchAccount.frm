VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmSearchAccount 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search For Account"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6015
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Refered By"
      Height          =   8055
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   5805
      Begin VB.Frame Frame5 
         Caption         =   "Search Options"
         Height          =   1605
         Left            =   150
         TabIndex        =   1
         Top             =   330
         Width           =   5535
         Begin VB.TextBox txtSearchName 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   180
            TabIndex        =   5
            ToolTipText     =   "[Shift + Enter] to search database on your string entry"
            Top             =   300
            Width           =   5205
         End
         Begin VB.CheckBox chkSearchAccountNames 
            Caption         =   "Search Account Names"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            TabIndex        =   4
            Top             =   840
            Value           =   1  'Checked
            Width           =   3675
         End
         Begin VB.CheckBox chkSearchContacts 
            Caption         =   "Search All Contact Names"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            TabIndex        =   3
            Top             =   1140
            Value           =   1  'Checked
            Width           =   3675
         End
         Begin MSComctlLib.ProgressBar pb1 
            Height          =   105
            Left            =   0
            TabIndex        =   2
            Top             =   1440
            Width           =   5505
            _ExtentX        =   9710
            _ExtentY        =   185
            _Version        =   393216
            Appearance      =   0
            Max             =   6
         End
      End
      Begin MSComctlLib.ListView lvReferedBy 
         Height          =   5835
         Left            =   120
         TabIndex        =   6
         Top             =   2010
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   10292
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   16777215
         BackColor       =   8421504
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Account Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Contact Name"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "frmSearchAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lRecID As Variant
Public sAccountName As String


Private Sub lvReferedBy_DblClick()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvReferedBy_DblClick"
    Const ContainerName = "frmSearchAccount"
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


    If Not lvReferedBy.SelectedItem Is Nothing Then
            
        Select Case Left(lvReferedBy.SelectedItem.Key, 1)
        Case "a"
            Me.lRecID = Val(Mid(lvReferedBy.SelectedItem.Key, 2))
        Case "d"
            Me.lRecID = Val(Mid(lvReferedBy.SelectedItem.Key, InStr(lvReferedBy.SelectedItem.Key, "_") + 1))
        End Select
        Me.sAccountName = lvReferedBy.SelectedItem.Text
        Unload Me
    Else
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


Private Sub txtSearchName_KeyPress(KeyAscii As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtSearchName_KeyPress"
    Const ContainerName = "frmSearchAccount"
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

    Select Case KeyAscii
    Case 13
        KeyAscii = 0
        pb1.Value = 0
        If Trim(txtSearchName) = "" Then Exit Sub
        Dim rsload As adodb.Recordset
        
        If chkSearchAccountNames.Value <> 0 Then
            bResult = MySQL.OpenTable(directConn, rsload, , MySQL.virtualisp("select * from accountinfo Where AccountName " & IIf(InStr(LCase(txtSearchName.Text), "null") > 0, " = NULL ", " Like '%" & txtSearchName.Text & "%'"), "accountinfo"))
            
            If rsload.RecordCount > 0 Then
                rsload.MoveFirst
                While Not rsload.EOF And Err.Number = 0
                    Set itmX = lvReferedBy.ListItems.Add(, "a" & rsload!RecID, rsload!AccountName)
                    rsload.MoveNext
                Wend
            End If
        End If
        
        pb1.Value = pb1.Value + 1
        
        If chkSearchContacts.Value <> 0 Then
        
            bResult = MySQL.OpenTable(directConn, rsload, , MySQL.virtualisp("select accountinfo.AccountName, acci_addresses.Contactname, acci_addresses.AccI_RecID, acci_addresses.RecID from accountinfo, acci_addresses Where acci_addresses.AccI_RecID = accountinfo.RecID AND acci_addresses.ContactName " & IIf(InStr(LCase(txtSearchName.Text), "null") > 0, " = NULL ", " Like '%" & txtSearchName.Text & "%'"), "accountinfo"))
            
            If rsload.RecordCount > 0 Then
                rsload.MoveFirst
                While Not rsload.EOF And Err.Number = 0
                    Set itmX = lvReferedBy.ListItems.Add(, "d" & rsload!RecID & "_" & rsload!acci_RecID, rsload!AccountName)
                    itmX.SubItems(1) = rsload!ContactName
                    rsload.MoveNext
                Wend
            End If
        
        End If
        pb1.Value = pb1.Value + 1
        
        If chkSearchContacts.Value <> 0 Then
        
            bResult = MySQL.OpenTable(directConn, rsload, , MySQL.virtualisp("select accountinfo.AccountName, acci_emailaddresses.Contactname, acci_emailaddresses.AccI_RecID, acci_emailaddresses.RecID from accountinfo, acci_emailaddresses Where acci_emailaddresses.AccI_RecID = accountinfo.RecID AND acci_emailaddresses.ContactName " & IIf(InStr(LCase(txtSearchName.Text), "null") > 0, " = NULL ", " Like '%" & txtSearchName.Text & "%'"), "accountinfo"))
            
            If rsload.RecordCount > 0 Then
                rsload.MoveFirst
                While Not rsload.EOF And Err.Number = 0
                    Set itmX = lvReferedBy.ListItems.Add(, "d" & rsload!RecID & "_" & rsload!acci_RecID, rsload!AccountName)
                    itmX.SubItems(1) = rsload!ContactName
                    rsload.MoveNext
                Wend
            End If
        
        End If
        pb1.Value = pb1.Value + 1
        
        If chkSearchContacts.Value <> 0 Then
        
            bResult = MySQL.OpenTable(directConn, rsload, , MySQL.virtualisp("select accountinfo.AccountName, acci_services.Contactname, acci_services.AccI_RecID, acci_services.RecID from accountinfo, acci_services Where acci_services.AccI_RecID = accountinfo.RecID AND acci_services.ContactName " & IIf(InStr(LCase(txtSearchName.Text), "null") > 0, " = NULL ", " Like '%" & txtSearchName.Text & "%'"), "accountinfo"))
            
            If rsload.RecordCount > 0 Then
                rsload.MoveFirst
                While Not rsload.EOF And Err.Number = 0
                    Set itmX = lvReferedBy.ListItems.Add(, "d" & rsload!RecID & "_" & rsload!acci_RecID, rsload!AccountName)
                    itmX.SubItems(1) = rsload!ContactName
                    rsload.MoveNext
                Wend
            End If
        
        End If
        pb1.Value = pb1.Value + 1
        
        pb1.Value = pb1.Value + 1
        
        If chkSearchContacts.Value <> 0 Then
        
            bResult = MySQL.OpenTable(directConn, rsload, , MySQL.virtualisp("select accountinfo.AccountName, acci_phonenumbers.Contactname, acci_phonenumbers.AccI_RecID, acci_phonenumbers.RecID from accountinfo, acci_phonenumbers Where acci_phonenumbers.AccI_RecID = accountinfo.RecID AND acci_phonenumbers.ContactName " & IIf(InStr(LCase(txtSearchName.Text), "null") > 0, " = NULL ", " Like '%" & txtSearchName.Text & "%'"), "accountinfo"))
            
            If rsload.RecordCount > 0 Then
                rsload.MoveFirst
                While Not rsload.EOF And Err.Number = 0
                    Set itmX = lvReferedBy.ListItems.Add(, "d" & rsload!RecID & "_" & rsload!acci_RecID, rsload!AccountName)
                    itmX.SubItems(1) = rsload!ContactName
                    rsload.MoveNext
                Wend
            End If
        End If
        pb1.Value = pb1.Value + 1
       
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
