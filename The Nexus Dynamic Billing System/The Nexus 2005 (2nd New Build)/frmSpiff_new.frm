VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmSpiff 
   BackColor       =   &H00AF94D3&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bonuses and Spiff"
   ClientHeight    =   8160
   ClientLeft      =   990
   ClientTop       =   2220
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      BackColor       =   &H00AF94D3&
      Caption         =   "Current plan or service you are working with"
      Height          =   945
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7665
      Begin VB.Label lblDesc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   540
         Left            =   180
         TabIndex        =   9
         Top             =   270
         Width           =   7380
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00AF94D3&
      Caption         =   "&Save And Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4980
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7500
      Width           =   2805
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00AF94D3&
      Caption         =   "Bonues and Spiff's"
      Height          =   6225
      Left            =   120
      TabIndex        =   0
      Top             =   1140
      Width           =   7665
      Begin VB.Frame Frame3 
         BackColor       =   &H00AF94D3&
         Height          =   5055
         Left            =   150
         TabIndex        =   6
         Top             =   1050
         Width           =   7365
         Begin MSComctlLib.ListView lvSpiff 
            Height          =   4635
            Left            =   150
            TabIndex        =   7
            Top             =   240
            Width           =   7065
            _ExtentX        =   12462
            _ExtentY        =   8176
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            PictureAlignment=   1
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Type"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Units"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Award To"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Minimum"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "Maximum"
               Object.Width           =   2540
            EndProperty
            Picture         =   "frmSpiff_new.frx":0000
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00AF94D3&
         ClipControls    =   0   'False
         Height          =   735
         Left            =   150
         TabIndex        =   1
         Top             =   240
         Width           =   7365
         Begin VB.CommandButton cmdAdd 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6810
            TabIndex        =   4
            Top             =   240
            Width           =   435
         End
         Begin VB.ComboBox cmbUnit 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   210
            Width           =   5625
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Unit Type:"
            Height          =   255
            Left            =   150
            TabIndex        =   3
            Top             =   330
            Width           =   795
         End
      End
   End
End
Attribute VB_Name = "frmSpiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ptRecID As Long

Dim oSpiff As New clsBonuses
Dim oComm As New clsCommission

Private Sub cmdAdd_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdAdd_Click"
    Const ContainerName = "frmSpiff"
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


    If cmbUnit.ListIndex = -1 Then
        MsgBox "You must select a Unit Type of Award to give"
        Exit Sub
    End If
    
    Dim fSpiffCon As New frmSpiffCon
    Set fSpiffCon.oBonus = oComm.clsSpiff.colSpiff.Add(Me.ptRecID & "_s" & oComm.clsSpiff.colSpiff.Count + 1, toAgency, cmbUnit.ItemData(cmbUnit.ListIndex), 0, 0, 0, Login.lVirtualID, lblDesc.Tag, Me.ptRecID & "_s" & oComm.clsSpiff.colSpiff.Count + 1)
    
    fSpiffCon.Show 1
    
    Dim itmX As ListItem
    
    fSpiffCon.oBonus.SaveFlag = 1
    
    Set itmX = lvSpiff.ListItems.Add(, fSpiffCon.oBonus.Key, cmbUnit.Text)
    itmX.SubItems(1) = fSpiffCon.oBonus.UnitsAwarded
    itmX.SubItems(2) = IIf(fSpiffCon.oBonus.enumSource = 3, "Developer", IIf(fSpiffCon.oBonus.enumSource = 2, "Agency", IIf(fSpiffCon.oBonus.enumSource = 1, "Site", IIf(fSpiffCon.oBonus.enumSource = 0, "Sysop", "Not Set"))))
    itmX.SubItems(3) = fSpiffCon.oBonus.Min
    itmX.SubItems(4) = fSpiffCon.oBonus.Max
        
    
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

Private Sub Command1_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Command1_Click"
    Const ContainerName = "frmSpiff"
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

Private Sub Command2_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Command2_Click"
    Const ContainerName = "frmSpiff"
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


    
    Dim ix As Integer
    Dim oSpif As clsBonuses
    
    If oComm.clsSpiff.colSpiff.Count > 0 Then
        For ix = 1 To oComm.clsSpiff.colSpiff.Count
            Set oSpif = oComm.clsSpiff.colSpiff(ix)
            Select Case oSpif.SaveFlag
            Case 0
            
            Case 1
                MySQL.Execute directConn, "INSERT INTO bonus_awards (ptRecID, AwardTo, Min, Max, UnitType, Units, VirtualID, sKey) " + "VALUES(" & oSpif.ptRecID & "," & oSpif.enumSource & "," & oSpif.Min & "," & oSpif.Max & "," & oSpif.UnitType & "," & oSpif.UnitsAwarded & "," & oSpif.VirtualID & ",'" & oSpif.Key & "')"
            Case 2
                MySQL.Execute directConn, "UPDATE bonus_awards SET ptRecID = " & oSpif.ptRecID & ", AwardTo = " & oSpif.enumSource & ", Min = " & oSpif.Min & ", Max = " & oSpif.Max & ", UnitType = " & oSpif.UnitType & ", Units = " & oSpif.UnitsAwarded & " where sKey = '" & oSpif.Key & "'"
            End Select
            oSpif.SaveFlag = 0
        Next
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

Private Sub Form_Load()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Load"
    Const ContainerName = "frmSpiff"
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


    Call GUI.LoadColWidths(lvSpiff, Me)
    

    Dim rsOpen As adodb.Recordset
    
    If MySQL.OpenTable(directConn, rsOpen, , "select RecID, Description from plantypes where RecID = " & ptRecID) = True Then
        If rsOpen.BOF And rsOpen.EOF Then
        
        Else
            If rsOpen.RecordCount > 0 Then
                lblDesc.Caption = rsOpen!Description
                lblDesc.Tag = rsOpen!RecID
            End If
        End If
    End If
    
    
    If MySQL.OpenTable(directConn, rsOpen, , "select * from bonus_units") = True Then
        If rsOpen.BOF And rsOpen.EOF Then
        
        Else
            While Not rsOpen.EOF And Err.Number = 0
                oComm.clsSpiff.colUnitType.Add "s" & rsOpen!RecID, rsOpen!UnitName, rsOpen!FieldName
                cmbUnit.AddItem rsOpen!UnitName
                cmbUnit.ItemData(cmbUnit.ListCount - 1) = rsOpen!RecID
                rsOpen.MoveNext
            Wend
            
        End If
    End If
    
  Dim itmX As ListItem
  
    If MySQL.OpenTable(directConn, rsOpen, , MySQL.virtualisp("select bonus_units.UnitName, bonus_awards.* from bonus_awards, bonus_units where bonus_units.RecID = bonus_awards.UnitType and ptRecID = " & Me.ptRecID, "bonus_awards")) = True Then
        If rsOpen.EOF And rsOpen.BOF Then
        
        Else
            While Not rsOpen.EOF And Err.Number = 0
                
                oComm.clsSpiff.colSpiff.Add rsOpen!sKey, rsOpen!AwardTo, rsOpen!UnitType, rsOpen!Min, rsOpen!Max, rsOpen!units, rsOpen!VirtualID, rsOpen!ptRecID, rsOpen!sKey
                
                Set itmX = lvSpiff.ListItems.Add(, rsOpen!sKey, rsOpen!UnitName)
                itmX.SubItems(1) = rsOpen!units
                itmX.SubItems(2) = IIf(rsOpen!AwardTo = 3, "Developer", IIf(rsOpen!AwardTo = 2, "Agency", IIf(rsOpen!AwardTo = 1, "Site", IIf(rsOpen!AwardTo = 0, "Sysop", "Not Set"))))
                itmX.SubItems(3) = rsOpen!Min
                itmX.SubItems(4) = rsOpen!Max
                rsOpen.MoveNext
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

Private Sub Form_Unload(Cancel As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Unload"
    Const ContainerName = "frmSpiff"
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


    Call GUI.SaveColWidths(lvSpiff, Me)
    
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

Private Sub txtQuanity_KeyPress(KeyAscii As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtQuanity_KeyPress"
    Const ContainerName = "frmSpiff"
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
    Case 48 To 57, 8
    Case Asc(".")
        If InStr(txtunit, ".") > 0 Then KeyAscii = 0
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

Private Sub lvSpiff_DblClick()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvSpiff_DblClick"
    Const ContainerName = "frmSpiff"
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


    If lvSpiff.SelectedItem Is Nothing Then
    
    Else
        Dim fSpiffCon As New frmSpiffCon
        Set fSpiffCon.oBonus = oComm.clsSpiff.colSpiff(lvSpiff.SelectedItem.Key)
        fSpiffCon.Show 1
            
        Set itmX = lvSpiff.SelectedItem
        itmX.SubItems(1) = fSpiffCon.oBonus.UnitsAwarded
        itmX.SubItems(2) = IIf(fSpiffCon.oBonus.enumSource = 3, "Developer", IIf(fSpiffCon.oBonus.enumSource = 2, "Agency", IIf(fSpiffCon.oBonus.enumSource = 1, "Site", IIf(fSpiffCon.oBonus.enumSource = 0, "Sysop", "Not Set"))))
        itmX.SubItems(3) = fSpiffCon.oBonus.Min
        itmX.SubItems(4) = fSpiffCon.oBonus.Max
        
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
