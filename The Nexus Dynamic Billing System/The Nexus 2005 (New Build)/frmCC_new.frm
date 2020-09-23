VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmCC 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Credit Card"
   ClientHeight    =   5325
   ClientLeft      =   1395
   ClientTop       =   4185
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCC_new.frx":0000
   ScaleHeight     =   5325
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Use Selected Credit Card"
      Height          =   375
      Left            =   9030
      TabIndex        =   2
      Top             =   4800
      Width           =   2115
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add a Credit Card"
      Height          =   375
      Left            =   90
      TabIndex        =   1
      Top             =   4770
      Width           =   2115
   End
   Begin MSComctlLib.ListView lvCC 
      Height          =   2745
      Left            =   120
      TabIndex        =   0
      Top             =   1290
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   4842
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   8421504
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CC Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name On Card"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Expiry Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Security Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   120
      Picture         =   "frmCC_new.frx":6D05
      Stretch         =   -1  'True
      Top             =   120
      Width           =   930
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Credit Bank Card from this Subscriber."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   1110
      TabIndex        =   3
      Top             =   90
      Width           =   2655
   End
End
Attribute VB_Name = "frmCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CCID As Long


Private Sub Command1_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Command1_Click"
    Const ContainerName = "frmCC"
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


    Dim fCC As New frmCCAdd
    Dim rsCC As adodb.Recordset
    
    fCC.Show 1
    
    If fCC.ccName <> "" Then
        
        MySQL.Execute directConn, "Insert into creditcard (bType, CardNumber, Name, ExpiryDate, SecurityNumber) " + _
                                "VALUES (" & fCC.ccType & ", AES_ENCRYPT('" & MySQL.NumCrypt(fCC.ccNum) & "','" & odb.colSalts.ReturnSalt("CCSalt") & "'), " + _
                                "'" & MySQL.ESC(fCC.ccName) & "','" & Format(fCC.ccExpire, "yyyy-mm-dd") & "','" & fCC.ccSec & "')"
    
        lvCC.ListItems.Clear
        
        If MySQL.OpenTable(directConn, rsCC, , "select RecID, bType, AES_DECRYPT(CardNumber,'" + odb.colSalts.ReturnSalt("CCSalt") + "') as CardNumber, Name, ExpiryDate, SecurityNumber from creditcard") = True Then
            If Not rsCC.BOF And Not rsCC.EOF Then
                If rsCC.RecordCount > 0 Then
                    While Not rsCC.EOF And Err.Number = 0
                        Set itmX = lvCC.ListItems.Add(, "c" & rsCC!RecID, MySQL.NumDecrypt(rsCC!CardNumber))
                        itmX.SubItems(1) = rsCC!Name
                        itmX.SubItems(2) = Format(rsCC!ExpiryDate, "mm/yyyy")
                        itmX.SubItems(3) = rsCC!SecurityNumber
                            
                        Select Case Val(rsCC!bType)
                        Case 0
                            itmX.SubItems(4) = "EFTPOS"
                        Case 1
                            itmX.SubItems(4) = "Visa"
                        Case 2
                            itmX.SubItems(4) = "Master card"
                        Case 3
                            itmX.SubItems(4) = "American Express"
                        Case 4
                            itmX.SubItems(4) = "Dinners Club"
                        Case 5
                            itmX.SubItems(4) = "Discover"
                        Case 6
                            itmX.SubItems(4) = "JCB"
                        End Select
                        rsCC.MoveNext
                    Wend
                End If
            End If
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

Private Sub Command2_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Command2_Click"
    Const ContainerName = "frmCC"
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


    If lvCC.SelectedItem Is Nothing Then
        
        
    Else
        CCID = Val(Mid(lvCC.SelectedItem.Key, 2))
    
        Select Case MsgBox("You must now do a manual tranasction on this credit card:" & vbCrLf & vbCrLf & "Cc Number: " & lvCC.SelectedItem.Text & vbCrLf & "Cc Name: " & lvCC.SelectedItem.SubItems(1) & vbCrLf & "Cc Expiry: " & lvCC.SelectedItem.SubItems(2) & vbCrLf & "Cc Security: " & lvCC.SelectedItem.SubItems(3) & vbCrLf & "Cc Type: " & lvCC.SelectedItem.SubItems(4) & vbCrLf & vbCrLf & "Have you completed this transaction?", vbCritical + vbYesNo, "Manual Transaction")
        Case vbYes
            Unload Me
        Case vbNo
        
        End Select
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

Private Sub Form_Load()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Load"
    Const ContainerName = "frmCC"
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


    Call GUI.LoadColWidths(lvCC, Me)
    
    If bBigFont = True Then
        lvCC.Font.Size = 18
    End If

    Dim rsCC As adodb.Recordset
    Dim itmX As ListItem
    
    lvCC.ListItems.Clear
    
    If MySQL.OpenTable(directConn, rsCC, , "select RecID, bType, AES_DECRYPT(CardNumber,'" + odb.colSalts.ReturnSalt("CCSalt") + "') as CardNumber, Name, ExpiryDate, SecurityNumber from creditcard") = True Then
        If Not rsCC.BOF And Not rsCC.EOF Then
            If rsCC.RecordCount > 0 Then
                While Not rsCC.EOF And Err.Number = 0
                    Set itmX = lvCC.ListItems.Add(, "c" & rsCC!RecID, MySQL.NumDecrypt(rsCC!CardNumber))
                    itmX.SubItems(1) = rsCC!Name
                    itmX.SubItems(2) = Format(rsCC!ExpiryDate, "mm/yyyy")
                    itmX.SubItems(3) = rsCC!SecurityNumber
                        
                    Select Case Val(rsCC!bType)
                    Case 0
                        itmX.SubItems(4) = "EFTPOS"
                    Case 1
                        itmX.SubItems(4) = "Visa"
                    Case 2
                        itmX.SubItems(4) = "Master card"
                    Case 3
                        itmX.SubItems(4) = "American Express"
                    Case 4
                        itmX.SubItems(4) = "Dinners Club"
                    Case 5
                        itmX.SubItems(4) = "Discover"
                    Case 6
                        itmX.SubItems(4) = "JCB"
                    End Select
                    rsCC.MoveNext
                Wend
            End If
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
    Const ContainerName = "frmCC"
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


    Call GUI.SaveColWidths(lvCC, Me)
    
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

Private Sub lvCC_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvCC_ColumnClick"
    Const ContainerName = "frmCC"
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

    
    Call GUI.ColumnSort(ColumnHeader, lvCC)
    
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

Private Sub lvCC_DblClick()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvCC_DblClick"
    Const ContainerName = "frmCC"
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


    Call Command2_Click
    
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
