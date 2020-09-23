VERSION 5.00
Begin VB.Form frmBroadband 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Broadband Phone Number"
   ClientHeight    =   6480
   ClientLeft      =   4455
   ClientTop       =   3090
   ClientWidth     =   5520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   3105
      Index           =   1
      Left            =   90
      TabIndex        =   18
      Top             =   2880
      Width           =   5325
      Begin VB.CheckBox Check1 
         Caption         =   "Churn Exisiting DSL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2250
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2310
         Width           =   2895
      End
      Begin VB.TextBox txtAdd 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   1290
         MaxLength       =   100
         TabIndex        =   5
         Top             =   270
         Width           =   1035
      End
      Begin VB.ComboBox cmbType 
         Height          =   360
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   930
         Width           =   1785
      End
      Begin VB.TextBox txtAdd 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   2370
         MaxLength       =   100
         TabIndex        =   6
         Top             =   270
         Width           =   2835
      End
      Begin VB.TextBox txtAdd 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   180
         MaxLength       =   100
         TabIndex        =   4
         Top             =   270
         Width           =   1065
      End
      Begin VB.TextBox txtAdd 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   8
         Top             =   930
         Width           =   3165
      End
      Begin VB.TextBox txtAdd 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   6
         Left            =   180
         MaxLength       =   4
         TabIndex        =   11
         Top             =   2310
         Width           =   1905
      End
      Begin VB.TextBox txtAdd 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   4
         Left            =   180
         MaxLength       =   3
         TabIndex        =   9
         Top             =   1620
         Width           =   1905
      End
      Begin VB.TextBox txtAdd 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   5
         Left            =   2130
         MaxLength       =   50
         TabIndex        =   10
         Text            =   "Australia"
         Top             =   1620
         Width           =   3075
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Street No."
         Height          =   225
         Index           =   9
         Left            =   1290
         TabIndex        =   26
         Top             =   660
         Width           =   1035
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Street Name"
         Height          =   225
         Index           =   10
         Left            =   2400
         TabIndex        =   25
         Top             =   630
         Width           =   2805
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Unit Number"
         Height          =   225
         Index           =   8
         Left            =   180
         TabIndex        =   24
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Street Type"
         Height          =   225
         Index           =   7
         Left            =   180
         TabIndex        =   23
         Top             =   1320
         Width           =   1755
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Suburb"
         Height          =   225
         Index           =   3
         Left            =   2130
         TabIndex        =   22
         Top             =   1320
         Width           =   3045
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Postcode"
         Height          =   225
         Index           =   4
         Left            =   180
         TabIndex        =   21
         Top             =   2700
         Width           =   1905
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "State"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   20
         Top             =   2010
         Width           =   1905
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Country"
         Height          =   225
         Index           =   6
         Left            =   2130
         TabIndex        =   19
         Top             =   2010
         Width           =   3075
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Line Information"
      Height          =   2745
      Index           =   0
      Left            =   90
      TabIndex        =   13
      Top             =   60
      Width           =   5325
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   210
         MaxLength       =   3
         TabIndex        =   1
         Top             =   1170
         Width           =   1185
      End
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   1440
         MaxLength       =   8
         TabIndex        =   2
         Top             =   1170
         Width           =   3795
      End
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         Height          =   330
         Index           =   3
         Left            =   180
         TabIndex        =   3
         Top             =   2070
         Width           =   5025
      End
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   180
         MaxLength       =   255
         TabIndex        =   0
         Top             =   300
         Width           =   5055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Area Code"
         Height          =   255
         Left            =   390
         TabIndex        =   17
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Phone Number"
         Height          =   255
         Index           =   0
         Left            =   1590
         TabIndex        =   16
         Top             =   1680
         Width           =   3405
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "e-Mail address to notify of when line is broadband enabled."
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   2400
         Width           =   5085
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Phone Line Account Name"
         Height          =   255
         Index           =   2
         Left            =   210
         TabIndex        =   14
         Top             =   810
         Width           =   5025
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply For DSL Service"
      Height          =   345
      Left            =   90
      TabIndex        =   12
      Top             =   6060
      Width           =   5325
   End
End
Attribute VB_Name = "frmBroadband"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sName As String
Public sAreaCode As String
Public sPhoneNum As String
Public sEmail As String
Public UnitNumber As String
Public StreetNo As String
Public StreetName As String
Public StreetType As String
Public Suburb As String



Public State As String
Public Country As String
Public PostCode As String
Public Churn As Byte

Public iCloseState As frm_CloseStates

Private Sub Command1_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Command1_Click"
    Const ContainerName = "frmBroadband"
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


    Dim bx As Byte
    
    For bx = txtField.LBound + 1 To txtField.UBound
        If Trim(txtField(bx).Text) = "" Then bBlank = True
    Next
    
    If bBlank = True Then
        MsgBox "All Fields must be completed.", vbInformation, "Field Data Missing"
        Exit Sub
    End If
        
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

Private Sub Form_Load()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Load"
    Const ContainerName = "frmBroadband"
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


    If bDebug = True Then On Error GoTo 0 Else On Error GoTo ErrorOccur

    cmbType.AddItem "AVENUE"
    cmbType.AddItem "COURT"
    cmbType.AddItem "DRIVE"
    cmbType.AddItem "PLACE"
    cmbType.AddItem "ROAD"
    cmbType.AddItem "STREET"
    cmbType.AddItem "ALLEY"
    cmbType.AddItem "ARCADE"
    cmbType.AddItem "BEND"
    cmbType.AddItem "BOULEVARD"
    cmbType.AddItem "CHASE"
    cmbType.AddItem "CIRCLE"
    cmbType.AddItem "CIRCUIT"
    cmbType.AddItem "CLOSE"
    cmbType.AddItem "CRESCENT"
    cmbType.AddItem "ENTRANCE"
    cmbType.AddItem "ESPLANADE"
    cmbType.AddItem "FREEWAY"
    cmbType.AddItem "GARDENS"
    cmbType.AddItem "GLADE"
    cmbType.AddItem "GLEN"
    cmbType.AddItem "GROVE"
    cmbType.AddItem "HEIGHTS"
    cmbType.AddItem "HIGHWAY"
    cmbType.AddItem "HILL"
    cmbType.AddItem "KEY"
    cmbType.AddItem "LANE"
    cmbType.AddItem "LOOP"
    cmbType.AddItem "MALL"
    cmbType.AddItem "MEWS"
    cmbType.AddItem "PARADE"
    cmbType.AddItem "PROMENADE"
    cmbType.AddItem "RETREAT"
    cmbType.AddItem "RISE"
    cmbType.AddItem "ROW"
    cmbType.AddItem "SQUARE"
    cmbType.AddItem "TERRACE"
    cmbType.AddItem "TRACK"
    cmbType.AddItem "TRAIL"
    cmbType.AddItem "WALK"
    cmbType.AddItem "WAY"
    cmbType.AddItem "WYND"

    txtField(2) = sName
    txtField(0) = sAreaCode
    txtField(1) = sPhoneNum
    txtField(3) = sEmail
    
    txtAdd(0) = Me.UnitNumber
    txtAdd(1) = Me.StreetNo
    txtAdd(2) = Me.StreetName
    If Me.StreetType <> "" Then cmbType.Text = Me.StreetType
    txtAdd(3) = Me.Suburb
    txtAdd(4) = Me.State
    txtAdd(5) = Me.Country
    txtAdd(6) = Me.PostCode
    
    If Err.Number = 0 Then Exit Sub
    
    
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
    Const ContainerName = "frmBroadband"
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

    
    If bDebug = True Then On Error GoTo 0 Else On Error GoTo ErrorOccur
    
    sName = txtField(2)
    sAreaCode = txtField(0)
    sPhoneNum = txtField(1)
    sEmail = txtField(3)

    Me.UnitNumber = txtAdd(0)
    Me.StreetNo = txtAdd(1)
    Me.StreetName = txtAdd(2)
    Me.StreetType = cmbType.Text
    Me.Suburb = txtAdd(3)
    Me.State = txtAdd(4)
    Me.Country = txtAdd(5)
    Me.PostCode = txtAdd(6)
    Me.Churn = Check1.Value
    
    If iCloseState <> frmCloseSave Then iCloseState = frmCloseCancel
    
    If Err.Number = 0 Then Exit Sub
    
    
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
    Const ContainerName = "frmBroadband"
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
    txtField(Index).SelLength = Len(txtField(Index))
    
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
    Const ContainerName = "frmBroadband"
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
    Case 13
        KeyAscii = 0
        SendKeys "{TAB}"
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
