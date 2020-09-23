VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSysopDetails 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "System Operator"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   3120
      TabIndex        =   6
      Top             =   5880
      Width           =   1425
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save And Close"
      Height          =   345
      Left            =   90
      TabIndex        =   5
      Top             =   5880
      Width           =   1425
   End
   Begin VB.PictureBox picContents 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   150
      ScaleHeight     =   4575
      ScaleWidth      =   4305
      TabIndex        =   8
      Top             =   1110
      Width           =   4305
      Begin MSComctlLib.ImageList ilSysops 
         Left            =   300
         Top             =   2010
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
               Picture         =   "frmSysopDetails.frx":0000
               Key             =   "k100"
               Object.Tag             =   "100% Security Access"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSysopDetails.frx":031A
               Key             =   "k080_099"
               Object.Tag             =   "High Level Access"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSysopDetails.frx":0634
               Key             =   "k060_079"
               Object.Tag             =   "Server Administrator"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSysopDetails.frx":0F0E
               Key             =   "k040_059"
               Object.Tag             =   "Network Admin"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSysopDetails.frx":1360
               Key             =   "k001_020"
               Object.Tag             =   "Service Plans Editor"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSysopDetails.frx":17B2
               Key             =   "k020_039"
               Object.Tag             =   "Rainbow Warrior"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSysopDetails.frx":1C04
               Key             =   "k000"
               Object.Tag             =   "Low Level Access"
            EndProperty
         EndProperty
      End
      Begin VB.Frame frmUserdetails 
         Caption         =   "User Details"
         Height          =   3255
         Left            =   60
         TabIndex        =   11
         Top             =   1290
         Width           =   4215
         Begin VB.TextBox txtField 
            BorderStyle     =   0  'None
            Height          =   1980
            Index           =   0
            Left            =   1710
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   2
            Top             =   390
            Width           =   2325
         End
         Begin VB.TextBox txtField 
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   1
            Left            =   1710
            TabIndex        =   3
            Top             =   2460
            Width           =   2325
         End
         Begin VB.TextBox txtField 
            BorderStyle     =   0  'None
            Height          =   270
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   1710
            PasswordChar    =   "#"
            TabIndex        =   4
            Top             =   2820
            Width           =   2325
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Description:"
            Height          =   240
            Left            =   210
            TabIndex        =   14
            Top             =   390
            Width           =   1350
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Username:"
            Height          =   240
            Left            =   210
            TabIndex        =   13
            Top             =   2490
            Width           =   1350
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Password:"
            Height          =   240
            Left            =   810
            TabIndex        =   12
            Top             =   2850
            Width           =   765
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Securtity Level"
         Height          =   1185
         Left            =   60
         TabIndex        =   9
         Top             =   90
         Width           =   4215
         Begin MSComctlLib.Slider sldSecLevel 
            Height          =   465
            Left            =   600
            TabIndex        =   1
            Top             =   570
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   820
            _Version        =   393216
            Max             =   100
            TickFrequency   =   10
         End
         Begin VB.Label lblDescription 
            AutoSize        =   -1  'True
            Height          =   240
            Left            =   720
            TabIndex        =   10
            Top             =   330
            Width           =   45
         End
         Begin VB.Image imgSecurityLevel 
            Height          =   480
            Left            =   150
            Picture         =   "frmSysopDetails.frx":2056
            Top             =   270
            Width           =   480
         End
      End
   End
   Begin MSComctlLib.TabStrip tsContents 
      Height          =   5055
      Left            =   90
      TabIndex        =   7
      Top             =   750
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   8916
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Account Details"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sysop Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1020
      TabIndex        =   0
      Top             =   30
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   60
      Picture         =   "frmSysopDetails.frx":2498
      Stretch         =   -1  'True
      Top             =   0
      Width           =   810
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   4710
      X2              =   0
      Y1              =   330
      Y2              =   330
   End
End
Attribute VB_Name = "frmSysopDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public byLevel As Byte
Public sDescription As String
Public sUsername As String
Public sPassword As String
Public iCloseState As frm_CloseStates


Private Sub Command1_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Command1_Click"
    Const ContainerName = "frmSysopDetails"
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
    For bx = 0 To 2
        If Trim(txtField(bx)) = "" Then
            MsgBox "All three fields must be completed before saving Sysop Details", vbInformation, "Details Missing"
            Exit Sub
        End If
    Next
    
    iCloseState = frmCloseSave
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

Private Sub Command2_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Command2_Click"
    Const ContainerName = "frmSysopDetails"
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


    iCloseState = frmCloseCancel
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

Private Sub Form_Load()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Load"
    Const ContainerName = "frmSysopDetails"
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


    
    If byLevel > Login.lLevel Then
    
        txtField(0).Locked = True
        txtField(1).Locked = True
        txtField(2).Locked = True
        sldSecLevel.Enabled = False
        
    End If
    
    sldSecLevel.Value = byLevel
    txtField(0) = sDescription
    txtField(1) = sUsername
    txtField(2) = sPassword
    
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_QueryUnload"
    Const ContainerName = "frmSysopDetails"
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


    byLevel = sldSecLevel.Value
    sDescription = txtField(0)
    sUsername = txtField(1)
    sPassword = txtField(2)
    
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

Private Sub sldSecLevel_Change()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "sldSecLevel_Change"
    Const ContainerName = "frmSysopDetails"
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


    Dim imgX As Byte
    Dim imgMin As Byte
    Dim imgMax As Byte
        
    If sldSecLevel.Enabled = True Then
        If sldSecLevel.Value > Login.lLevel Then sldSecLevel.Value = Login.lLevel
    End If
    
    For imgX = 1 To ilSysops.ListImages.Count
        If InStr(ilSysops.ListImages(imgX).Key, "_") > 0 Then
            imgMin = CByte(Mid(ilSysops.ListImages(imgX).Key, 2, 3))
            imgMax = CByte(Mid(ilSysops.ListImages(imgX).Key, 6, 3))
        Else
            imgMin = CByte(Mid(ilSysops.ListImages(imgX).Key, 2, 3))
            imgMax = imgMin
        End If
        If sldSecLevel.Value >= imgMin And sldSecLevel.Value <= imgMax Then
            imgSecurityLevel.Picture = ilSysops.ListImages(imgX).Picture
            lblDescription.Caption = "" & sldSecLevel.Value & "% - " & ilSysops.ListImages(imgX).Tag
            Exit For
        End If
    Next imgX
    
    
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

Private Sub Slider1_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Slider1_Click"
    Const ContainerName = "frmSysopDetails"
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



ErrorOccur:
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Sub

