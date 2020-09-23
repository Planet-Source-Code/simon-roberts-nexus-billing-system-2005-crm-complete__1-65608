VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRadiusAccount 
   BackColor       =   &H006977A7&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Network Access Account (Radius Server)"
   ClientHeight    =   9735
   ClientLeft      =   1935
   ClientTop       =   2625
   ClientWidth     =   9315
   ClipControls    =   0   'False
   Icon            =   "frmRadiusAccount2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   649
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   621
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picTS 
      BackColor       =   &H006977A7&
      BorderStyle     =   0  'None
      Height          =   8715
      Left            =   150
      ScaleHeight     =   8715
      ScaleWidth      =   9045
      TabIndex        =   18
      Top             =   90
      Width           =   9045
      Begin VB.Frame Frame1 
         BackColor       =   &H006977A7&
         Caption         =   "Sessions Allowed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0084E8E8&
         Height          =   765
         Index           =   2
         Left            =   3600
         TabIndex        =   20
         Top             =   1920
         Width           =   2295
         Begin VB.TextBox txtSessions 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   150
            Locked          =   -1  'True
            TabIndex        =   9
            Text            =   "1"
            Top             =   300
            Width           =   1755
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   315
            Left            =   1950
            TabIndex        =   10
            Top             =   300
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtContact"
            BuddyDispid     =   196625
            OrigLeft        =   2970
            OrigTop         =   300
            OrigRight       =   3210
            OrigBottom      =   615
            Enabled         =   -1  'True
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H006977A7&
         Caption         =   "Radius Group"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0084E8E8&
         Height          =   855
         Left            =   3570
         TabIndex        =   30
         Top             =   3180
         Width           =   5385
         Begin VB.ComboBox cmbGroupname 
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
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   330
            Width           =   5145
         End
      End
      Begin VB.PictureBox picDatawave 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   4395
         Left            =   90
         ScaleHeight     =   289
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   587
         TabIndex        =   29
         Tag             =   "Data wave graph mesuring the history of data usage over time of the account creatation"
         Top             =   4200
         Width           =   8865
         Begin VB.Line Line4 
            BorderColor     =   &H00FF00FF&
            BorderStyle     =   3  'Dot
            DrawMode        =   4  'Mask Not Pen
            X1              =   0
            X2              =   588
            Y1              =   138
            Y2              =   138
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H006977A7&
         Caption         =   "Timeouts"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0084E8E8&
         Height          =   1275
         Index           =   3
         Left            =   6030
         TabIndex        =   26
         Top             =   1890
         Width           =   2925
         Begin VB.TextBox txtIdleTimeout 
            Height          =   360
            Left            =   1590
            TabIndex        =   13
            Text            =   "600"
            Top             =   750
            Width           =   975
         End
         Begin VB.TextBox txtSessionTimeout 
            Height          =   360
            Left            =   1590
            TabIndex        =   11
            Text            =   "10800"
            Top             =   300
            Width           =   975
         End
         Begin MSComCtl2.UpDown UpDown2 
            Height          =   360
            Left            =   2580
            TabIndex        =   14
            Top             =   750
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   635
            _Version        =   393216
            Value           =   600
            BuddyControl    =   "txtIdleTimeout"
            BuddyDispid     =   196616
            OrigLeft        =   2880
            OrigTop         =   750
            OrigRight       =   3120
            OrigBottom      =   1110
            Max             =   999999999
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udSessionTimeout 
            Height          =   360
            Left            =   2580
            TabIndex        =   12
            Top             =   300
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   635
            _Version        =   393216
            Value           =   10800
            BuddyControl    =   "txtSessionTimeout"
            BuddyDispid     =   196617
            OrigLeft        =   2850
            OrigTop         =   300
            OrigRight       =   3090
            OrigBottom      =   660
            Max             =   999999999
            Min             =   -1
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Idle Timeout:"
            Height          =   240
            Index           =   1
            Left            =   480
            TabIndex        =   28
            Top             =   780
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Session Timeout:"
            Height          =   240
            Index           =   0
            Left            =   210
            TabIndex        =   27
            Top             =   360
            Width           =   1245
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H006977A7&
         Caption         =   "Autoactivate/deactivate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0084E8E8&
         Height          =   2115
         Left            =   90
         TabIndex        =   23
         Top             =   1920
         Width           =   3375
         Begin VB.OptionButton optAutoTime 
            BackColor       =   &H006977A7&
            Caption         =   "Not Set"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   0
            Left            =   180
            TabIndex        =   4
            Top             =   270
            Value           =   -1  'True
            Width           =   2775
         End
         Begin VB.OptionButton optAutoTime 
            BackColor       =   &H006977A7&
            Caption         =   "Set Interval:"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   2
            Left            =   180
            TabIndex        =   6
            Top             =   810
            Width           =   2775
         End
         Begin VB.OptionButton optAutoTime 
            BackColor       =   &H006977A7&
            Caption         =   "Business Hours: 8:30am - 5:30pm"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   1
            Left            =   180
            TabIndex        =   5
            Top             =   540
            Width           =   2775
         End
         Begin MSComCtl2.DTPicker dtpActivation 
            Height          =   375
            Left            =   1440
            TabIndex        =   7
            Top             =   1080
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "HH:mm:ss"
            Format          =   68681731
            UpDown          =   -1  'True
            CurrentDate     =   0.354166666666667
         End
         Begin MSComCtl2.DTPicker dtpDeactivation 
            Height          =   375
            Left            =   1440
            TabIndex        =   8
            Top             =   1560
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "HH:mm:ss"
            Format          =   68681731
            UpDown          =   -1  'True
            CurrentDate     =   0.729166666666667
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H006977A7&
            Caption         =   "Deactivate:"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   25
            Top             =   1620
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H006977A7&
            Caption         =   "Activate Time:"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   24
            Top             =   1170
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H006977A7&
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0084E8E8&
         Height          =   885
         Index           =   0
         Left            =   90
         TabIndex        =   22
         Top             =   930
         Width           =   4305
         Begin VB.CommandButton cmdUsername 
            Caption         =   "..."
            Height          =   465
            Left            =   3690
            TabIndex        =   2
            Top             =   270
            Width           =   465
         End
         Begin VB.TextBox txtUID 
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
            Height          =   465
            Left            =   150
            MaxLength       =   50
            TabIndex        =   1
            Top             =   270
            Width           =   3465
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H006977A7&
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0084E8E8&
         Height          =   885
         Index           =   1
         Left            =   4500
         TabIndex        =   21
         Top             =   930
         Width           =   4455
         Begin VB.TextBox txtPWD 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   50
            PasswordChar    =   "¤"
            TabIndex        =   3
            Top             =   270
            Width           =   4245
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H006977A7&
         Caption         =   "Contact Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0084E8E8&
         Height          =   855
         Index           =   4
         Left            =   90
         TabIndex        =   19
         Top             =   60
         Width           =   8865
         Begin VB.TextBox txtContact 
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
            Height          =   465
            Left            =   150
            MaxLength       =   50
            TabIndex        =   0
            Top             =   270
            Width           =   8595
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H006977A7&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7140
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8970
      Width           =   2085
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H006977A7&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8910
      Width           =   2235
   End
End
Attribute VB_Name = "frmRadiusAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lRadiusID As Variant
Public p_ContactName As String
Public p_Username As String
Public p_Password As String
Public p_Sessions As Integer
Public p_AutoFlag As Byte
Public p_Activate As Date
Public p_Deactivate As Date
Public p_SessionTimeOut As Variant
Public p_IdleTimeout As Variant

Public iCloseState As frm_CloseStates
        

Private Sub cmdSave_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdSave_Click"
    Const ContainerName = "frmRadiusAccount"
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


    Dim bResult As Boolean
    Dim rsRadius As ADODB.Recordset
    Dim rsload As ADODB.Recordset
    Static FirstDone As Boolean
    
    If Me.lRadiusID <> 0 Then
        bResult = MySQL.OpenTable(ADOConn, rsload, , "select * from radiusaccounts where RecID = " & Me.lRadiusID & " Limit 1 ")
        If txtUID <> rsload!Username Then
            bResult = MySQL.OpenTable(ADOConn, rsRadius, , "select * from radiusaccounts where Username = '" & txtUID & "' Limit 1 ")
        
            If rsRadius.RecordCount > 0 Then
                MsgBox "Username " & txtUID & " already exists in system."
                If FirstDone = False Then
                    FirstDone = True
                    txtUID.Tag = txtUID.Text
                    txtUID.Text = txtUID.Tag & Round(Rnd * 255)
                    Exit Sub
                Else
                    txtUID.Text = txtUID.Tag & Round(Rnd * 255)
                End If
                
                Exit Sub
            End If
        End If
    Else
    
        bResult = MySQL.OpenTable(ADOConn, rsRadius, , "select * from radiusaccounts where Username = '" & txtUID & "' Limit 1 ")
    
        If rsRadius.RecordCount > 0 Then
            MsgBox "Username " & txtUID & " already exists in system."
                If FirstDone = False Then
                    FirstDone = True
                    txtUID.Tag = txtUID.Text
                    txtUID.Text = txtUID.Tag & Round(Rnd * 255)
                    Exit Sub
                Else
                    txtUID.Text = txtUID.Tag & Round(Rnd * 255)
                End If
        End If
    End If

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

Private Sub cmdCancel_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdCancel_Click"
    Const ContainerName = "frmRadiusAccount"
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

Private Sub Command1_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Command1_Click"
    Const ContainerName = "frmRadiusAccount"
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

Private Sub cmdUsername_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdUsername_Click"
    Const ContainerName = "frmRadiusAccount"
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


    Dim fUser As New frmUsername
    fUser.Show 1
    If fUser.sUsername <> "" Then txtUID = fUser.sUsername
        
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
    Const ContainerName = "frmRadiusAccount"
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

    If Me.p_ContactName <> "" Then txtContact.Text = Me.p_ContactName
    If Me.p_Password <> "" Then txtPWD.Text = Me.p_Password
    If Me.p_Username <> "" Then txtUID.Text = Me.p_Username
    
    If lRadiusID <> 0 Then
        LoadRadiusInfo lRadiusID
        DrawDatawave picDatawave, , CLng(lRadiusID)
    End If
    
    If Login.bMaster = False Then
        txtSessions.Enabled = False
        UpDown1.Enabled = False
        txtSessionTimeout.Enabled = False
        txtIdleTimeout.Enabled = False
        UpDown2.Enabled = False
        udSessionTimeout.Enabled = False
        cmbGroupname.Enabled = False
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

Public Function LoadRadiusInfo(lRadiusID As Variant)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "LoadRadiusInfo"
    Const ContainerName = "frmRadiusAccount"
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


    Dim bResult As Boolean
    Dim rsload As ADODB.Recordset
    
    bResult = MySQL.OpenTable(ADOConn, rsload, , "select * from radiusaccounts Where RecID = " & lRadiusID)
    
    If bResult = False Then Exit Function
    
    If rsload.RecordCount > 0 Then
        txtContact = p_ContactName
        txtUID = IIf(IsNull(rsload!Username), "", rsload!Username)
        txtUID.Tag = IIf(IsNull(rsload!Username), "", rsload!Username)
        txtUID.Locked = True
        'txtPWD = IIf(IsNull(rsLoad!Password), "", rsLoad!Password)
        txtSessions = rsload!SessionsAllowed
        dtpActivation.Value = IIf(IsNull(rsload!Activate), sysnow, rsload!Activate)
        dtpDeactivation.Value = IIf(IsNull(rsload!Deactivate), sysnow, rsload!Deactivate)
        txtSessionTimeout.Text = "" & Val(rsload!SessionTimeout)
        txtIdleTimeout.Text = "" & Val(rsload!IdleTimeout)
    
        Dim rsList As ADODB.Recordset
        Dim rsload2 As ADODB.Recordset
        
        If MySQL.OpenTable(ADOConn, rsload2, , "select * from radius.RadiusUserGroup where radius.RadiusUserGroup.username = '" & txtUID & "'") = True Then
            If rsload2.State = 0 Then
                            
            Else
                If rsload2.RecordCount = 0 Then cmbGroupname.Enabled = False
                If MySQL.OpenTable(ADOConn, rsList, , "select distinct groupname from radius.radiusgroupreply") = True Then
                    If rsload.RecordCount > 0 Then
                        If rsList.RecordCount > 0 Then
                            Do
                                cmbGroupname.AddItem rsList!groupname
                                If rsload2.RecordCount > 0 Then If rsList!groupname = rsload2!groupname Then cmbGroupname.ListIndex = cmbGroupname.ListCount - 1
                                    
                                rsList.MoveNext
                            Loop Until rsList.EOF Or Err.Number <> 0
                        End If
                    Else
                        If rsList.RecordCount > 0 Then
                            Do
                                cmbGroupname.AddItem rsList!groupname
                                rsList.MoveNext
                            Loop Until rsList.EOF Or Err.Number <> 0
                        End If
                    End If
                End If
            End If
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_QueryUnload"
    Const ContainerName = "frmRadiusAccount"
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


    If cmbGroupname.ListCount > 0 And cmbGroupname.ListIndex <> -1 Then
        
        MySQL.Execute ADOConn, "Update radius.RadiusUserGroup Set groupname='" & cmbGroupname.Text & "' where Username = '" & txtUID.Tag & "'"
        
    End If
    
    Me.p_Username = txtUID
    Me.p_Password = txtPWD
    Me.p_ContactName = txtContact
    Me.p_Sessions = Val(txtSessions)
    Me.p_Activate = dtpActivation.Value
    Me.p_Deactivate = dtpDeactivation.Value
    Me.p_SessionTimeOut = Val(txtSessionTimeout.Text)
    Me.p_IdleTimeout = Val(txtIdleTimeout.Text)
    
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

Private Sub optAutoTime_Click(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "optAutoTime_Click"
    Const ContainerName = "frmRadiusAccount"
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


    Select Case Index
    Case 0, 1
        dtpActivation.Enabled = False
        dtpDeactivation.Enabled = False
        If Index = 1 Then
            dtpActivation.Value = CDate("8:30am")
            dtpDeactivation.Value = CDate("5:30pm")
        End If
    Case 2
        dtpActivation.Enabled = True
        dtpDeactivation.Enabled = True
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

Private Sub txtIdleTimeout_KeyPress(KeyAscii As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtIdleTimeout_KeyPress"
    Const ContainerName = "frmRadiusAccount"
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


    Select Case KeyAscii
    Case 8, 48 To 57
    Case Else
        KeyAscii = 0
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

Private Sub txtSessions_KeyPress(KeyAscii As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtSessions_KeyPress"
    Const ContainerName = "frmRadiusAccount"
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


    Select Case KeyAscii
    Case 8, 48 To 57
    Case Else
        KeyAscii = 0
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

Private Sub txtSessionTimeout_KeyPress(KeyAscii As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtSessionTimeout_KeyPress"
    Const ContainerName = "frmRadiusAccount"
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


    Select Case KeyAscii
    Case 8, 48 To 57
    Case Else
        KeyAscii = 0
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

Private Sub txtUID_KeyPress(KeyAscii As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtUID_KeyPress"
    Const ContainerName = "frmRadiusAccount"
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


    Select Case KeyAscii
    Case Asc(" ")
        KeyAscii = 0
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
