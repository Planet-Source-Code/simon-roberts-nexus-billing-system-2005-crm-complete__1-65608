VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCCAdd 
   BackColor       =   &H00F3B4E3&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Credit Card"
   ClientHeight    =   5505
   ClientLeft      =   615
   ClientTop       =   1545
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCCAdd.frx":0000
   ScaleHeight     =   5505
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00F3B4E3&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4740
      Width           =   2265
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00F3B4E3&
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4740
      Width           =   2265
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F3B4E3&
      Caption         =   "Credit Card"
      Height          =   4485
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   5775
      Begin VB.Frame Frame3 
         BackColor       =   &H00F3B4E3&
         Caption         =   "Type"
         Height          =   915
         Left            =   120
         TabIndex        =   11
         Top             =   3360
         Width           =   5445
         Begin VB.ComboBox cmbType 
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
            TabIndex        =   12
            Top             =   330
            Width           =   5145
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00F3B4E3&
         Caption         =   "Name On Card"
         Height          =   975
         Left            =   120
         TabIndex        =   7
         Top             =   270
         Width           =   5505
         Begin VB.TextBox txtfield 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   150
            TabIndex        =   8
            Top             =   300
            Width           =   5205
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00F3B4E3&
         Caption         =   "Card Number"
         Height          =   945
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   5505
         Begin VB.TextBox txtfield 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Index           =   1
            Left            =   180
            MaxLength       =   25
            TabIndex        =   6
            Top             =   300
            Width           =   5175
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00F3B4E3&
         Caption         =   "Secuirty Number (CIC)"
         Height          =   945
         Left            =   120
         TabIndex        =   3
         Top             =   2340
         Width           =   2325
         Begin VB.TextBox txtfield 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   2
            Left            =   150
            TabIndex        =   4
            Top             =   300
            Width           =   2025
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00F3B4E3&
         Caption         =   "Expiry Date"
         Height          =   975
         Left            =   2520
         TabIndex        =   1
         Top             =   2310
         Width           =   3075
         Begin MSComCtl2.DTPicker dtpExpiry 
            Height          =   525
            Left            =   90
            TabIndex        =   2
            Top             =   300
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   926
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   57409537
            CurrentDate     =   37629
         End
      End
   End
End
Attribute VB_Name = "frmCCAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ccName As String
Public ccNum As String
Public ccSec As String
Public ccExpire As Date
Public ccType As Byte

Private Sub cmdAdd_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdAdd_Click"
    Const ContainerName = "frmCCAdd"
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


    If cmbType.ListIndex = -1 Then
        MsgBox "You must select a type of Card"
        Exit Sub
    End If
    
    If txtField(0) = "" And txtField(1) = "" Then
        MsgBox "You must enter a name and number for the card"
        Exit Sub
    End If
    
    ccName = txtField(0)
    ccNum = txtField(1)
    ccSec = txtField(2)
    ccExpire = dtpExpiry.Value
    ccType = cmbType.ListIndex
    
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

Private Sub cmdClose_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdClose_Click"
    Const ContainerName = "frmCCAdd"
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

Private Sub Form_Load()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Load"
    Const ContainerName = "frmCCAdd"
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


    cmbType.AddItem "EFTPOS"
    cmbType.AddItem "Visa"
    cmbType.AddItem "Mastercard"
    cmbType.AddItem "American Express"
    cmbType.AddItem "Dinners Club"
    cmbType.AddItem "Discover"
    cmbType.AddItem "JCB"
    
    If bBigFont = True Then
    '    cmbtype.
    
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
