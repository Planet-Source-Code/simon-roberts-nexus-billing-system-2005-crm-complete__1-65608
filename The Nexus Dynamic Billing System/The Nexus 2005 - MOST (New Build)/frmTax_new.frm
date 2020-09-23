VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmTax 
   Caption         =   "Tax"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10065
   Icon            =   "frmTax_new.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   10065
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   90
      TabIndex        =   2
      Top             =   750
      Width           =   9825
      Begin VB.CommandButton Command1 
         Caption         =   "+"
         Height          =   345
         Left            =   9330
         TabIndex        =   7
         Top             =   300
         Width           =   375
      End
      Begin VB.TextBox txtPercentage 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   6870
         TabIndex        =   5
         Text            =   "0"
         Top             =   270
         Width           =   2385
      End
      Begin VB.TextBox txtDesc 
         Height          =   375
         Left            =   150
         TabIndex        =   3
         Top             =   270
         Width           =   6675
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Percentage"
         Height          =   255
         Left            =   6870
         TabIndex        =   6
         Top             =   660
         Width           =   2385
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Description"
         Height          =   225
         Left            =   150
         TabIndex        =   4
         Top             =   660
         Width           =   6705
      End
   End
   Begin MSComctlLib.ListView lvTAX 
      Height          =   3765
      Left            =   90
      TabIndex        =   1
      Top             =   1800
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   6641
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Description"
         Object.Width           =   14473
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Percentage"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tax"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   720
      TabIndex        =   0
      Top             =   30
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   30
      Picture         =   "frmTax_new.frx":0ECA
      Top             =   30
      Width           =   720
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   720
      X2              =   10080
      Y1              =   360
      Y2              =   390
   End
End
Attribute VB_Name = "frmTax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Load"
    Const ContainerName = "frmTax"
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



    Dim bResult As Boolean
    Dim rstax As adodb.Recordset
    
    
    bResult = MySQL.OpenTable(directConn, rstax, , "select * from Tax")
    
    If rstax.RecordCount > 0 Then
        
        Dim itmX As ListItem
        While Not rstax.EOF And Err.Number = 0
            
            Set itmX = lvTAX.ListItems.Add(, "t" & rstax!RecID, rstax!Description)
            itmX.SubItems(1) = rstax!Percentage
            
        Wend
        
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

