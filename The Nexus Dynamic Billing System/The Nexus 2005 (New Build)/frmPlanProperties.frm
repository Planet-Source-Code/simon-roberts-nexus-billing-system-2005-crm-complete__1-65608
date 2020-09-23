VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPlanProperties 
   BackColor       =   &H00A9C9AE&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   9825
   ClientLeft      =   645
   ClientTop       =   1230
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9825
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      BackColor       =   &H00A9C9AE&
      Caption         =   "&Close"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   9180
      Width           =   3675
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00A9C9AE&
      Caption         =   "&Update"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   8670
      Width           =   3675
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00A9C9AE&
      Caption         =   "Other plans included with this service"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4185
      Left            =   3900
      TabIndex        =   32
      Top             =   5520
      Width           =   7695
      Begin MSComctlLib.ListView lvPlans 
         Height          =   3795
         Left            =   120
         TabIndex        =   35
         Top             =   300
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   6694
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         PictureAlignment=   1
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Contact Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Username"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "MD5 Encryption CRC"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Base URL"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Object.Width           =   2540
         EndProperty
         Picture         =   "frmPlanProperties.frx":0000
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00004080&
      Caption         =   "MB Quota"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   795
      Index           =   5
      Left            =   3870
      TabIndex        =   28
      Top             =   2940
      Width           =   3975
      Begin VB.TextBox txtxField 
         Alignment       =   2  'Center
         Height          =   345
         Index           =   5
         Left            =   90
         TabIndex        =   29
         Top             =   300
         Width           =   3765
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00004080&
      Caption         =   "Service Active"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   765
      Index           =   4
      Left            =   8010
      TabIndex        =   27
      Top             =   2970
      Width           =   3525
      Begin VB.OptionButton optActive 
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   1860
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   270
         Width           =   1545
      End
      Begin VB.OptionButton optActive 
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   270
         Value           =   -1  'True
         Width           =   1635
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00004080&
      Caption         =   "Cost for Plan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1605
      Left            =   3840
      TabIndex        =   20
      Top             =   3780
      Width           =   7695
      Begin VB.TextBox txtFee 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   5460
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   270
         Width           =   2085
      End
      Begin VB.TextBox txtFee 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   5460
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   690
         Width           =   2085
      End
      Begin VB.TextBox txtFee 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   5460
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1110
         Width           =   2085
      End
      Begin VB.TextBox txtFee 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   3270
         TabIndex        =   23
         Top             =   270
         Width           =   2145
      End
      Begin VB.TextBox txtFee 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3270
         TabIndex        =   22
         Top             =   690
         Width           =   2145
      End
      Begin VB.TextBox txtFee 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   3270
         TabIndex        =   21
         Top             =   1110
         Width           =   2145
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00004080&
         Caption         =   "Cost/Cycle Price (Excluding Tax):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0084E8E8&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   26
         Top             =   330
         Width           =   2880
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00004080&
         Caption         =   "Per Mb Block (Excluding Tax):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0084E8E8&
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   25
         Top             =   750
         Width           =   2595
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00004080&
         Caption         =   "Per Extra Hour (Excluding Tax):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0084E8E8&
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   24
         Top             =   1170
         Width           =   2700
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00004080&
      Caption         =   "Date Created"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   795
      Index           =   3
      Left            =   3870
      TabIndex        =   18
      Top             =   2100
      Width           =   3975
      Begin VB.TextBox txtxField 
         Alignment       =   2  'Center
         Height          =   345
         Index           =   3
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   300
         Width           =   3765
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00004080&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   765
      Index           =   2
      Left            =   8010
      TabIndex        =   16
      Top             =   2160
      Width           =   3525
      Begin VB.TextBox txtxField 
         Alignment       =   2  'Center
         Height          =   345
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   300
         Width           =   3285
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00004080&
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   795
      Index           =   1
      Left            =   8010
      TabIndex        =   14
      Top             =   1290
      Width           =   3555
      Begin VB.TextBox txtxField 
         Alignment       =   2  'Center
         Height          =   345
         Index           =   1
         Left            =   150
         TabIndex        =   15
         Top             =   300
         Width           =   3315
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00004080&
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   795
      Index           =   0
      Left            =   3870
      TabIndex        =   12
      Top             =   1260
      Width           =   3975
      Begin VB.TextBox txtxField 
         Alignment       =   2  'Center
         Height          =   345
         Index           =   0
         Left            =   90
         TabIndex        =   13
         Top             =   300
         Width           =   3765
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0043D143&
      Caption         =   "Previous Cycle"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Index           =   1
      Left            =   90
      TabIndex        =   9
      Top             =   4980
      Width           =   3705
      Begin MSComCtl2.DTPicker dtpCycle 
         Height          =   375
         Index           =   1
         Left            =   180
         TabIndex        =   10
         Top             =   3150
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   63832066
         UpDown          =   -1  'True
         CurrentDate     =   37795
      End
      Begin MSComCtl2.MonthView mvCycle 
         Height          =   2820
         Index           =   1
         Left            =   180
         TabIndex        =   11
         Top             =   300
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   4974
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
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
         StartOfWeek     =   63832065
         CurrentDate     =   37795
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0094BD7B&
      Caption         =   "Next Cycle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3645
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   1230
      Width           =   3735
      Begin MSComCtl2.DTPicker dtpCycle 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   3180
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   63832066
         UpDown          =   -1  'True
         CurrentDate     =   37795
      End
      Begin MSComCtl2.MonthView mvCycle 
         Height          =   2820
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   330
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   4974
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
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
         StartOfWeek     =   63832065
         CurrentDate     =   37795
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Header Descriptions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   11415
      Begin VB.Label lblPLan 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   5310
         TabIndex        =   2
         Top             =   690
         Width           =   735
      End
      Begin VB.Label lblPLan 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   5310
         TabIndex        =   1
         Top             =   300
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmPlanProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public acciRecID As Long
Public RecID As Long
Dim rsload As ADODB.Recordset
Dim rsPlans As ADODB.Recordset
Dim rsCosts As ADODB.Recordset
Dim rsSysops As ADODB.Recordset

Private Sub Command1_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Command1_Click"
    Const ContainerName = "frmPlanProperties"
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


'If bDebug = True Then On Error GoTo 0 Else On Error GoTo ErrorOccur

    MySQL.Execute ADOConn, "Update acci_services SET " + _
                            "ContactName = '" & MySQL.ESC(txtxField(0).Text) & "', " + _
                            "Username = '" & MySQL.ESC(txtxField(1).Text) & "', " + _
                            "Password = AES_ENCRYPT('" & MySQL.ESC(txtxField(2).Text) & "','" & odb.colSalts.ReturnSalt("md5Password") & "'), " + _
                            "MBQuota = " & Val(txtxField(5).Text) & ", " & _
                            "NextCycle = '" & Format(mvCycle(0).Value, "YYYY-MM-DD") & " " & Format(dtpCycle(0).Value, "Hh:Nn:Ss") & "', " & _
                            "PeriodFee = " & txtFee(1).Text & ", " & _
                            "PerMB = " & txtFee(2) & ", " & _
                            "PerHour = " & txtFee(3) & ", " & _
                            "Checked = " & IIf(optActive(0).Value = True, -1, 0) & " " & _
                            " WHERE RecID = " & RecID
                            

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
    Const ContainerName = "frmPlanProperties"
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
    Const ContainerName = "frmPlanProperties"
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

    
    
    Me.Caption = "Customer " & acciRecID & " - Plan " & RecID
    
    If MySQL.OpenTable(ADOConn, rsload, , "select *, AES_DECRYPT(acci_services.Password,'" & odb.colSalts.ReturnSalt("md5Password") & "') as DecPassword from acci_services where RecID = " & RecID) = True Then
        If rsload.RecordCount > 0 Then
            txtxField(0) = IIf(IsNull(rsload!ContactName), "::", rsload!ContactName)
            txtxField(1) = IIf(IsNull(rsload!Username), "::", rsload!Username)
            txtxField(2) = IIf(IsNull(rsload!DecPassword), "::", rsload!DecPassword)
            txtxField(3) = IIf(IsNull(rsload!DateCreated), Format(Now), Format(rsload!DateCreated))
            txtxField(5) = IIf(IsNull(rsload!MBQuota), "20", rsload!MBQuota)
            optActive(IIf(IsNull(rsload!Checked), 0, IIf(Val(rsload!Checked) = -1, 0, 1))).Value = True
            
            If MySQL.OpenTable(ADOConn, rsCosts, , "select plantemplates.PeriodFee as tmCycleCost, plantemplates.FeePerBlock as tmMBFee, plantemplates.ExtraPerHour as tmPerHour, plantypes.Description as ptDesc, servicetypes.Description as svDesc " + _
                                                "from plantemplates, servicetypes, plantypes, sysops " + _
                                                "WHERE plantemplates.RecID = plantypes.TemplateID and servicetypes.RecID = plantypes.ServiceID and plantypes.RecID = " & rsload!ptRecID) = True Then
                If rsCosts.RecordCount > 0 Then
                    txtFee(1).Tag = rsCosts!tmCycleCost
                    txtFee(2).Tag = rsCosts!tmMBFee
                    txtFee(3).Tag = rsCosts!tmPerHour
                
                    Call MySQL.OpenTable(ADOConn, rsSysops, , "select sysops.Username from sysops where RecID = " & rsload!SysopID)
                
                    lblPLan(0).Caption = rsCosts!ptDesc
                    lblPLan(1).Caption = rsCosts!svDesc & " Created By " & rsSysops!Username
                            
                End If
            End If
            
            txtFee(1) = IIf(IsNull(rsload!PeriodFee) = True, 0, rsload!PeriodFee)
            txtFee(2) = IIf(IsNull(rsload!PerMB) = True, 0, rsload!PerMB)
            txtFee(3) = IIf(IsNull(rsload!PerHour) = True, 0, rsload!PerHour)
            mvCycle(0).Value = IIf(IsNull(rsload!NextCycle), sysNOW, rsload!NextCycle)
            dtpCycle(0).Value = IIf(IsNull(rsload!NextCycle), sysNOW, rsload!NextCycle)
            mvCycle(1).Value = IIf(IsNull(rsload!PreviousCycle), sysNOW, rsload!PreviousCycle)
            dtpCycle(1).Value = IIf(IsNull(rsload!PreviousCycle), sysNOW, rsload!PreviousCycle)
        
            If MySQL.OpenTable(ADOConn, rsPlans, , "select acci_services.ServiceID, acci_services.RecID, acci_services.Checked, acci_services.ContactName, acci_services.Username, AES_DECRYPT(acci_services.Password,'" & odb.colSalts.ReturnSalt("md5Password") & "') as Password, acci_services.BaseURL, acci_services.DynamicField1, acci_services.DynamicField2, acci_services.DynamicField3, acci_services.DynamicField4, acci_services.DynamicField5 , plantypes.Description from acci_services, plantypes Where plantypes.RecID = acci_services.ptRecID AND acci_services.RecID <> " & RecID & " AND SubRecID = " & rsload!SubRecID) = True Then
                If rsPlans.RecordCount > 0 Then
                    rsPlans.MoveFirst
                    While Not rsPlans.EOF And Err.Number = 0
                        Set itmX = lvPlans.ListItems.Add(, "r" & rsPlans!RecID, IIf(IsNull(rsPlans!ContactName), "", rsPlans!ContactName))
                        itmX.Tag = rsPlans!ServiceID
                        itmX.SubItems(1) = IIf(IsNull(rsPlans!Description), "", rsPlans!Description)
                        itmX.SubItems(2) = IIf(IsNull(rsPlans!Username), "", rsPlans!Username)
                        itmX.SubItems(3) = IIf(IsNull(rsPlans!Password), "", rsPlans!Password)
                        itmX.SubItems(4) = IIf(IsNull(rsPlans!BaseURL), "", rsPlans!BaseURL)
                        itmX.SubItems(5) = IIf(IsNull(rsPlans!DynamicField1), "", rsPlans!DynamicField1)
                        itmX.SubItems(6) = IIf(IsNull(rsPlans!DynamicField2), "", rsPlans!DynamicField2)
                        itmX.SubItems(7) = IIf(IsNull(rsPlans!DynamicField3), "", rsPlans!DynamicField3)
                        itmX.SubItems(8) = IIf(IsNull(rsPlans!DynamicField4), "", rsPlans!DynamicField4)
                        itmX.SubItems(9) = IIf(IsNull(rsPlans!DynamicField5), "", rsPlans!DynamicField5)
                        itmX.Checked = rsPlans!Checked
                        rsPlans.MoveNext
                    Wend
                End If
            End If
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

Private Sub txtFee_Change(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtFee_Change"
    Const ContainerName = "frmPlanProperties"
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


    If Index > 3 Then Exit Sub
    txtFee(Index + 3).Text = Format(Val(txtFee(Index).Text) - Val(txtFee(Index).Tag), "Currency")
    
    txtFee(Index).ToolTipText = "GST Inclusive Price " & Format(Val(txtFee(Index)) * oTax(Login.TaxCode, Login.TaxCountry) + Val(txtFee(Index)), "Currency")
    
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

Private Sub txtFee_DblClick(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtFee_DblClick"
    Const ContainerName = "frmPlanProperties"
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


    frmGSTCalc.cAmount = Val(txtFee(Index)) * oTax(Login.TaxCode, Login.TaxCountry) + txtFee(Index)
    frmGSTCalc.Show 1
    txtFee(Index) = "" & frmGSTCalc.cAmount
    
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

Private Sub txtFee_KeyPress(Index As Integer, KeyAscii As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtFee_KeyPress"
    Const ContainerName = "frmPlanProperties"
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
    Case 13
        KeyAscii = 0
        SendKeys "{TAB}"
    Case 48 To 57
    Case Asc(".")
        If InStr(txtFee(Index).Text, ".") > 0 Then KeyAscii = 0
    Case 8
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
