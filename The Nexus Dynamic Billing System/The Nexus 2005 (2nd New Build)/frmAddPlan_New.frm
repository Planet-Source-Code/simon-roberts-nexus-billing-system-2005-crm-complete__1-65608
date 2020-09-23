VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAddPlan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Plan"
   ClientHeight    =   9285
   ClientLeft      =   3750
   ClientTop       =   1620
   ClientWidth     =   10740
   Icon            =   "frmAddPlan_New.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAddPlan_New.frx":030A
   ScaleHeight     =   619
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   716
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   120
      Left            =   4770
      Top             =   4410
   End
   Begin VB.CommandButton cmbCreate 
      Caption         =   "&Create Services"
      Enabled         =   0   'False
      Height          =   345
      Left            =   120
      TabIndex        =   32
      Top             =   8850
      Width           =   1575
   End
   Begin VB.PictureBox picts 
      BackColor       =   &H0092BA5A&
      BorderStyle     =   0  'None
      Height          =   7065
      Index           =   0
      Left            =   210
      ScaleHeight     =   7065
      ScaleWidth      =   10275
      TabIndex        =   1
      Top             =   1590
      Width           =   10275
      Begin VB.Frame frameExtras 
         BackColor       =   &H0092BA5A&
         Caption         =   "Extra Services in Subscription Package "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         Left            =   90
         TabIndex        =   2
         Top             =   4650
         Width           =   10095
         Begin MSComctlLib.ListView lvPlans 
            Height          =   1605
            Left            =   150
            TabIndex        =   3
            Top             =   330
            Width           =   9825
            _ExtentX        =   17330
            _ExtentY        =   2831
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   11138478
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Plan Type"
               Object.Width           =   8899
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Number Of"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H0092BA5A&
         Caption         =   "Service and Plans"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4545
         Index           =   0
         Left            =   30
         TabIndex        =   4
         Top             =   0
         Width           =   10215
         Begin MSComctlLib.TreeView tvServiceTypes 
            Height          =   4155
            Left            =   120
            TabIndex        =   5
            Top             =   270
            Width           =   3225
            _ExtentX        =   5689
            _ExtentY        =   7329
            _Version        =   393217
            Indentation     =   441
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            HotTracking     =   -1  'True
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
         End
         Begin MSComctlLib.ListView lvAccounts 
            Height          =   4185
            Left            =   3420
            TabIndex        =   6
            Top             =   240
            Width           =   6645
            _ExtentX        =   11721
            _ExtentY        =   7382
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   11138478
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
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Description"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Monthly Fee"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Montly Data"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Monthly Time"
               Object.Width           =   2540
            EndProperty
         End
      End
   End
   Begin MSComctlLib.TabStrip ts 
      Height          =   7485
      Left            =   150
      TabIndex        =   0
      Top             =   1260
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   13203
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Service or Plan to add"
            Object.ToolTipText     =   "Select the service and plan you want to added to the customer account information."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Constraints"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Item Text"
            Object.ToolTipText     =   "This is where the text description and further information about the service is displayed."
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picts 
      BorderStyle     =   0  'None
      Height          =   7125
      Index           =   1
      Left            =   150
      ScaleHeight     =   7125
      ScaleWidth      =   10365
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   10365
      Begin VB.Frame Frame7 
         Caption         =   "Setup Fees and Setup Contracts to choose from"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2505
         Left            =   90
         TabIndex        =   34
         Top             =   4560
         Width           =   10185
         Begin MSComctlLib.ListView lvContracts 
            Height          =   2085
            Left            =   150
            TabIndex        =   35
            Tag             =   "0"
            Top             =   330
            Width           =   9945
            _ExtentX        =   17542
            _ExtentY        =   3678
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   16777215
            BackColor       =   7697007
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   9
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "Description"
               Text            =   "Description"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "NoPeriods"
               Text            =   "Interval Length (ttl)"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "TypePeriods"
               Text            =   "Interval Code"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Tag             =   "Termination"
               Text            =   "Termination Fee"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Tag             =   "JoiningFee"
               Text            =   "Joining Fee"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Tag             =   "PeriodFee"
               Text            =   "Billing Cycle Fee"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Object.Tag             =   "FeePerBlock"
               Text            =   "Fee Per MB Block"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Object.Tag             =   "FeePerHour"
               Text            =   "Per Extra Hour"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Object.Tag             =   "RecID"
               Text            =   "Contract ID"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Adjusts Inital Package Costs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4455
         Left            =   4500
         TabIndex        =   11
         Top             =   60
         Width           =   5745
         Begin VB.Frame Frame5 
            Caption         =   "Total Package Cost"
            Height          =   705
            Index           =   1
            Left            =   150
            TabIndex        =   27
            Top             =   3600
            Width           =   5445
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Total:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   3
               Left            =   180
               TabIndex        =   31
               Top             =   300
               Width           =   600
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Margin:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   2
               Left            =   2760
               TabIndex        =   30
               Top             =   300
               Width           =   795
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "$ 0.00"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   225
               Index           =   3
               Left            =   720
               TabIndex        =   29
               Top             =   300
               Width           =   1815
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "$ 0.00"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000080FF&
               Height          =   225
               Index           =   2
               Left            =   3360
               TabIndex        =   28
               Top             =   300
               Width           =   1905
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Total Package Value"
            Height          =   705
            Index           =   0
            Left            =   150
            TabIndex        =   18
            Top             =   2850
            Width           =   5445
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "$ 0.00"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   1
               Left            =   3570
               TabIndex        =   26
               Top             =   270
               Width           =   1695
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "$ 0.00"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   1050
               TabIndex        =   25
               Top             =   270
               Width           =   1485
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Inc. Tax:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   1
               Left            =   2640
               TabIndex        =   20
               Top             =   270
               Width           =   1035
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Ex Tax:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   0
               Left            =   180
               TabIndex        =   19
               Top             =   270
               Width           =   780
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Cost for Plan"
            Height          =   2055
            Left            =   150
            TabIndex        =   13
            Top             =   780
            Width           =   5445
            Begin VB.TextBox txtFee 
               Alignment       =   2  'Center
               Height          =   360
               Index           =   3
               Left            =   2760
               TabIndex        =   24
               Top             =   1560
               Width           =   2535
            End
            Begin VB.TextBox txtFee 
               Alignment       =   2  'Center
               Height          =   360
               Index           =   2
               Left            =   2760
               TabIndex        =   23
               Top             =   1140
               Width           =   2535
            End
            Begin VB.TextBox txtFee 
               Alignment       =   2  'Center
               Height          =   360
               Index           =   1
               Left            =   2760
               TabIndex        =   22
               Top             =   720
               Width           =   2535
            End
            Begin VB.TextBox txtFee 
               Alignment       =   2  'Center
               Height          =   360
               Index           =   0
               Left            =   2760
               TabIndex        =   21
               Top             =   300
               Width           =   2535
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Per Extra Hour (Excluding Tax):"
               Height          =   240
               Index           =   3
               Left            =   180
               TabIndex        =   17
               Top             =   1620
               Width           =   2370
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Per Mb Block (Excluding Tax):"
               Height          =   240
               Index           =   2
               Left            =   180
               TabIndex        =   16
               Top             =   1200
               Width           =   2190
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Cost/Cycle Price (Excluding Tax):"
               Height          =   240
               Index           =   1
               Left            =   180
               TabIndex        =   15
               Top             =   780
               Width           =   2520
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Joining Fee (Excluding Tax):"
               Height          =   240
               Index           =   0
               Left            =   180
               TabIndex        =   14
               Top             =   360
               Width           =   2085
            End
         End
         Begin VB.ComboBox cmbPlans 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   150
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   330
            Width           =   5475
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Activation Date Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4425
         Left            =   120
         TabIndex        =   8
         Top             =   90
         Width           =   4305
         Begin MSComCtl2.DTPicker dtpActivation 
            Height          =   525
            Left            =   270
            TabIndex        =   10
            Top             =   3660
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   926
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   53149698
            UpDown          =   -1  'True
            CurrentDate     =   37738
         End
         Begin MSComCtl2.MonthView mvActivation 
            Height          =   3270
            Left            =   270
            TabIndex        =   9
            Top             =   360
            Width           =   3750
            _ExtentX        =   6615
            _ExtentY        =   5768
            _Version        =   393216
            ForeColor       =   -2147483630
            BackColor       =   -2147483633
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            StartOfWeek     =   53149697
            CurrentDate     =   37738
         End
      End
   End
   Begin VB.PictureBox picts 
      BorderStyle     =   0  'None
      Height          =   7125
      Index           =   2
      Left            =   180
      ScaleHeight     =   7125
      ScaleWidth      =   10365
      TabIndex        =   33
      Top             =   1560
      Visible         =   0   'False
      Width           =   10365
      Begin RichTextLib.RichTextBox txtProductText 
         Height          =   6855
         Left            =   150
         TabIndex        =   36
         Top             =   150
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   12091
         _Version        =   393217
         BackColor       =   14718607
         Enabled         =   -1  'True
         TextRTF         =   $"frmAddPlan_New.frx":1444B4
      End
   End
End
Attribute VB_Name = "frmAddPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsPlanType As adodb.Recordset

Public Activation As Date

Public sDesc As String

Public lServiceID As Double

Public ptRecID As Double


Private Sub cmbCreate_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmbCreate_Click"
    Const ContainerName = "frmAddPlan"
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


    If lvContracts.ListItems.Count > 0 Then
    
    
        If lvAccounts.Tag <> "" And Not lvContracts.SelectedItem Is Nothing Then
            ptRecID = lvAccounts.Tag
            sDesc = lvAccounts.SelectedItem.Text
            
            SetConstraints
            
            
        Else
             If lvAccounts.Tag <> "" Then
                ptRecID = lvAccounts.Tag
                sDesc = lvAccounts.SelectedItem.Text
                oService.ContractID = 0
                oService.IntervalLength = 5
                oService.IntervalType = "s"
    
                
                SetConstraints
                
                
            End If
        End If
    Else
    
        If lvAccounts.Tag <> "" Then
            ptRecID = lvAccounts.Tag
            sDesc = lvAccounts.SelectedItem.Text
            oService.ContractID = 0
            oService.IntervalLength = 5
            oService.IntervalType = "s"

            
            SetConstraints
            
            
        End If
        
    End If
    
    oService.Activated = CDate(Format(mvActivation.Value, "dd mmm yyyy") & " " & Format(dtpActivation.Value, "ttttt"))
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

Private Sub cmbPlans_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmbPlans_Click"
    Const ContainerName = "frmAddPlan"
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


    Dim itmX As New clsPackageNode
    
    Set itmX = oService(cmbPlans.ListIndex + 1)
    txtFee(0).Text = itmX.JoiningFee
    txtFee(1).Text = itmX.PeriodFee
    txtFee(2).Text = itmX.PerMBBlock
    txtFee(3).Text = itmX.PerHour
    
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
    Const ContainerName = "frmAddPlan"
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


    Set lvPlans.SmallIcons = fIcon.il16x16
            Set lvPlans.Icons = fIcon.il32x32
            Set lvAccounts.SmallIcons = fIcon.il16x16
            Set lvAccounts.Icons = fIcon.il32x32
            
    If Login.lLevel > 85 Then
        txtFee(0).Locked = False
    Else
        txtFee(0).Locked = True
    End If
        
    Set oService = New colPackages
    
    
    mvActivation.Value = sysnow
    dtpActivation.Value = sysnow
    
        Call GUI.LoadColWidths(lvPlans, Me)
        Call GUI.LoadColWidths(lvContracts, Me)
    
    If bBigFont = True Then

        lvAccounts.Font.Size = 16
        lvPlans.Font.Size = 16
        lvContracts.Font.Size = 16
        ts.Font.Size = 14
        tvServiceTypes.Font.Size = 16
        cmbPlans.Font.Size = 15
        
    End If
    
    PopulateList
    
        
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
    Const ContainerName = "frmAddPlan"
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


    Call GUI.SaveColWidths(lvPlans, Me)
    Call GUI.SaveColWidths(lvContracts, Me)
    
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

Private Sub lvAccounts_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvAccounts_ColumnClick"
    Const ContainerName = "frmAddPlan"
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


    Call GUI.ColumnSort(ColumnHeader, lvAccounts)
    
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

Private Sub lvAccounts_DblClick()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvAccounts_DblClick"
    Const ContainerName = "frmAddPlan"
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


    If lvAccounts.Tag <> "" Then
        ptRecID = lvAccounts.Tag
        sDesc = lvAccounts.SelectedItem.Text
        
        SetConstraints
        
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

Private Sub lvAccounts_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvAccounts_ItemClick"
    Const ContainerName = "frmAddPlan"
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


    BuildServicesNode Val(Mid(Item.Key, 2))
    
    Dim rsload As adodb.Recordset
    
    Call MySQL.OpenTable(directConn, rsload, , "select plantemplates.ProductText as ProductText from plantemplates, plantypes where plantypes.TemplateID = plantemplates.RecID and plantypes.RecID = " & Mid(Item.Key, 2))
    If rsload.RecordCount > 0 Then
    
        txtProductText.TextRTF = IIf(IsNull(rsload!ProductText), "", rsload!ProductText)
    
    End If
    LoadContracts Mid(Item.Key, 2)
    
    cmbCreate.Enabled = True
    
    'LoadFlags Val(Mid(Item.Key, 2))
    lvAccounts.Tag = Mid(Item.Key, 2)
    
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


Public Function LoadFlags(lRecID As Variant)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "LoadFlags"
    Const ContainerName = "frmAddPlan"
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


    Exit Function
    
    Dim bResult As Boolean
    Dim rsload As adodb.Recordset
    Dim itmX As ListItem
    
    bResult = MySQL.OpenTable(directConn, rsload, , "select plantypes.Description, flags_plantype.* from plantypes, flags_plantype Where flags_plantype.PlanType = plantypes.RecID AND ptRecID = " & lRecID)
    
    lvPlans.ListItems.Clear
    
    If rsload.RecordCount > 0 Then
        rsload.MoveFirst
        While Not rsload.EOF And Err.Number = 0
            Set itmX = lvPlans.ListItems.Add(, "r" & rsload!RecID, rsload!Description)
            itmX.SubItems(1) = rsload!NumberOf
            itmX.Checked = rsload!Checked
            rsload.MoveNext
        Wend
    End If


Exit Function



ErrorOccur:
Select Case oErr.chkError(directConn, Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function


Private Sub TabStrip1_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "TabStrip1_Click"
    Const ContainerName = "frmAddPlan"
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

Private Sub lvContracts_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvContracts_ItemClick"
    Const ContainerName = "frmAddPlan"
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




    oService.ClearContract
    
    Dim ix As Integer
    
    For ix = cmbPlans.ListCount To oService.Count + 1 Step -1
        cmbPlans.RemoveItem ix - 1
    Next
   
    Dim br As Boolean
    Dim rsPt As adodb.Recordset
    Dim rsSv As adodb.Recordset
    Dim rsEx As adodb.Recordset
    
    br = MySQL.OpenTable(directConn, rsSv, , "select servicetypes.RecID, servicetypes.ServiceKey, servicetypes.Description, servicetypes.SubofRecID, servicetypes.ListOnRadius, servicetypes.HasUID, servicetypes.HasSysUID, servicetypes.BillImmediately, servicetypes_matrix.SubServiceID  from servicetypes inner join servicetypes_matrix on servicetypes.RecID = servicetypes_matrix.ServiceID and servicetypes_matrix.VirtualID = '" & Login.lVirtualID & "'")
    br = MySQL.OpenTable(directConn, rsPt, , "select * from plantypes where RecID = '" & Login.MISCFee & "' Limit 1")
    
    br = MySQL.OpenTable(directConn, rsEx, , "select flags_tempextras.NumberOf, flags_tempextras.Checked, flags_tempextras.PlanType, plantypes.SessionsAllowed, plantypes.SessionTimeout, plantypes.IdleTimeout, plantypes.TemplateID, plantypes.ServiceID, plantypes.MBQuota, plantypes.CostPrice, plantypes.JoiningFee, plantypes.BillImmediately, plantypes.chgInterval, plantypes.chgIntervalType, plantypes.ExtraPerHour, plantypes.FeePerBlock, plantypes.PeriodFee, plantypes.Description, plantypes.RecID as ExtraRecID, plantypes.PeriodFee, plantypes.MBPerPeriod  from flags_tempextras, plantypes, contractsruntime " + _
                                            " Where plantypes.RecID = flags_tempextras.PlanType And flags_tempextras.ContractID = contractsruntime.ContractID " + _
                                            " AND contractsruntime.RecID = '" & Mid(Item.Key, 2) & "'")
      
        rsSv.Filter = "RecID = " & rsPt!ServiceID
    Randomize Now
    Set oItem = oService.Add(rsPt!Description, 1, Val(Item.SubItems(5)), Val(Item.SubItems(4)), Val(Item.SubItems(6)), Val(Item.SubItems(7)), rsPt!chgIntervalType, rsPt!chgInterval, rsSv!ServiceKey, rsSv!BillImmediately, rsPt!MBQuota, 0, 0, 0, 0, rsPt!SessionTimeout, rsPt!IdleTimeout, rsPt!SessionsAllowed, rsPt!ServiceID, Val(rsSv!ListOnRadius), 0, "r" & oService.Count + 1)
    
    cmbPlans.AddItem rsPt!Description
    
    oItem.cJoiningFee = Val(Item.SubItems(4))
    oItem.cPerHour = 0
    oItem.cPerMBBlock = 0
    oItem.cPeriodFee = 0
    oItem.bContract = True
    
    
    Dim itmX  As ListItem
    
    While Not rsEx.EOF And Err.Number = 0
        rsSv.Filter = "RecID = " & rsEx!ServiceID
        Set oItem = oService.Add(rsEx!Description, rsEx!NumberOf, rsEx!PeriodFee, rsEx!JoiningFee, rsEx!FeePerBlock, rsEx!ExtraPerHour, rsEx!chgIntervalType, rsEx!chgInterval, rsSv!ServiceKey, rsSv!BillImmediately, rsEx!MBQuota, 0, 0, 0, 0, rsEx!SessionTimeout, rsEx!IdleTimeout, rsEx!SessionsAllowed, rsEx!ServiceID, Val(rsSv!ListOnRadius), rsEx!ExtraRecID, "r" & oService.Count + 1)
        
        br = MySQL.OpenTable(directConn, rsPt, , "select PeriodFee, FeePerBlock, ExtraPerHour from plantemplates where RecID = " & rsEx!TemplateID & " Limit 1")
        oItem.cJoiningFee = 0
        oItem.cPerHour = 0
        oItem.cPerMBBlock = 0
        oItem.cPeriodFee = 0
        oItem.bContract = True
        cmbPlans.AddItem rsEx!Description
        
        rsEx.MoveNext
    Wend
    
    oService.ContractID = Mid(Item.Key, 2)
    oService.IntervalLength = Item.SubItems(1)
    oService.IntervalType = Item.SubItems(2)
    
    oService.calTotals oTax(Login.TaxCode, Login.TaxCountry)
    
    Label3(0).Caption = Format(oService.ttlExTax, "Currency")
    Label3(1).Caption = Format(oService.ttlTax, "Currency")
    If oService.ttlMargin < 0 Then Label3(2).ForeColor = RGB(255, 0, 0) Else Label3(2).ForeColor = &H80FF&
    Label3(2).Caption = Format(oService.ttlMargin, "Currency")
    Label3(3).Caption = Format(oService.ttlCost, "Currency")

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

Private Sub Timer1_Timer()

    Static bSet As Boolean
    
    If Login.IconsSet = True Then
        If bSet = False Then
            
            
            Set tvServiceTypes.ImageList = fIcon.il16x16
            'Set tvCat.ImageList = fIcon.il16x16
            Set lvPlans.SmallIcons = fIcon.il16x16
            Set lvPlans.Icons = fIcon.il32x32
            Set lvAccounts.SmallIcons = fIcon.il16x16
            Set lvAccounts.Icons = fIcon.il32x32
            
            bSet = True
        End If
    End If
    
    gSleep
    
End Sub

Private Sub ts_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "ts_Click"
    Const ContainerName = "frmAddPlan"
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
    
    For bx = picTS.LBound To picTS.UBound
        If bx <> ts.SelectedItem.Index - 1 Then
            picTS(bx).Visible = False
        Else
            picTS(bx).Move ts.ClientLeft, ts.ClientTop, ts.ClientWidth, ts.ClientHeight
            picTS(bx).Visible = True
            picTS(bx).ZOrder 0
            
        End If
    Next
    
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

Private Sub tvservicetypes_NodeClick(ByVal Node As MSComctlLib.Node)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "tvservicetypes_NodeClick"
    Const ContainerName = "frmAddPlan"
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
    
    Dim bResult As Boolean
    Dim itmX As ListItem
    lvAccounts.ListItems.Clear
    
    bResult = MySQL.SetColumnHeaders("accounttype", lvAccounts, Node.Tag, directConn)
    
    cmbAllPLans.Clear
    bResult = MySQL.OpenTable(directConn, rsPlanType, , "select * from plantypes")
    If rsPlanType.RecordCount > 0 Then
        While Not rsPlanType.EOF And Err.Number = 0
            cmbAllPLans.AddItem rsPlanType!Description
            cmbAllPLans.ItemData(cmbAllPLans.ListCount - 1) = rsPlanType!RecID
            rsPlanType.MoveNext
        Wend
    End If
    
    lvPlans.ListItems.Clear
    
    If Login.lAgencyID <> 0 Then
        bResult = MySQL.OpenTable(directConn, rsPlanType, , "select plantypes.* from plantypes, agencyplans Where agencyplans.ptRecID = plantypes.RecID and agencyplans.IsAvailable = -1 and ServiceID = " & Mid(Node.Key, 2))
    Else
        bResult = MySQL.OpenTable(directConn, rsPlanType, , "select * from plantypes Where ServiceID = " & Mid(Node.Key, 2) & " and VirtualID = " & Login.lVirtualID)
    End If
    
    Call MySQL.fillLV(directConn, rsPlanType, lvAccounts, False)
    
    Exit Sub
    cmbRollover.Clear
    Dim rsload As adodb.Recordset
    Dim SQL As String
    Dim sResult As String
    Dim bNumeric As Boolean
    
        While Not rsPlanType.EOF And Err.Number = 0
            Set itmX = lvAccounts.ListItems.Add(, "r" & rsPlanType!RecID, "")
            For X = 1 To lvAccounts.ColumnHeaders.Count
                If lvAccounts.ColumnHeaders(X).Tag <> "" Then
                    Select Case Mid(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") + 1, Len(lvAccounts.ColumnHeaders(X).Tag) - InStr(lvAccounts.ColumnHeaders(X).Tag, "^"))
                    Case "No Format"
                        Select Case X
                        Case 1
                            itmX.Text = MySQL.OP(rsPlanType, Left(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") - 1))
                            If InStr(itmX.Text, "-1") > 0 Then itmX.SubItems(X - 1) = sSTR.ReplaceString(itmX.Text, "-1", "Unlimited")
                        Case Else
                            itmX.SubItems(X - 1) = MySQL.OP(rsPlanType, Left(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") - 1))
                            If InStr(itmX.SubItems(X - 1), "-1") > 0 Then itmX.SubItems(X - 1) = sSTR.ReplaceString(itmX.SubItems(X - 1), "-1", "Unlimited")
                        End Select
                    Case Else
                        If rsPlanType(lvAccounts.ColumnHeaders(X).Tag) <> -1 Then
                            If InStr(LCase(Mid(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") + 1)), "select") > 0 Then
                                sResult = MySQL.fldType(rsPlanType(Left(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") - 1)).Type, bNumeric)
                                Select Case bNumeric
                                Case True
                                    If Not IsNull(rsPlanType(Left(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") - 1))) Then
                                        SQL = LCase(Mid(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") + 1)) & "'" & Val(rsPlanType(Left(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") - 1))) & "'"
                                    Else
                                        SQL = LCase(Mid(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") + 1)) & "'0'"
                                    End If
                                Case False
                                    If Not IsNull(rsPlanType(Left(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") - 1))) Then
                                        SQL = LCase(Mid(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") + 1)) & "'" & MySQL.ESC(rsPlanType(Left(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") - 1))) & "'"
                                    Else
                                        SQL = LCase(Mid(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") + 1)) & "''"
                                    End If
                                End Select
                                
                                Call MySQL.OpenTable(directConn, rsload, , SQL)
                                If rsload.RecordCount > 0 Then
                                   sResult = MySQL.fldType(rsload("nResult").Type, bNumeric)
                                   Select Case bNumeric
                                   Case True
                                        
                                        If X = 1 Then
                                            itmX.Text = "" & Val(IIf(IsNull(rsload("nResult")), 0, rsload("nResult")))
                                        Else
                                            itmX.SubItems(X - 1) = "" & Val(IIf(IsNull(rsload("nResult")), 0, rsload("nResult")))
                                        End If
                                        
                                   Case False

                                        If X = 1 Then
                                            itmX.Text = "" & IIf(IsNull(rsload("nResult")), 0, rsload("nResult"))
                                            If Left(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") - 1) = "vendorid" Then itmX.Text = MySQL.CodeME(itmX.Text)
                                        Else
                                            itmX.SubItems(X - 1) = "" & IIf(IsNull(rsload("nResult")), "", rsload("nResult"))
                                            If Left(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") - 1) = "VendorID" Then itmX.SubItems(X - 1) = MySQL.CodeME(itmX.SubItems(X - 1))
                                        End If
                                                                           End Select
                                   'itmX.SubItems(X-1) = Format(MySQL.OP(rsPlanType, Left(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") - 1)), Mid(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") + 1, Len(lvAccounts.ColumnHeaders(X).Tag) - InStr(lvAccounts.ColumnHeaders(X).Tag, "^")))
                                End If
                            Else
                                
                                If X = 1 Then
                                    itmX.Text = "" & Format(MySQL.OP(rsPlanType, Left(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") - 1)), Mid(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") + 1, Len(lvAccounts.ColumnHeaders(X).Tag) - InStr(lvAccounts.ColumnHeaders(X).Tag, "^")))
                                    If InStr(itmX.Text, "-1") > 0 Then itmX.Text = sSTR.ReplaceString(itmX.Text, "-1", "Unlimited")
                                Else
                                    itmX.SubItems(X - 1) = "" & Format(MySQL.OP(rsPlanType, Left(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") - 1)), Mid(lvAccounts.ColumnHeaders(X).Tag, InStr(lvAccounts.ColumnHeaders(X).Tag, "^") + 1, Len(lvAccounts.ColumnHeaders(X).Tag) - InStr(lvAccounts.ColumnHeaders(X).Tag, "^")))
                                    If InStr(itmX.SubItems(X - 1), "-1") > 0 Then itmX.SubItems(X - 1) = sSTR.ReplaceString(itmX.SubItems(X - 1), "-1", "Unlimited")
                                End If
                                        
                            End If
                        End If
                    End Select
                End If
                gSleep
            Next X
        rsPlanType.MoveNext
    Wend
    
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

Public Function PopulateList()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "PopulateList"
    Const ContainerName = "frmAddPlan"
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
    Dim rsload As adodb.Recordset
    Dim NodX As Node
    Dim NodeX As Node
    
    bResult = MySQL.OpenTable(directConn, rsload, , "select servicetypes.RecID, servicetypes.ServiceKey, servicetypes.Description, servicetypes.SubofRecID, servicetypes.ListOnRadius, servicetypes.HasUID, servicetypes.HasSysUID, servicetypes.BillImmediately, servicetypes_matrix.SubServiceID  from servicetypes inner join servicetypes_matrix on servicetypes.RecID = servicetypes_matrix.ServiceID and servicetypes_matrix.VirtualID = '" & Login.lVirtualID & "'")
    
    If bResult = True Then

        
        tvServiceTypes.NodeS.Clear
        
        If bResult = False Then Exit Function
        
        If rsload.RecordCount > 0 Then
            While Not rsload.EOF And Err.Number = 0
                If Not IsNull(rsload!SubofRecID) Then
                    If tvServiceTypes.NodeS("k" & rsload!SubofRecID) Is Nothing Then
                        Set NodeX = tvServiceTypes.NodeS("k" & rsload!SubServiceID)
                        Set NodX = tvServiceTypes.NodeS.Add(NodeX.Key, tvwChild, "k" & rsload!RecID, IIf(IsNull(rsload!Description), "Description Not set", rsload!Description))
                        NodX.Tag = rsload!ServiceKey
                        NodX.Image = IIf(fIcon.il16x16.ListImages.Count = 1, 1, Round(Rnd * 110))
                        NodeX.Expanded = False
                    ElseIf rsload!SubofRecID <> 0 Then
                        Set NodeX = tvServiceTypes.NodeS("k" & rsload!SubofRecID)
                        Set NodX = tvServiceTypes.NodeS.Add(NodeX.Key, tvwChild, "k" & rsload!RecID, IIf(IsNull(rsload!Description), "Description Not set", rsload!Description))
                        NodX.Image = IIf(fIcon.il16x16.ListImages.Count = 1, 1, Round(Rnd * 110))
                        NodX.Tag = rsload!ServiceKey
                        NodeX.Expanded = False
                    Else
                        Set NodX = tvServiceTypes.NodeS.Add(, , "k" & rsload!RecID, IIf(IsNull(rsload!Description), "Description Not set", rsload!Description))
                        NodX.Image = IIf(fIcon.il16x16.ListImages.Count = 1, 1, Round(Rnd * 110))
                        NodX.Tag = rsload!ServiceKey
                    End If
                Else
                    Set NodX = tvServiceTypes.NodeS.Add(, , "k" & rsload!RecID, IIf(IsNull(rsload!Description), "Description Not set", rsload!Description))
                    NodX.Image = IIf(fIcon.il16x16.ListImages.Count = 1, 1, Round(Rnd * 110))
                    NodX.Tag = rsload!ServiceKey
                End If
                rsload.MoveNext
            Wend
    
        End If
    End If
    
Exit Function



ErrorOccur:
Select Case oErr.chkError(directConn, Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function



Public Sub SetConstraints()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "SetConstraints"
    Const ContainerName = "frmAddPlan"
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


    Me.Activation = CDate(Format(mvActivation.Value, "dd-mm-yyyy") & " " & Format(dtpActivation.Value, "Hh:Nn:Ss"))
    
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

Public Sub BuildServicesNode(ptRecID As Long)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "BuildServicesNode"
    Const ContainerName = "frmAddPlan"
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


    oService.Clear
    
    Dim br As Boolean
    Dim rsPt As adodb.Recordset
    Dim rsSv As adodb.Recordset
    Dim rsEx As adodb.Recordset
    
    br = MySQL.OpenTable(directConn, rsSv, , "select servicetypes.RecID, servicetypes.ServiceKey, servicetypes.Description, servicetypes.SubofRecID, servicetypes.ListOnRadius, servicetypes.HasUID, servicetypes.HasSysUID, servicetypes.BillImmediately, servicetypes_matrix.SubServiceID  from servicetypes inner join servicetypes_matrix on servicetypes.RecID = servicetypes_matrix.ServiceID and servicetypes_matrix.VirtualID = '" & Login.lVirtualID & "'")
    br = MySQL.OpenTable(directConn, rsPt, , "select * from plantypes where RecID = " & ptRecID & " Limit 1")
    br = MySQL.OpenTable(directConn, rsEx, , "select flags_plantype.*, plantypes.SessionsAllowed, plantypes.SessionTimeout, plantypes.IdleTimeout, plantypes.TemplateID, plantypes.ServiceID, plantypes.MBQuota, plantypes.CostPrice, plantypes.JoiningFee, plantypes.BillImmediately, plantypes.chgInterval, plantypes.chgIntervalType, plantypes.ExtraPerHour, plantypes.FeePerBlock, plantypes.PeriodFee, plantypes.Description, plantypes.RecID as ExtraRecID, plantypes.PeriodFee, plantypes.MBPerPeriod from flags_plantype, plantypes where PlanType = plantypes.RecID AND flags_plantype.Checked = -1 and ptRecID = " & ptRecID)

    Dim oItem As New clsPackageNode
    
    cmbPlans.Clear
    
    rsSv.Filter = "RecID = " & rsPt!ServiceID
    Randomize Now
    Set oItem = oService.Add(rsPt!Description, 1, rsPt!PeriodFee, rsPt!JoiningFee, rsPt!FeePerBlock, rsPt!ExtraPerHour, rsPt!chgIntervalType, rsPt!chgInterval, rsSv!ServiceKey, rsSv!BillImmediately, rsPt!MBQuota, 0, 0, 0, 0, rsPt!SessionTimeout, rsPt!IdleTimeout, rsPt!SessionsAllowed, rsPt!ServiceID, Val(rsSv!ListOnRadius), ptRecID, "r" & oService.Count + 1)
    
    cmbPlans.AddItem rsPt!Description
    
    br = MySQL.OpenTable(directConn, rsPt, , "select PeriodFee, FeePerBlock, ExtraPerHour from plantemplates where RecID = " & rsPt!TemplateID & " Limit 1")
    oItem.cJoiningFee = 0
    oItem.cPerHour = rsPt!ExtraPerHour
    oItem.cPerMBBlock = rsPt!FeePerBlock
    oItem.cPeriodFee = rsPt!PeriodFee
    

    
    
    lvPlans.ListItems.Clear
    
    Dim itmX  As ListItem
    
    While Not rsEx.EOF And Err.Number = 0
        On Error Resume Next
        rsSv.Filter = "RecID = " & rsEx!ServiceID
        Set oItem = oService.Add(rsEx!Description, rsEx!NumberOf, rsEx!PeriodFee, rsEx!JoiningFee, rsEx!FeePerBlock, rsEx!ExtraPerHour, rsEx!chgIntervalType, rsEx!chgInterval, rsSv!ServiceKey, rsSv!BillImmediately, rsEx!MBQuota, 0, 0, 0, 0, rsEx!SessionTimeout, rsEx!IdleTimeout, rsEx!SessionsAllowed, rsEx!ServiceID, Val(rsSv!ListOnRadius), rsEx!ExtraRecID, "r" & oService.Count + 1)
        
        br = MySQL.OpenTable(directConn, rsPt, , "select PeriodFee, FeePerBlock, ExtraPerHour from plantemplates where RecID = " & rsEx!TemplateID & " Limit 1")
        oItem.cJoiningFee = 0
        oItem.cPerHour = rsPt!ExtraPerHour
        oItem.cPerMBBlock = rsPt!FeePerBlock
        oItem.cPeriodFee = rsPt!PeriodFee
        
        cmbPlans.AddItem rsEx!Description
        
        Set itmX = lvPlans.ListItems.Add(, "r" & rsEx!ExtraRecID, rsEx!Description)
        itmX.SubItems(1) = rsEx!NumberOf
        itmX.Checked = -1
        
        rsEx.MoveNext
    Wend
    
    cmbPlans.ListIndex = 0
    
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

Private Sub txtFee_Change(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtFee_Change"
    Const ContainerName = "frmAddPlan"
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


    Select Case Index
    Case 0
        oService(cmbPlans.ListIndex + 1).JoiningFee = Val(txtFee(Index).Text)
    Case 1
        oService(cmbPlans.ListIndex + 1).PeriodFee = Val(txtFee(Index))
    Case 2
        oService(cmbPlans.ListIndex + 1).PerMBBlock = Val(txtFee(Index))
    Case 3
        oService(cmbPlans.ListIndex + 1).PerHour = Val(txtFee(Index))
    End Select
    
    oService.calTotals oTax(Login.TaxCode, Login.TaxCountry)
    
    Label3(0).Caption = Format(oService.ttlExTax, "Currency")
    Label3(1).Caption = Format(oService.ttlTax, "Currency")
    If oService.ttlMargin < 0 Then Label3(2).ForeColor = RGB(255, 0, 0) Else Label3(2).ForeColor = &H80FF&
    Label3(2).Caption = Format(oService.ttlMargin, "Currency")
    Label3(3).Caption = Format(oService.ttlCost, "Currency")
    
    txtFee(Index).ToolTipText = "GST Inclusive Price " & Format(Val(txtFee(Index)) * oTax(Login.TaxCode, Login.TaxCountry) + Val(txtFee(Index)), "Currency")
    
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

Private Sub txtFee_DblClick(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtFee_DblClick"
    Const ContainerName = "frmAddPlan"
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


    frmGSTCalc.cAmount = Val(txtFee(Index)) * oTax(Login.TaxCode, Login.TaxCountry) + txtFee(Index)
    frmGSTCalc.Show 1
    txtFee(Index) = "" & frmGSTCalc.cAmount
    
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

Private Sub txtFee_GotFocus(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtFee_GotFocus"
    Const ContainerName = "frmAddPlan"
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


    txtFee(Index).SelStart = 0
    txtFee(Index).SelLength = Len(txtFee(Index).Text)
    
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

Private Sub txtFee_KeyPress(Index As Integer, KeyAscii As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtFee_KeyPress"
    Const ContainerName = "frmAddPlan"
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
    Case 48 To 57
    Case Asc(".")
        If InStr(txtFee(Index).Text, ".") > 0 Then KeyAscii = 0
    Case 8
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

Public Sub LoadContracts(ptRecID As Variant)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "LoadContracts"
    Const ContainerName = "frmAddPlan"
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


    Dim rsload As adodb.Recordset


        Call MySQL.OpenTable(directConn, rsload, , "select contracttemplates.Description, contracttemplates.NoPeriods, contracttemplates.TypePeriods, contractsruntime.* from contracttemplates, contractsruntime where contractsruntime.ContractID = contracttemplates.RecID and contractsruntime.ptRecID = " & ptRecID & " and contracttemplates.bDeleted = 0")
        Dim itmX As ListItem
        Dim bx As Byte
        lvContracts.ListItems.Clear
        If rsload.State = adStateOpen Then
            If rsload.RecordCount > 0 Then
                While Not rsload.EOF And Err.Number = 0
                    Set itmX = lvContracts.ListItems.Add(, "c" & rsload!RecID, IIf(IsNull(rsload!Description), "(null)", rsload!Description))
                    For bx = 2 To lvContracts.ColumnHeaders.Count

                        itmX.SubItems(bx - 1) = IIf(IsNull(rsload(lvContracts.ColumnHeaders(bx).Tag)), "0", rsload(lvContracts.ColumnHeaders(bx).Tag))

                    Next
                    itmX.Checked = True
                    rsload.MoveNext
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
