VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmSupItems 
   Caption         =   "Supplier Items"
   ClientHeight    =   9585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15000
   Icon            =   "frmSupItems.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9585
   ScaleWidth      =   15000
   Begin VB.Frame frameItems 
      Caption         =   "Items"
      Height          =   9315
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin MSComctlLib.ListView lvItems 
         Height          =   8865
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   270
         Width           =   7545
         _ExtentX        =   13309
         _ExtentY        =   15637
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         PictureAlignment=   5
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
            Text            =   "RRP"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cost"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Category"
            Object.Width           =   2540
         EndProperty
         Picture         =   "frmSupItems.frx":0442
      End
   End
   Begin VB.PictureBox picts 
      BorderStyle     =   0  'None
      Height          =   8955
      Index           =   1
      Left            =   8040
      ScaleHeight     =   8955
      ScaleWidth      =   6795
      TabIndex        =   24
      Top             =   450
      Visible         =   0   'False
      Width           =   6795
      Begin VB.Frame Frame1 
         Caption         =   "Measured Units"
         Height          =   2295
         Left            =   120
         TabIndex        =   39
         Top             =   4770
         Width           =   2715
         Begin VB.Frame frameUnits 
            Enabled         =   0   'False
            Height          =   1575
            Left            =   120
            TabIndex        =   41
            Top             =   600
            Width           =   2475
            Begin VB.OptionButton optUnits 
               Caption         =   "Measured in Set Units"
               Height          =   285
               Index           =   3
               Left            =   120
               TabIndex        =   45
               Top             =   1140
               Value           =   -1  'True
               Width           =   2205
            End
            Begin VB.OptionButton optUnits 
               Caption         =   "Measured in Days"
               Height          =   285
               Index           =   2
               Left            =   120
               TabIndex        =   44
               Top             =   840
               Width           =   2205
            End
            Begin VB.OptionButton optUnits 
               Caption         =   "Measured in Minutes"
               Height          =   285
               Index           =   1
               Left            =   120
               TabIndex        =   43
               Top             =   540
               Width           =   2205
            End
            Begin VB.OptionButton optUnits 
               Caption         =   "Measured in Hours"
               Height          =   285
               Index           =   0
               Left            =   120
               TabIndex        =   42
               Top             =   240
               Width           =   2205
            End
         End
         Begin VB.CheckBox chkUnits 
            Caption         =   "Item can have multiple Units"
            Height          =   255
            Left            =   180
            TabIndex        =   40
            Top             =   300
            Width           =   2415
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frame6"
         Height          =   4155
         Left            =   3300
         TabIndex        =   31
         Top             =   4770
         Visible         =   0   'False
         Width           =   3075
         Begin VB.CheckBox Check4 
            Height          =   285
            Index           =   2
            Left            =   150
            TabIndex        =   34
            Top             =   1020
            Width           =   2805
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Return to Base Warranty"
            Height          =   285
            Index           =   1
            Left            =   150
            TabIndex        =   33
            Top             =   600
            Width           =   2805
         End
         Begin VB.CheckBox Check4 
            Caption         =   "On Site Warranty is included"
            Height          =   285
            Index           =   0
            Left            =   150
            TabIndex        =   32
            Top             =   300
            Width           =   2805
         End
         Begin VB.Line Line1 
            X1              =   60
            X2              =   3000
            Y1              =   930
            Y2              =   930
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Components"
         Height          =   3375
         Left            =   120
         TabIndex        =   27
         ToolTipText     =   "This is where you set up product linking so that a product is included with extra and options"
         Top             =   1320
         Width           =   6555
         Begin VB.CheckBox Check3 
            Caption         =   "Hide sub items from browser"
            Height          =   315
            Left            =   4050
            TabIndex        =   30
            Top             =   270
            Width           =   2385
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Item is a master node of multiple components"
            Height          =   375
            Left            =   150
            TabIndex        =   29
            Top             =   240
            Width           =   4155
         End
         Begin MSComctlLib.ListView lvItems 
            Height          =   2595
            Index           =   1
            Left            =   120
            TabIndex        =   28
            Top             =   630
            Width           =   6315
            _ExtentX        =   11139
            _ExtentY        =   4577
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            GridLines       =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            Enabled         =   0   'False
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Item Description"
               Object.Width           =   10936
            EndProperty
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Viewability"
         Height          =   1215
         Left            =   120
         TabIndex        =   25
         Top             =   90
         Width           =   2745
         Begin VB.CheckBox Check1 
            Caption         =   "Hidden from VISP"
            Height          =   255
            Index           =   0
            Left            =   300
            TabIndex        =   26
            Top             =   330
            Width           =   2295
         End
      End
   End
   Begin VB.PictureBox picts 
      BorderStyle     =   0  'None
      Height          =   8925
      Index           =   0
      Left            =   8070
      ScaleHeight     =   8925
      ScaleWidth      =   6765
      TabIndex        =   3
      Top             =   510
      Width           =   6765
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   345
         Left            =   5070
         TabIndex        =   38
         Top             =   8400
         Width           =   1605
      End
      Begin VB.CommandButton cmdClone 
         Caption         =   "&Clone Entry"
         Enabled         =   0   'False
         Height          =   345
         Left            =   3420
         TabIndex        =   37
         Top             =   8400
         Width           =   1605
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update Database"
         Enabled         =   0   'False
         Height          =   345
         Left            =   5070
         TabIndex        =   36
         Top             =   8010
         Width           =   1605
      End
      Begin VB.CommandButton cmdCreate 
         Caption         =   "&Create New"
         Height          =   345
         Left            =   3420
         TabIndex        =   35
         Top             =   8010
         Width           =   1605
      End
      Begin VB.Frame Frame2 
         Caption         =   "Warranty"
         Height          =   765
         Index           =   8
         Left            =   60
         TabIndex        =   22
         Top             =   7950
         Width           =   3285
         Begin VB.TextBox txtField 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   7
            Left            =   150
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   23
            Top             =   270
            Width           =   2955
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Manufacture"
         Height          =   765
         Index           =   7
         Left            =   3420
         TabIndex        =   20
         Top             =   7140
         Width           =   3255
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   150
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   270
            Width           =   2985
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Model"
         Height          =   765
         Index           =   6
         Left            =   60
         TabIndex        =   18
         Top             =   7140
         Width           =   3285
         Begin VB.TextBox txtField 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   5
            Left            =   150
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   19
            Top             =   270
            Width           =   2955
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Image"
         Height          =   2355
         Left            =   2370
         TabIndex        =   16
         Top             =   3870
         Width           =   4305
         Begin VB.CommandButton cmdImage 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   255
            Left            =   3690
            TabIndex        =   17
            Top             =   1980
            Width           =   465
         End
         Begin VB.Image Image1 
            Height          =   1965
            Left            =   120
            Stretch         =   -1  'True
            Top             =   270
            Width           =   4035
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Category"
         Height          =   765
         Index           =   5
         Left            =   60
         TabIndex        =   14
         Top             =   6330
         Width           =   6615
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   150
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   270
            Width           =   6345
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "RRP (Ex GST)"
         Height          =   765
         Index           =   4
         Left            =   60
         TabIndex        =   12
         Top             =   5460
         Width           =   2235
         Begin VB.TextBox txtField 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   4
            Left            =   150
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   13
            Top             =   270
            Width           =   1905
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Cost (Ex GST)"
         Height          =   765
         Index           =   3
         Left            =   60
         TabIndex        =   10
         Top             =   4650
         Width           =   2235
         Begin VB.TextBox txtField 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   3
            Left            =   150
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   11
            Top             =   270
            Width           =   1905
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Supplier Code"
         Height          =   765
         Index           =   2
         Left            =   60
         TabIndex        =   8
         Top             =   3870
         Width           =   2235
         Begin VB.TextBox txtField 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   2
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   9
            Top             =   270
            Width           =   1965
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Item Feature Text"
         Height          =   2955
         Index           =   1
         Left            =   60
         TabIndex        =   6
         Top             =   870
         Width           =   6645
         Begin VB.TextBox txtField 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2520
            Index           =   1
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   65535
            MultiLine       =   -1  'True
            TabIndex        =   7
            Top             =   270
            Width           =   6375
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Item Description"
         Height          =   765
         Index           =   0
         Left            =   60
         TabIndex        =   4
         Top             =   60
         Width           =   6645
         Begin VB.TextBox txtField 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   5
            Top             =   270
            Width           =   6375
         End
      End
   End
   Begin MSComctlLib.TabStrip ts 
      Height          =   9375
      Left            =   8010
      TabIndex        =   2
      Top             =   120
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   16536
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Item"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Constraints"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSupItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Resize"
    Const ContainerName = "frmSupItems"
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


    If Me.ScaleWidth > 7500 Then
        frameItems.Move frameItems.Left, frameItems.Top, Me.ScaleWidth - ts.Width - frameItems.Left * 3, Me.ScaleHeight - frameItems.Top * 2
        lvItems(0).Move lvItems(0).Left, lvItems(0).Top, frameItems.Width - lvItems(0).Left * 2, frameItems.Height - lvItems(0).Top - lvItems(0).Left
        ts.Move frameItems.Width + frameItems.Left * 2, ts.Top, ts.Width, Me.ScaleHeight - ts.Top * 2
        Call ts_Click
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

Private Sub ts_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "ts_Click"
    Const ContainerName = "frmSupItems"
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


    Dim ix As Byte
    
    For ix = picts.LBound To picts.UBound
        If ix <> ts.SelectedItem.Index - 1 Then
            picts(ix).Visible = False
        Else
            picts(ix).Move ts.ClientLeft, ts.ClientTop, ts.ClientWidth, ts.ClientHeight
            picts(ix).ZOrder 0
            picts(ix).Visible = True
        End If
    Next
    
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
