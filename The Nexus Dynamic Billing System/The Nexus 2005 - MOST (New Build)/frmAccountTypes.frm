VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAccountTypes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Service & Plan Types"
   ClientHeight    =   11820
   ClientLeft      =   3390
   ClientTop       =   1710
   ClientWidth     =   10245
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAccountTypes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11820
   ScaleWidth      =   10245
   Begin VB.CommandButton cmdSetup 
      Caption         =   "Setup Contracts and Setup Fees"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4470
      TabIndex        =   74
      Top             =   5310
      Width           =   2595
   End
   Begin VB.CommandButton cmdSpiff 
      Caption         =   "Bonuses and Spiff"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7110
      TabIndex        =   60
      Top             =   5310
      Width           =   1935
   End
   Begin VB.Frame Frame11 
      Caption         =   "Billing Period"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   1
      Left            =   150
      TabIndex        =   42
      Top             =   9720
      Width           =   10005
      Begin VB.CheckBox chkBillOnce 
         Caption         =   "Bill Only Once"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1680
         TabIndex        =   59
         Top             =   330
         Width           =   1485
      End
      Begin VB.OptionButton optBillingCycle 
         Caption         =   "Days:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   5
         Left            =   3240
         TabIndex        =   48
         Tag             =   "d"
         Top             =   330
         Width           =   855
      End
      Begin VB.TextBox txtBillingCycle 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   5
         Left            =   4170
         MaxLength       =   5
         TabIndex        =   47
         Text            =   "14"
         Top             =   330
         Width           =   675
      End
      Begin VB.OptionButton optBillingCycle 
         Caption         =   "Months:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   4950
         TabIndex        =   46
         Tag             =   "m"
         Top             =   330
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.TextBox txtBillingCycle 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   4
         Left            =   6030
         MaxLength       =   5
         TabIndex        =   45
         Text            =   "1"
         Top             =   330
         Width           =   675
      End
      Begin VB.OptionButton optBillingCycle 
         Caption         =   "Hours:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   3
         Left            =   6870
         TabIndex        =   44
         Tag             =   "h"
         Top             =   330
         Width           =   975
      End
      Begin VB.TextBox txtBillingCycle 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   3
         Left            =   7890
         MaxLength       =   5
         TabIndex        =   43
         Text            =   "96"
         Top             =   330
         Width           =   675
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Rollover Period"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Index           =   0
      Left            =   150
      TabIndex        =   6
      Top             =   10560
      Width           =   10005
      Begin VB.TextBox txtBillingCycle 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   7890
         MaxLength       =   5
         TabIndex        =   14
         Text            =   "96"
         Top             =   330
         Width           =   675
      End
      Begin VB.OptionButton optBillingCycle 
         Caption         =   "Hours:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   6870
         TabIndex        =   13
         Tag             =   "h"
         Top             =   330
         Width           =   975
      End
      Begin VB.TextBox txtBillingCycle 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   6030
         MaxLength       =   5
         TabIndex        =   12
         Text            =   "1"
         Top             =   330
         Width           =   675
      End
      Begin VB.OptionButton optBillingCycle 
         Caption         =   "Months:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   4950
         TabIndex        =   11
         Tag             =   "m"
         Top             =   330
         Width           =   1035
      End
      Begin VB.TextBox txtBillingCycle 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   4170
         MaxLength       =   5
         TabIndex        =   10
         Text            =   "14"
         Top             =   330
         Width           =   675
      End
      Begin VB.OptionButton optBillingCycle 
         Caption         =   "Days:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3240
         TabIndex        =   9
         Tag             =   "d"
         Top             =   330
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.ComboBox cmbRollover 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3300
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   690
         Width           =   5265
      End
      Begin VB.CheckBox chkLimit 
         Caption         =   "Roll Over"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   300
         TabIndex        =   7
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "To Plan Type:"
         Enabled         =   0   'False
         Height          =   240
         Left            =   2190
         TabIndex        =   15
         Top             =   750
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9060
      TabIndex        =   4
      Top             =   5310
      Width           =   1035
   End
   Begin VB.CommandButton cmdAddType 
      Caption         =   "Add Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   150
      TabIndex        =   3
      Top             =   5310
      Width           =   2175
   End
   Begin VB.Frame frameServices 
      Caption         =   "Services and Plans"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4545
      Left            =   150
      TabIndex        =   1
      Top             =   690
      Width           =   10005
      Begin MSComctlLib.TreeView tvServiceTypes 
         Height          =   4155
         Left            =   120
         TabIndex        =   5
         Top             =   270
         Width           =   3105
         _ExtentX        =   5477
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ListView lvAccounts 
         Height          =   4215
         Left            =   3270
         TabIndex        =   2
         Top             =   240
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   7435
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
   Begin VB.Frame frameExtras 
      Caption         =   "Extra's"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   2580
      TabIndex        =   51
      Top             =   2760
      Width           =   6105
      Begin MSComctlLib.ListView lvPlans 
         Height          =   2145
         Left            =   150
         TabIndex        =   52
         Top             =   270
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   3784
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
   Begin VB.Frame frameConfig 
      BackColor       =   &H0047ADE4&
      BorderStyle     =   0  'None
      Height          =   1725
      Index           =   1
      Left            =   7410
      TabIndex        =   37
      Top             =   8010
      Width           =   2805
      Begin VB.OptionButton Option1 
         BackColor       =   &H001394F2&
         Caption         =   "Backorder"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   765
         Index           =   7
         Left            =   8580
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   2700
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H001394F2&
         Caption         =   "No more stock not reorderable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   765
         Index           =   6
         Left            =   8580
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   1860
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H001394F2&
         Caption         =   "Not available anymore"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   765
         Index           =   5
         Left            =   8580
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   1050
         Width           =   1185
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H001394F2&
         Caption         =   "Item is Orderable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   765
         Index           =   4
         Left            =   8580
         MaskColor       =   &H0086D28D&
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   240
         Width           =   1155
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H0047ADE4&
         Caption         =   "Product Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   2100
         TabIndex        =   64
         Top             =   990
         Width           =   3735
         Begin VB.TextBox txtCat 
            Appearance      =   0  'Flat
            BackColor       =   &H0047ADE4&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   1
            Left            =   180
            MaxLength       =   30
            TabIndex        =   65
            Text            =   "INT0000"
            Top             =   240
            Width           =   3375
         End
      End
      Begin VB.TextBox txtFee 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   5
         Left            =   6900
         TabIndex        =   54
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtFee 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   2100
         TabIndex        =   39
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox txtDescription 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2100
         Locked          =   -1  'True
         MaxLength       =   128
         TabIndex        =   38
         Top             =   270
         Width           =   6375
      End
      Begin VB.Label lblMBQuota 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Quota: 10 MB's"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   6690
         TabIndex        =   58
         Top             =   960
         Width           =   1785
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H001394F2&
         BackStyle       =   0  'Transparent
         Caption         =   "Joining Fee:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   5940
         TabIndex        =   55
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H001394F2&
         BackStyle       =   0  'Transparent
         Caption         =   "Fee:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   1515
         TabIndex        =   41
         Top             =   600
         Width           =   315
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H001394F2&
         BackStyle       =   0  'Transparent
         Caption         =   "Plan Description:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   615
         TabIndex        =   40
         Top             =   270
         Width           =   1200
      End
   End
   Begin VB.Frame frameConfig 
      BackColor       =   &H0086D28D&
      BorderStyle     =   0  'None
      Height          =   1755
      Index           =   2
      Left            =   6510
      TabIndex        =   61
      Top             =   7860
      Width           =   3585
      Begin MSComctlLib.ListView lvContracts 
         Height          =   3315
         Left            =   90
         TabIndex        =   75
         Tag             =   "0"
         Top             =   90
         Width           =   9675
         _ExtentX        =   17066
         _ExtentY        =   5847
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
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
   Begin VB.Frame frameConfig 
      BackColor       =   &H00EC7A71&
      BorderStyle     =   0  'None
      Height          =   3525
      Index           =   0
      Left            =   240
      TabIndex        =   17
      Top             =   6150
      Width           =   9885
      Begin VB.OptionButton Option1 
         BackColor       =   &H001394F2&
         Caption         =   "Backorder"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   765
         Index           =   3
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   2670
         Width           =   1185
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H001394F2&
         Caption         =   "No more stock not reorderable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   765
         Index           =   2
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   1830
         Width           =   1185
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H001394F2&
         Caption         =   "Not available anymore"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   765
         Index           =   1
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   1020
         Width           =   1185
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H001394F2&
         Caption         =   "Item is orderable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   765
         Index           =   0
         Left            =   8640
         MaskColor       =   &H0086D28D&
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   210
         Width           =   1185
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00EC7A71&
         Caption         =   "Product Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   62
         Top             =   150
         Width           =   2535
         Begin VB.TextBox txtCat 
            Appearance      =   0  'Flat
            BackColor       =   &H00EC7A71&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   405
            Index           =   0
            Left            =   210
            MaxLength       =   30
            TabIndex        =   63
            Text            =   "INT0000"
            Top             =   240
            Width           =   2115
         End
      End
      Begin VB.TextBox txtFee 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   4
         Left            =   6870
         TabIndex        =   53
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox cmdRadius 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5130
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   2910
         Width           =   3405
      End
      Begin VB.TextBox txtDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4200
         Locked          =   -1  'True
         MaxLength       =   128
         TabIndex        =   33
         Top             =   270
         Width           =   4245
      End
      Begin VB.TextBox txtFee 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   4200
         TabIndex        =   32
         Top             =   600
         Width           =   1515
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EC7A71&
         Height          =   945
         Index           =   0
         Left            =   1230
         TabIndex        =   24
         Top             =   870
         Width           =   7335
         Begin VB.CheckBox chkLimit 
            Caption         =   "Data Limits"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   1
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   210
            Width           =   1335
         End
         Begin VB.TextBox txtMB 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   3060
            TabIndex        =   27
            Top             =   240
            Width           =   1425
         End
         Begin VB.TextBox txtMB 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   3060
            TabIndex        =   26
            Top             =   570
            Width           =   1425
         End
         Begin VB.TextBox txtFee 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   5790
            TabIndex        =   25
            Top             =   210
            Width           =   1425
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MB's Per Month:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   2
            Left            =   1815
            TabIndex        =   31
            Top             =   240
            Width           =   1170
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MB's Per Block:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   3
            Left            =   1860
            TabIndex        =   30
            Top             =   600
            Width           =   1125
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fee Per Block:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   4
            Left            =   4620
            TabIndex        =   29
            Top             =   270
            Width           =   1050
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EC7A71&
         Height          =   945
         Index           =   1
         Left            =   1230
         TabIndex        =   18
         Top             =   1830
         Width           =   7335
         Begin VB.TextBox txtFee 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   2
            Left            =   5790
            TabIndex        =   21
            Top             =   570
            Width           =   1425
         End
         Begin VB.TextBox txtHours 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5790
            TabIndex        =   20
            Top             =   240
            Width           =   1425
         End
         Begin VB.CheckBox chkLimit 
            Caption         =   "Time Limits"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   210
            Width           =   1335
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fee Per Extra Hour:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   5
            Left            =   4305
            TabIndex        =   23
            Top             =   570
            Width           =   1395
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hours Per Cycle:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   6
            Left            =   4515
            TabIndex        =   22
            Top             =   240
            Width           =   1185
         End
      End
      Begin VB.Label lblMBQuota 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Quota: 10 MB's"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   1290
         TabIndex        =   57
         Top             =   2970
         Width           =   1485
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Joining Fee:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   5910
         TabIndex        =   56
         Top             =   630
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Radius Pool:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   4170
         TabIndex        =   50
         Top             =   2970
         Width           =   915
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plan Description:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   2850
         TabIndex        =   35
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fee Per Bill Cycle:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   2835
         TabIndex        =   34
         Top             =   630
         Width           =   1275
      End
   End
   Begin MSComctlLib.TabStrip tsPlan 
      Height          =   3915
      Left            =   150
      TabIndex        =   16
      Top             =   5760
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   6906
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Plan Configuration"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Extra's"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Contracts"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   2865
      Left            =   240
      TabIndex        =   36
      Top             =   6180
      Width           =   8355
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Types"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   960
      TabIndex        =   0
      Top             =   30
      Width           =   2520
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   -30
      Picture         =   "frmAccountTypes.frx":0442
      Stretch         =   -1  'True
      Top             =   -60
      Width           =   930
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   10230
      Y1              =   420
      Y2              =   420
   End
End
Attribute VB_Name = "frmAccountTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VendorID As Long
Dim RecID As Long

Private Sub Check2_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Check2_Click"
    Const ContainerName = "frmAccountTypes"
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

Private Sub chkLimit_Click(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "chkLimit_Click"
    Const ContainerName = "frmAccountTypes"
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
    Case 0 ' Time Limits
        
        txtHours.Enabled = IIf(chkLimit(Index).Value = 1, True, False)
        txtFee(2).Enabled = IIf(chkLimit(Index).Value = 1, True, False)
        
    Case 1 ' Data Limits
    
        txtMB(0).Enabled = IIf(chkLimit(Index).Value = 1, True, False)
        txtMB(1).Enabled = IIf(chkLimit(Index).Value = 1, True, False)
        txtFee(1).Enabled = IIf(chkLimit(Index).Value = 1, True, False)
        
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



Private Sub cmdAddType_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdAddType_Click"
    Const ContainerName = "frmAccountTypes"
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

    
    If Not tvServiceTypes.SelectedItem Is Nothing Then
        
        Dim ffrmTemp As New frmTemplates
        ffrmTemp.ServiceID = Mid(tvServiceTypes.SelectedItem.Key, 2)
        ffrmTemp.Show 1
        
            
        If ffrmTemp.vRecID <> 0 Then
        
            cmdSave.Enabled = True
            
            Dim rsload As ADODB.Recordset
            Dim bResult As Boolean
            
            bResult = MySQL.OpenTable(ADOConn, rsload, , "select * from plantemplates Where RecID = " & ffrmTemp.vRecID)
            
            txtFee(4).Tag = Login.lVirtualID
            
            VendorID = rsload!VendorID
            txtDesc = rsload!Description
            txtFee(0).Tag = ffrmTemp.vRecID
            txtDesc.Locked = False
            txtDesc.Tag = 0
            txtDescription.Tag = 0
            txtDescription = rsload!Description
            txtDescription.Locked = False
            txtFee(3).Text = rsload!PeriodFee
            txtFee(0).Text = rsload!PeriodFee
            txtFee(3).Tag = rsload!PeriodFee
            
            If rsload!MBQuota = 0 Then
                lblMBQuota(0).Visible = False
                lblMBQuota(1).Visible = False
            Else
                lblMBQuota(0).Visible = True
                lblMBQuota(1).Visible = True
            End If
            
            lblMBQuota(0).Caption = "Data Quota: " & rsload!MBQuota & " MB's"
            lblMBQuota(0).Tag = rsload!MBQuota
            lblMBQuota(1).Caption = "Data Quota: " & rsload!MBQuota & " MB's"
            
            If rsload!MBPerPeriod <> -1 Then
                chkLimit(1).Value = 1
                Call chkLimit_Click(1)
                txtMB(1) = rsload!MBPerPeriod
                txtMB(0) = rsload!MBBlockSize
                txtFee(1) = rsload!FeePerBlock
            Else
                chkLimit(1).Value = 0
                Call chkLimit_Click(0)
                txtMB(0) = "0"
                txtMB(1) = "0"
                txtFee(1) = "0.00"
            End If
            
            
            If rsload!HoursPerPeriod <> -1 Then
                chkLimit(0).Value = 1
                Call chkLimit_Click(0)
                txtHours = rsload!HoursPerPeriod
                txtFee(2) = rsload!ExtraPerHour
            Else
                chkLimit(0).Value = 0
                Call chkLimit_Click(0)
                txtHours = "0"
                txtFee(2) = "0.00"
            End If
                        
            If Not IsNull(rsload!RadiusID) Then
                Dim ix As Integer
                For ix = 0 To cmdRadius.ListCount - 1
                    If cmdRadius.ItemData(ix) = rsload!RadiusID Then
                       cmdRadius.ListIndex = ix
                       Exit For
                    End If
                Next
            Else
                cmdRadius.Enabled = True
            End If
                        
            Dim itmX As ListItem
            
            bResult = MySQL.OpenTable(ADOConn, rsload, , "select plantypes.Description, flags_tempextras.* from plantypes, flags_tempextras Where flags_tempextras.PlanType = plantypes.RecID AND flags_tempextras.ptRecID = " & rsload!RecID)
            
            lvPlans.ListItems.Clear
            
            If rsload.RecordCount > 0 Then
                rsload.MoveFirst
                While Not rsload.EOF And Err.Number = 0
                    Set itmX = lvPlans.ListItems.Add(, "r" & rsload!RecID, rsload!Description)
                    itmX.Tag = rsload!PlanType
                    itmX.SubItems(1) = rsload!NumberOf
                    itmX.Checked = rsload!Checked
                    rsload.MoveNext
                Wend
            End If
            
            
            cmdSpiff.Enabled = False
            cmdSetup.Enabled = False
            
            'txtDesc.SetFocus
            cmbRollover.ListIndex = -1
            
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

Private Sub cmdSave_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdSave_Click"
    Const ContainerName = "frmAccountTypes"
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
    Dim rsServices As ADODB.Recordset
    Dim rsSave As ADODB.Recordset

    On Error Resume Next
    Dim tmpRecID As Double
    Dim tmpRecID3 As Double
    Dim tmpRecID2 As Double
    Dim RecID As Double
    Dim bNEW As Boolean
    
    Dim bx As Byte
    Dim rsPSType As ADODB.Recordset
    
    If Val(txtDescription.Tag) <> 0 Then RecID = Val(txtDescription.Tag)
    If Val(txtDesc.Tag) <> 0 Then RecID = Val(txtDesc.Tag)
    
    bResult = MySQL.OpenTable(ADOConn, rsPSType, , "select * from plantypes Where RecID = " & Val(RecID))
    bResult = MySQL.OpenTable(ADOConn, rsServices, , "select BillImmediately from servicetypes where RecID = " & Mid(tvServiceTypes.SelectedItem.Key, 2))
    
    Select Case tvServiceTypes.SelectedItem.Tag
    
    Case "FTP", "POP3", "WWW", "DESIGN", "DOMAIN", "GATEWAY", "TRAINING", "SALES", "CONSULT", "COLO", "HOSTING", "ALIAS", "HOST"
        Select Case Val(RecID)
        
        Case 0
            
            bNEW = True
            
            Do
                On Error Resume Next
                Err.Clear
                RecID = MySQL.GetTMPRecID("plantypes", ADOConn)
                Call MySQL.Execute(ADOConn, "Insert into plantypes (RecID, Description, PeriodFee, JoiningFee, Rollover, ServiceID, VirtualID, BillImmediately, TemplateID, MbQuota, VendorID ) " + "VALUES('" & RecID & "','" & MySQL.ESC(IIf(txtDesc <> "", txtDesc, txtDescription)) & "','" & Val(txtFee(3)) & "','" & Val(txtFee(5)) & "','" & chkLimit(2).Value & "','" & Val(Mid(tvServiceTypes.SelectedItem.Key, 2)) & "','" & txtFee(4).Tag & "','" & rsServices!BillImmediately & "','" & txtFee(0).Tag & "','" & lblMBQuota(0).Tag & "','" & VendorID & "')")
            Loop Until Err.Number = 0
            
            For bx = 0 To 2
                If optBillingCycle(bx).Value = True Then
                    Call MySQL.Execute(ADOConn, "update plantypes set roIntervalType = '" & optBillingCycle(bx).Tag & "' where RecID = " & RecID)
                    Call MySQL.Execute(ADOConn, "update plantypes set roInterval = '" & Val(IIf(txtBillingCycle(bx).Text = "", "1", txtBillingCycle(bx).Text)) & "' where RecID = " & RecID)
                    If cmbRollover.ListIndex <> -1 Then
                        Call MySQL.Execute(ADOConn, "update plantypes set roRecID = '" & cmbRollover.ItemData(cmbRollover.ListIndex) & "' where RecID = " & RecID)
                    Else
                        Call MySQL.Execute(ADOConn, "update plantypes set roRecID = '" & 0 & "' where RecID = " & RecID)
                    End If
                    Exit For
                End If
            Next
            
            For bx = 3 To 5
                If optBillingCycle(bx).Value = True Then
                    Call MySQL.Execute(ADOConn, "update plantypes set chgIntervalType = '" & optBillingCycle(bx).Tag & "' where RecID = " & RecID)
                    Call MySQL.Execute(ADOConn, "update plantypes set chgInterval = '" & txtBillingCycle(bx).Text & "' where RecID = " & RecID)
                    Exit For
                End If
            Next

            Call MySQL.Execute(ADOConn, "update plantypes set MBPerPeriod = -1 where RecID = " & RecID)
            Call MySQL.Execute(ADOConn, "update plantypes set MBBlockSize = 0 where RecID = " & RecID)
            Call MySQL.Execute(ADOConn, "update plantypes set FeePerBlock = 0 where RecID = " & RecID)
            Call MySQL.Execute(ADOConn, "update plantypes set HoursPerPeriod = -1 where RecID = " & RecID)
            Call MySQL.Execute(ADOConn, "update plantypes set ExtraPerHour = 0 where RecID = " & RecID)
            Call MySQL.Execute(ADOConn, "update plantypes set BillOnce = '" & chkBillOnce.Value & "' where RecID = " & RecID)
            Call MySQL.Execute(ADOConn, "update plantypes set CatNo = '" & txtCat(0) & "' where RecID = " & RecID)
            Call MySQL.Execute(ADOConn, "update plantypes set CostPrice = '" & txtFee(3).Tag & "' where RecID = " & RecID)
            
        Case Else
            
                        
            If MySQL.OpenTable(ADOConn, rsSave, , "select * from plantypes where RecID = " & RecID) = True Then
            
            
                If rsSave.RecordCount > 0 Then
                
                    If rsSave!Description <> txtDescription Then
                    
                        MySQL.Execute ADOConn, "Update accountviewer Set Description='View " & MySQL.ESC(txtDescription) & " Users' Where selectStatement Like '%" & rsSave!RecID & "%' and Description <> 'Active' and Description <> 'Cancelled'"
                        
                    End If
                    
                    Call MySQL.Execute(ADOConn, "update plantypes set Description ='" & MySQL.ESC(IIf(txtDesc <> "", txtDesc, txtDescription)) & "' where RecID = " & RecID)
                    Call MySQL.Execute(ADOConn, "update plantypes set PeriodFee = '" & Val(txtFee(3)) & "' where RecID = " & RecID)
                    Call MySQL.Execute(ADOConn, "update plantypes set JoiningFee = '" & Val(txtFee(5)) & "' where RecID = " & RecID)
                    Call MySQL.Execute(ADOConn, "update plantypes set Rollover = '" & chkLimit(2).Value & "' where RecID = " & RecID)
                    Call MySQL.Execute(ADOConn, "update plantypes set BillImmediately = '" & rsServices!BillImmediately & "' where RecID = " & RecID)
                    Call MySQL.Execute(ADOConn, "update plantypes set TemplateID = '" & txtFee(0).Tag & "' where RecID = " & RecID)
                    Call MySQL.Execute(ADOConn, "update plantypes set MbQuota = '" & lblMBQuota(0).Tag & "' where RecID = " & RecID)
                    Call MySQL.Execute(ADOConn, "update plantypes set TemplateID = '" & txtFee(0).Tag & "' where RecID = " & RecID)
                    Call MySQL.Execute(ADOConn, "update plantypes set VendorID = '" & VendorID & "' where RecID = " & RecID)
                    Call MySQL.Execute(ADOConn, "update plantypes set CatNo = '" & txtCat(0) & "' where RecID = " & RecID)
                    Call MySQL.Execute(ADOConn, "update plantypes set MbQuota = '" & lblMBQuota(0).Tag & "' where RecID = " & RecID)
                    Call MySQL.Execute(ADOConn, "update plantypes set TemplateID = '" & txtFee(0).Tag & "' where RecID = " & RecID)
                    Call MySQL.Execute(ADOConn, "update plantypes set CostPrice = '" & txtFee(3).Tag & "' where RecID = " & RecID)
                    
                    For bx = 0 To 2
                        If optBillingCycle(bx).Value = True Then
                            Call MySQL.Execute(ADOConn, "update plantypes set roIntervalType = '" & optBillingCycle(bx).Tag & "' where RecID = " & RecID)
                            Call MySQL.Execute(ADOConn, "update plantypes set roInterval = '" & Val(IIf(txtBillingCycle(bx).Text = "", "1", txtBillingCycle(bx).Text)) & "' where RecID = " & RecID)
                            If cmbRollover.ListIndex <> -1 Then
                                Call MySQL.Execute(ADOConn, "update plantypes set roRecID = '" & cmbRollover.ItemData(cmbRollover.ListIndex) & "' where RecID = " & RecID)
                            Else
                                Call MySQL.Execute(ADOConn, "update plantypes set roRecID = '" & 0 & "' where RecID = " & RecID)
                            End If
                            Exit For
                        End If
                    Next
                    
                    For bx = 3 To 5
                        If optBillingCycle(bx).Value = True Then
                            MySQL.Execute ADOConn, "Update plantypes Set chgIntervalType = '" & optBillingCycle(bx).Tag & "', chgInterval = " & txtBillingCycle(bx).Text & " where RecID = " & rsSave!RecID
                            Exit For
                        End If
                    Next
                    
                    Call MySQL.Execute(ADOConn, "update plantypes set MBPerPeriod = -1 where RecID = " & RecID)
                    Call MySQL.Execute(ADOConn, "update plantypes set MBBlockSize = 0 where RecID = " & RecID)
                    Call MySQL.Execute(ADOConn, "update plantypes set FeePerBlock = 0 where RecID = " & RecID)
                    Call MySQL.Execute(ADOConn, "update plantypes set HoursPerPeriod = -1 where RecID = " & RecID)
                    Call MySQL.Execute(ADOConn, "update plantypes set ExtraPerHour = 0 where RecID = " & RecID)
                    Call MySQL.Execute(ADOConn, "update plantypes set BillOnce = '" & chkBillOnce.Value & "' where RecID = " & RecID)
                    Call MySQL.Execute(ADOConn, "update plantypes set CatNo = '" & txtCat(0) & "' where RecID = " & RecID)
                
                End If
            End If
            'PopulateList
            Call tvservicetypes_NodeClick(tvServiceTypes.SelectedItem)
            
        End Select
    Case "DIALUP", "ADSL", "SHDSL"
            
        If cmdRadius.ListIndex = -1 Then
            MsgBox "You must set the radius pool that is being used!"
            Exit Sub
        End If

        Select Case RecID
        
        Case 0
            
            bNEW = True
            
            Do
                On Error Resume Next
                Err.Clear
                RecID = MySQL.GetTMPRecID("plantypes", ADOConn)

                Call MySQL.Execute(ADOConn, "Insert into plantypes (RecID, Description, PeriodFee, JoiningFee, Rollover, ServiceID, VirtualID, BillImmediately, TemplateID, MbQuota, VendorID ) " + "VALUES('" & RecID & "','" & MySQL.ESC(txtDesc) & "','" & Val(txtFee(0)) & "','" & Val(txtFee(4)) & "','" & chkLimit(2).Value & "','" & Val(Mid(tvServiceTypes.SelectedItem.Key, 2)) & "','" & txtFee(4).Tag & "','" & Val(rsServices!BillImmediately) & "','" & txtFee(0).Tag & "','" & lblMBQuota(0).Tag & "','" & VendorID & "')")
            Loop Until Err.Number = 0
            
            
            
            For bx = 0 To 2
                If optBillingCycle(bx).Value = True Then
                    Call MySQL.Execute(ADOConn, "update plantypes set roIntervalType = '" & optBillingCycle(bx).Tag & "' where RecID = " & RecID)
                    Call MySQL.Execute(ADOConn, "update plantypes set roInterval = '" & Val(IIf(txtBillingCycle(bx).Text = "", "1", txtBillingCycle(bx).Text)) & "' where RecID = " & RecID)
                    If cmbRollover.ListIndex <> -1 Then
                        Call MySQL.Execute(ADOConn, "update plantypes set roRecID = '" & cmbRollover.ItemData(cmbRollover.ListIndex) & "' where RecID = " & RecID)
                    Else
                        Call MySQL.Execute(ADOConn, "update plantypes set roRecID = '" & 0 & "' where RecID = " & RecID)
                    End If
                    Exit For
                End If
            Next
            
            For bx = 3 To 5
                If optBillingCycle(bx).Value = True Then
                    Call MySQL.Execute(ADOConn, "update plantypes set chgIntervalType = '" & optBillingCycle(bx).Tag & "' where RecID = " & RecID)
                    Call MySQL.Execute(ADOConn, "update plantypes set chgInterval = '" & txtBillingCycle(bx).Text & "' where RecID = " & RecID)
                    Exit For
                End If
            Next
                        
            If chkLimit(1).Value = 1 Then
                Call MySQL.Execute(ADOConn, "update plantypes set MBPerPeriod = '" & Val(txtMB(1)) & "' where RecID = " & RecID)
                Call MySQL.Execute(ADOConn, "update plantypes set MBBlockSize = '" & Val(txtMB(0)) & "' where RecID = " & RecID)
                Call MySQL.Execute(ADOConn, "update plantypes set FeePerBlock = '" & Val(txtFee(1)) & "' where RecID = " & RecID)
            Else
                Call MySQL.Execute(ADOConn, "update plantypes set MBPerPeriod = '" & 1 & "' where RecID = " & RecID)
                Call MySQL.Execute(ADOConn, "update plantypes set MBBlockSize = '" & 0 & "' where RecID = " & RecID)
                Call MySQL.Execute(ADOConn, "update plantypes set FeePerBlock = '" & 0 & "' where RecID = " & RecID)
            End If
            
            If chkLimit(0).Value = 1 Then
                Call MySQL.Execute(ADOConn, "update plantypes set HoursPerPeriod = '" & txtHours & "' where RecID = " & RecID)
                Call MySQL.Execute(ADOConn, "update plantypes set ExtraPerHour = '" & Val(txtFee(2)) & "' where RecID = " & RecID)
            Else
                Call MySQL.Execute(ADOConn, "update plantypes set HoursPerPeriod = '" & -1 & "' where RecID = " & RecID)
                Call MySQL.Execute(ADOConn, "update plantypes set ExtraPerHour = '" & 0 & "' where RecID = " & RecID)
            End If
                       
            
            Call MySQL.Execute(ADOConn, "update plantypes set BillOnce = '" & chkBillOnce.Value & "' where RecID = " & RecID)
            Call MySQL.Execute(ADOConn, "update plantypes set CostPrice = '" & txtFee(3).Tag & "' where RecID = " & RecID)
           
        Case Else
        
            If cmdRadius.ListIndex = -1 Then
                MsgBox "You must set the radius pool that is being used!"
                Exit Sub
            End If
            
                        
            
            If MySQL.OpenTable(ADOConn, rsSave, , "select * from plantypes where RecID = " & RecID) = True Then
                If rsSave.RecordCount > 0 Then
                    If rsSave!Description <> txtDescription Then
                        MySQL.Execute ADOConn, "Update accountviewer Set Description='View " & MySQL.ESC(txtDesc) & " Users' Where selectStatement Like '\'" & rsSave!RecID & "\'' and Description <> 'Active' and Description <> 'Cancelled'"
                        'rsSave!Description = txtDescription
                    End If
        
                    If rsSave.RecordCount > 0 Then
                        Call MySQL.Execute(ADOConn, "update plantypes set Description = '" & MySQL.ESC(txtDesc) & "' where RecID = " & RecID)
                        Call MySQL.Execute(ADOConn, "update plantypes set PeriodFee = '" & Val(txtFee(0)) & "' where RecID = " & RecID)
                        Call MySQL.Execute(ADOConn, "update plantypes set JoiningFee = '" & Val(txtFee(4)) & "' where RecID = " & RecID)
                        Call MySQL.Execute(ADOConn, "update plantypes set Rollover = '" & chkLimit(2).Value & "' where RecID = " & RecID)
                        Call MySQL.Execute(ADOConn, "update plantypes set BillImmediately = '" & rsServices!BillImmediately & "' where RecID = " & RecID)
                        Call MySQL.Execute(ADOConn, "update plantypes set RadiusID = '" & cmdRadius.ItemData(cmdRadius.ListIndex) & "' where RecID = " & RecID)
                        Call MySQL.Execute(ADOConn, "update plantypes set TemplateID = '" & txtFee(0).Tag & "' where RecID = " & RecID)
                        Call MySQL.Execute(ADOConn, "update plantypes set VirtualID = '" & txtFee(4).Tag & "' where RecID = " & RecID)
                        Call MySQL.Execute(ADOConn, "update plantypes set VendorID = '" & VendorID & "' where RecID = " & RecID)
                        Call MySQL.Execute(ADOConn, "update plantypes set CatNo = '" & txtCat(0) & "' where RecID = " & RecID)
                        Call MySQL.Execute(ADOConn, "update plantypes set CostPrice = '" & txtFee(3).Tag & "' where RecID = " & RecID)
                        
                        For bx = 0 To 2
                            If optBillingCycle(bx).Value = True Then
                                Call MySQL.Execute(ADOConn, "update plantypes set roIntervalType = '" & optBillingCycle(bx).Tag & "' where RecID = " & RecID)
                                Call MySQL.Execute(ADOConn, "update plantypes set roInterval = '" & Val(IIf(txtBillingCycle(bx).Text = "", "1", txtBillingCycle(bx).Text)) & "' where RecID = " & RecID)
                                If cmbRollover.ListIndex <> -1 Then
                                    Call MySQL.Execute(ADOConn, "update plantypes set roRecID = '" & cmbRollover.ItemData(cmbRollover.ListIndex) & "' where RecID = " & RecID)
                                Else
                                    Call MySQL.Execute(ADOConn, "update plantypes set roRecID = '" & 0 & "' where RecID = " & RecID)
                                End If
                                Exit For
                            End If
                        Next
                        
                        For bx = 3 To 5
                            If optBillingCycle(bx).Value = True Then
                                Call MySQL.Execute(ADOConn, "update plantypes set chgIntervalType = '" & optBillingCycle(bx).Tag & "' where RecID = " & RecID)
                                Call MySQL.Execute(ADOConn, "update plantypes set chgInterval = '" & txtBillingCycle(bx).Text & "' where RecID = " & RecID)
                                Exit For
                            End If
                        Next
                        
                        If chkLimit(1).Value = 1 Then
                            Call MySQL.Execute(ADOConn, "update plantypes set MBPerPeriod = '" & Val(txtMB(1)) & "' where RecID = " & RecID)
                            Call MySQL.Execute(ADOConn, "update plantypes set MBBlockSize = '" & Val(txtMB(0)) & "' where RecID = " & RecID)
                            Call MySQL.Execute(ADOConn, "update plantypes set FeePerBlock = '" & Val(txtFee(1)) & "' where RecID = " & RecID)
                        Else
                            Call MySQL.Execute(ADOConn, "update plantypes set MBPerPeriod = '" & 1 & "' where RecID = " & RecID)
                            Call MySQL.Execute(ADOConn, "update plantypes set MBBlockSize = '" & 0 & "' where RecID = " & RecID)
                            Call MySQL.Execute(ADOConn, "update plantypes set FeePerBlock = '" & 0 & "' where RecID = " & RecID)
                        End If
                        
                        If chkLimit(0).Value = 1 Then
                            Call MySQL.Execute(ADOConn, "update plantypes set HoursPerPeriod = '" & txtHours & "' where RecID = " & RecID)
                            Call MySQL.Execute(ADOConn, "update plantypes set ExtraPerHour = '" & Val(txtFee(2)) & "' where RecID = " & RecID)
                        Else
                            Call MySQL.Execute(ADOConn, "update plantypes set HoursPerPeriod = '" & -1 & "' where RecID = " & RecID)
                            Call MySQL.Execute(ADOConn, "update plantypes set ExtraPerHour = '" & 0 & "' where RecID = " & RecID)
                        End If
                    
                        Call MySQL.Execute(ADOConn, "update plantypes set BillOnce = '" & chkBillOnce.Value & "' where RecID = " & RecID)
                        
                    End If
               End If
            End If
        End Select
    End Select
    
    SaveFlags Val(RecID)
    
    Dim descrip As String
    
    If txtDesc <> "" Then descrip = txtDesc
    If txtDescription <> "" Then descrip = txtDescription
    
    MySQL.Execute ADOConn, "Update plantypes set Description = '" & MySQL.ESC(descrip) & "' where RecID = " & RecID
    MySQL.Execute ADOConn, "Update plantypes set CatNo = '" & MySQL.ESC(txtCat(0)) & "' where RecID = " & RecID
    
    Dim ox As Byte
    For ox = 0 To 3
        If Option1(ox).Value = 1 Then
            MySQL.Execute ADOConn, "Update plantypes set OptionalText = '" & MySQL.ESC(Option1(ox).Caption) & "' where RecID = " & RecID
        End If
    Next
    
    cmdSpiff.Enabled = True
    cmdSpiff.Tag = RecID
    
    txtDescription.Tag = RecID
    lvAccounts.Enabled = False
    tvServiceTypes.Enabled = False
    
    Call MySQL.OpenTable(ADOConn, rsSave, , "select * from plantypes where RecID = '" & RecID & "'")
    
    If bNEW = False Then
        Call MySQL.fillLV(ADOConn, rsSave, lvAccounts, True, lvAccounts.SelectedItem)
    Else
        Call MySQL.fillLV(ADOConn, rsSave, lvAccounts, True)
    End If
    
    lvAccounts.Enabled = True
    tvServiceTypes.Enabled = True
    
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

Private Sub cmdSetup_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdSetup_Click"
    Const ContainerName = "frmAccountTypes"
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


    Dim fSetup As New frmContracts1
    
    fSetup.ptRecID = Val(cmdSpiff.Tag)
    fSetup.PeriodFee = Val(IIf(Val(txtFee(0)) = 0, txtFee(3), txtFee(0)))
    fSetup.Show 1
    
    
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

Private Sub cmdSpiff_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdSpiff_Click"
    Const ContainerName = "frmAccountTypes"
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


    Dim fSpiff As New frmSpiff
    fSpiff.ptRecID = cmdSpiff.Tag
    
    fSpiff.Show 1
    
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
    Const ContainerName = "frmAccountTypes"
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


    frameConfig(0).Visible = True
    frameConfig(0).ZOrder 0
    frameConfig(0).Move tsPlan.ClientLeft, tsPlan.ClientTop, tsPlan.ClientWidth, tsPlan.ClientHeight
    
    If bBigFont = True Then
        Dim iCnt As Long
        For iCnt = frameConfig.LBound To frameConfig.UBound
            frameConfig(iCnt).Font.Size = 13
        Next
        lvAccounts.Font.Size = 14
        lvContracts.Font.Size = 16
        lvPlans.Font.Size = 16
        tvServiceTypes.Font.Size = 16
        tsPlan.Font.Size = 16
    End If
    
    PopulateList
    
    'Loads the first entry in
    tvServiceTypes.NodeS(1).Selected = True
    Call tvservicetypes_NodeClick(tvServiceTypes.SelectedItem)
    
    Call GUI.LoadColWidths(lvContracts, Me)
    
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

Public Function PopulateList()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "PopulateList"
    Const ContainerName = "frmAccountTypes"
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
    Dim NodX As Node
    Dim NodeX As Node
    
    bResult = MySQL.OpenTable(ADOConn, rsload, , "select * from servicetypes")
    
    tvServiceTypes.NodeS.Clear
    
    If rsload.RecordCount > 0 Then
        While Not rsload.EOF And Err.Number = 0
            If Not IsNull(rsload!SubofRecID) Then
                If rsload!SubofRecID <> 0 Then
                    Set NodeX = tvServiceTypes.NodeS("k" & rsload!SubofRecID)
                    Set NodX = tvServiceTypes.NodeS.Add(NodeX.Key, tvwChild, "k" & rsload!RecID, IIf(IsNull(rsload!Description), "Description Not set", rsload!Description))
                    NodX.Tag = rsload!ServiceKey
                    NodeX.Expanded = False
                Else
                    Set NodX = tvServiceTypes.NodeS.Add(, , "k" & rsload!RecID, IIf(IsNull(rsload!Description), "Description Not set", rsload!Description))
                    NodX.Tag = rsload!ServiceKey
                End If
            Else
                Set NodX = tvServiceTypes.NodeS.Add(, , "k" & rsload!RecID, IIf(IsNull(rsload!Description), "Description Not set", rsload!Description))
                NodX.Tag = rsload!ServiceKey
            End If
            rsload.MoveNext
        Wend

    End If
        
    bResult = MySQL.OpenTable(ADOConn, rsload, , "select * from radiuspools")
    If rsload.RecordCount > 0 Then
        While Not rsload.EOF And Err.Number = 0
            cmdRadius.AddItem rsload!Description
            cmdRadius.ItemData(cmdRadius.ListCount - 1) = rsload!RecID
            rsload.MoveNext
            gSleep
        Wend
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
    Const ContainerName = "frmAccountTypes"
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


    Call GUI.SaveColWidths(lvContracts, Me)
    
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

Private Sub lvAccounts_AfterLabelEdit(Cancel As Integer, NewString As String)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvAccounts_AfterLabelEdit"
    Const ContainerName = "frmAccountTypes"
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



    'MySQL.Execute ADOConn, "Update accountviewer Set Description='View " & MySQL.ESC(NewString) & " Users' Where selectStatement Like '\'" & txtDescription.Tag & "\'' and Description <> 'Active' and Description <> 'Cancelled'"
    MySQL.Execute ADOConn, "Update plantypes Set `" & Left(lvAccounts.ColumnHeaders(1).Tag, InStr(lvAccounts.ColumnHeaders(1).Tag, ">") - 1) & "`='" & MySQL.ESC(NewString) & "' Where RecID = " & txtDescription.Tag
    

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

Private Sub lvAccounts_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvAccounts_ColumnClick"
    Const ContainerName = "frmAccountTypes"
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


    Call GUI.ColumnSort(ColumnHeader, lvAccounts)
    
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

Private Sub lvAccounts_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvAccounts_ItemClick"
    Const ContainerName = "frmAccountTypes"
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


    Dim rsPSType As ADODB.Recordset
    bResult = MySQL.OpenTable(ADOConn, rsPSType, , "select * from plantypes Where RecID = " & Mid(Item.Key, 2))
    RecID = Val(Mid(Item.Key, 2))
    
    If rsPSType.State = adStateOpen Then
        If rsPSType.RecordCount > 0 Then
        
            txtFee(4).Tag = rsPSType!VirtualID
            txtFee(3).Tag = rsPSType!CostPrice
            
            Dim ox As Byte
            For ox = 0 To 3
                If Option1(ox).Value = 1 Then
                    MySQL.Execute ADOConn, "Update plantypes set OptionalText = '" & MySQL.ESC(Option1(ox).Caption) & "' where RecID = " & ltmpRecID
                End If
            Next
            
            If rsPSType!MBQuota = 0 Then
                lblMBQuota(0).Visible = False
                lblMBQuota(1).Visible = False
            Else
                lblMBQuota(0).Visible = True
                lblMBQuota(1).Visible = True
            End If
            
            lblMBQuota(0).Caption = "Data Quota: " & rsPSType!MBQuota & " MB's"
            lblMBQuota(0).Tag = rsPSType!MBQuota
            lblMBQuota(1).Caption = "Data Quota: " & rsPSType!MBQuota & " MB's"
            cmdSave.Enabled = True
            
            chkBillOnce.Value = IIf(IsNull(rsPSType!BillOnce), 0, rsPSType!BillOnce)
            
            txtCat(0) = IIf(IsNull(rsPSType!CatNo), "INT" & Round(Rnd * 65535), rsPSType!CatNo)
            
            Select Case tvServiceTypes.SelectedItem.Tag
            
            Case "FTP", "POP3", "WWW", "DESIGN", "DOMAIN", "GATEWAY", "TRAINING", "SALES", "CONSULT", "COLO", "HOST", "ALIAS"
                If rsPSType.RecordCount > 0 Then
                    txtDescription = IIf(IsNull(rsPSType!Description), "", rsPSType!Description)
                    txtDescription.Tag = rsPSType!RecID
                    txtFee(3) = rsPSType!PeriodFee
                    txtFee(5) = IIf(IsNull(rsPSType!JoiningFee), 0, rsPSType!JoiningFee)
                    txtFee(0).Tag = rsPSType!TemplateID
                End If
                
            Case "DIALUP", "ADSL", "SHDSL"
            
                If rsPSType.RecordCount > 0 Then
                
                    txtDesc = IIf(IsNull(rsPSType!Description), "", rsPSType!Description)
                    txtDesc.Tag = rsPSType!RecID
                    txtFee(0) = rsPSType!PeriodFee
                    txtFee(4) = IIf(IsNull(rsPSType!JoiningFee), 0, rsPSType!JoiningFee)
                    txtFee(0).Tag = rsPSType!TemplateID
                    If rsPSType!MBPerPeriod <> -1 Then
                        chkLimit(1).Value = 1
                        Call chkLimit_Click(1)
                        txtMB(1) = rsPSType!MBPerPeriod
                        txtMB(0) = rsPSType!MBBlockSize
                        txtFee(1) = rsPSType!FeePerBlock
                    Else
                        chkLimit(1).Value = 0
                        Call chkLimit_Click(0)
                        txtMB(0) = "0"
                        txtMB(1) = "0"
                        txtFee(1) = "0.00"
                    End If
                    
                    
                    If rsPSType!HoursPerPeriod <> -1 Then
                        chkLimit(0).Value = 1
                        Call chkLimit_Click(0)
                        txtHours = rsPSType!HoursPerPeriod
                        txtFee(2) = rsPSType!ExtraPerHour
                    Else
                        chkLimit(0).Value = 0
                        Call chkLimit_Click(0)
                        txtHours = "0"
                        txtFee(2) = "0.00"
                    End If
                    
                    If Not IsNull(rsPSType!RadiusID) Then
                        Dim ix As Integer
                        For ix = 0 To cmdRadius.ListCount - 1
                            If cmdRadius.ItemData(ix) = rsPSType!RadiusID Then
                               cmdRadius.ListIndex = ix
                               Exit For
                            End If
                        Next
                    End If
                End If
            End Select
            
            chkLimit(2).Value = rsPSType!Rollover
            
            Dim bx As Byte
            Dim lx As Variant
            
            For bx = 0 To 2
                If rsPSType!roIntervalType = optBillingCycle(bx).Tag Then
                    optBillingCycle(bx).Value = True
                    txtBillingCycle(bx).Text = "" & rsPSType!roInterval
                    txtBillingCycle(bx).Locked = False
                    For lx = 0 To cmbRollover.ListCount - 1
                        If cmbRollover.ItemData(lx) = rsPSType!roRecID Then
                            cmbRollover.ListIndex = lx
                            Exit For
                        End If
                    Next
                    Exit For
                End If
            Next
            
            For bx = 3 To 5
                If rsPSType!chgIntervalType = optBillingCycle(bx).Tag Then
                    optBillingCycle(bx).Value = True
                    txtBillingCycle(bx).Text = rsPSType!chgInterval
                    txtBillingCycle(bx).Locked = False
                    Exit For
                End If
            Next
            
            txtDesc.Locked = False
            txtDescription.Locked = False
            LoadFlags rsPSType!RecID
            LoadContracts rsPSType!RecID
            
            cmdSpiff.Enabled = True
            cmdSetup.Enabled = True
            cmdSpiff.Tag = rsPSType!RecID
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

Private Sub lvAccounts_KeyUp(KeyCode As Integer, Shift As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvAccounts_KeyUp"
    Const ContainerName = "frmAccountTypes"
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


    Select Case KeyCode
    Case vbKeyF2
        lvAccounts.StartLabelEdit
    Case vbKeyDelete
    
        If lvAccounts.ListItems.Count > 0 Then
        Select Case MsgBox("Are you sure you wish to delete this Sevice or Plan?", vbYesNo + vbCritical, "Caution")
        Case vbYes
            MySQL.Execute ADOConn, "Delete from plantypes where RecID = " & Mid(lvAccounts.SelectedItem.Key, 2)
            MySQL.Execute ADOConn, "Delete from accountviewer where selectStatement Like '%\'" & Mid(lvAccounts.SelectedItem.Key, 2) & "\'%'"
            lvAccounts.ListItems.Remove lvAccounts.SelectedItem.Index
        End Select
        End If
    Case vbKeyF12
    
            On Error Resume Next
            Do
                Err.Clear
                ltmpRecID2 = MySQL.GetTMPRecID("accountviewer", ADOConn)
                MySQL.Execute ADOConn, "Insert Into accountviewer (RecID, SubofRecID, Description, IconNum, Action, selectStatement, CountStatement, VirtualID) VALUES (" & Val(ltmpRecID2) & ",1,'" & "View " & MySQL.ESC(lvAccounts.SelectedItem.Text) & " Users',40,0,'" & "select distinct accountinfo.* from accountinfo, acci_services Where accountinfo.RecID = acci_services.acci_RecID And acci_services.ptRecID = \'" & Val(Mid(lvAccounts.SelectedItem.Key, 2)) & "\'','" & "select distinct count(*) from accountinfo, acci_services Where accountinfo.RecID = acci_services.acci_RecID And acci_services.ptRecID = \'" & Val(Mid(lvAccounts.SelectedItem.Key, 2)) & "\''," & Login.lVirtualID & ")"
                gSleep
                If Err.Number > 0 Then cDebug Err.Description
            Loop Until Err.Number = 0
            
            Do
                Err.Clear
                ltmpRecID3 = MySQL.GetTMPRecID("accountviewer", ADOConn)
                MySQL.Execute ADOConn, "Insert Into accountviewer (RecID, SubofRecID, Description, IconNum, Action, selectStatement, CountStatement, VirtualID) VALUES (" & ltmpRecID3 & ", " & Val(ltmpRecID2) & ",'Active',41,0,'" & "select distinct accountinfo.* from accountinfo, acci_services Where accountinfo.RecID = acci_services.acci_RecID And acci_services.ptRecID = \'" & Val(Mid(lvAccounts.SelectedItem.Key, 2)) & "\' and accountinfo.Cancelled = 0','" & "select distinct count(*) from accountinfo, acci_services Where accountinfo.RecID = acci_services.acci_RecID And acci_services.ptRecID = \'" & Val(Mid(lvAccounts.SelectedItem.Key, 2)) & "\' and accountinfo.Cancelled = 0'," & Login.lVirtualID & ")"
                gSleep
                If Err.Number > 0 Then cDebug Err.Description
            Loop Until Err.Number = 0
            
            Do
                Err.Clear
                ltmpRecID3 = MySQL.GetTMPRecID("accountviewer", ADOConn)
                MySQL.Execute ADOConn, "Insert Into accountviewer (Recid, SubofRecID, Description, IconNum, Action, selectStatement, CountStatement, VirtualID) VALUES (" & ltmpRecID3 & ", " & Val(ltmpRecID2) & ",'Cancelled',41,0,'" & "select distinct accountinfo.* from accountinfo, acci_services Where accountinfo.RecID = acci_services.acci_RecID And acci_services.ptRecID = \'" & Val(Mid(lvAccounts.SelectedItem.Key, 2)) & "\' and accountinfo.Cancelled <> 0','" & "select distinct count(*) from accountinfo, acci_services Where accountinfo.RecID = acci_services.acci_RecID And acci_services.ptRecID = \'" & Val(ltmpRecID) & "\' and accountinfo.Cancelled <> 0'," & Login.lVirtualID & ")"
                gSleep
                If Err.Number > 0 Then cDebug Err.Description
            Loop Until Err.Number = 0
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

Private Sub lvContracts_DblClick()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvContracts_DblClick"
    Const ContainerName = "frmAccountTypes"
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


    If lvContracts.SelectedItem Is Nothing Then
    
    Else
    
        Dim fEdit As New frmContractEditor
        
        fEdit.txtField(0) = lvContracts.SelectedItem.Text
        fEdit.txtFee(6) = lvContracts.SelectedItem.SubItems(3)
        fEdit.txtFee(7) = lvContracts.SelectedItem.SubItems(5)
        fEdit.txtFee(8) = lvContracts.SelectedItem.SubItems(4)
        fEdit.txtFee(9) = lvContracts.SelectedItem.SubItems(6)
        fEdit.txtFee(10) = lvContracts.SelectedItem.SubItems(7)
        fEdit.RecID = lvContracts.SelectedItem.SubItems(8)
        fEdit.Parnt = Me
        fEdit.Show 1
        
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

Private Sub lvPlans_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvPlans_ColumnClick"
    Const ContainerName = "frmAccountTypes"
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


    Call GUI.ColumnSort(ColumnHeader, lvPlans)
    
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

Private Sub lvPlans_ItemCheck(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvPlans_ItemCheck"
    Const ContainerName = "frmAccountTypes"
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


    If Item.Key <> "" Then Item.Key = "e" & Mid(Item.Key, 2)
    
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

Private Sub optBillingCycle_Click(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "optBillingCycle_Click"
    Const ContainerName = "frmAccountTypes"
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
    
    For bx = txtBillingCycle.LBound To txtBillingCycle.UBound
        If bx = Index Then
            txtBillingCycle(bx).Enabled = True
        Else
            txtBillingCycle(bx).Enabled = False
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

Private Sub Option1_Click(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Option1_Click"
    Const ContainerName = "frmAccountTypes"
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
    
    Select Case Index
    Case 0 To 3
        For ix = 4 To 7
            Option1(ix) = Option1(ix - 4)
        Next
    Case 4 To 7
        For ix = 0 To 3
            Option1(ix) = Option1(ix + 4)
        Next
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

Public Sub tsPlan_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "tsPlan_Click"
    Const ContainerName = "frmAccountTypes"
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
    
    cmdRadius.Enabled = False
    
    Select Case tsPlan.SelectedItem.Index
    Case 1
        frameExtras.Visible = False
        For bx = frameConfig.LBound To frameConfig.UBound
            frameConfig(bx).Visible = False
        Next
        Select Case tvServiceTypes.SelectedItem.Tag
        Case "DIALUP", "ADSL", "SHDSL"
            frameConfig(0).Visible = True
            frameConfig(0).ZOrder 0
            frameConfig(0).Move tsPlan.ClientLeft, tsPlan.ClientTop, tsPlan.ClientWidth, tsPlan.ClientHeight
            cmdRadius.Enabled = True
        Case "FTP", "POP3", "WWW", "DESIGN", "DOMAIN", "GATEWAY", "TRAINING", "SALES", "CONSULT", "COLO", "HOST", "ALIAS"
            frameConfig(1).Visible = True
            frameConfig(1).ZOrder 0
            frameConfig(1).Move tsPlan.ClientLeft, tsPlan.ClientTop, tsPlan.ClientWidth, tsPlan.ClientHeight
        End Select
    Case 2
        frameExtras.Visible = True
        frameExtras.ZOrder 0
        frameExtras.Move tsPlan.ClientLeft, tsPlan.ClientTop, tsPlan.ClientWidth, tsPlan.ClientHeight
        For bx = frameConfig.LBound To frameConfig.UBound
            frameConfig(bx).Visible = False
        Next
    Case 3
        frameExtras.Visible = False
        For bx = frameConfig.LBound To frameConfig.UBound
            frameConfig(bx).Visible = False
        Next
        frameConfig(2).Visible = True
        frameConfig(2).ZOrder 0
        frameConfig(2).Move tsPlan.ClientLeft, tsPlan.ClientTop, tsPlan.ClientWidth, tsPlan.ClientHeight
        
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

Private Sub tvservicetypes_NodeClick(ByVal Node As MSComctlLib.Node)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "tvservicetypes_NodeClick"
    Const ContainerName = "frmAccountTypes"
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


    On Error Resume Next
    
    Dim rsPSType As ADODB.Recordset
    Dim bResult As Boolean
    Dim itmX As ListItem
    lvAccounts.ListItems.Clear
    
    bResult = MySQL.SetColumnHeaders("accounttype", lvAccounts, Node.Tag, ADOConn)
        
    frameServices.Caption = "Services and Plans: " & Node.Text
    
    lvPlans.ListItems.Clear
    
    
    bResult = MySQL.OpenTable(ADOConn, rsPSType, , MySQL.virtualisp("select * from plantypes Where plantypes.ServiceID = '" & Mid(Node.Key, 2) + "'", "plantypes", True))
    
    Call MySQL.fillLV(ADOConn, rsPSType, lvAccounts, False)

    Call tsPlan_Click
    
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

Private Sub txtBillingCycle_KeyPress(Index As Integer, KeyAscii As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtBillingCycle_KeyPress"
    Const ContainerName = "frmAccountTypes"
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
    Case 8
    Case Asc("0") To Asc("9")
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

Private Sub txtCat_Change(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtCat_Change"
    Const ContainerName = "frmAccountTypes"
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
    Case 0
        If txtCat(1) <> txtCat(0) Then txtCat(1) = txtCat(0)
    Case 1
        If txtCat(1) <> txtCat(0) Then txtCat(0) = txtCat(1)
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

Private Sub txtDesc_Change()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtDesc_Change"
    Const ContainerName = "frmAccountTypes"
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


    If Not txtDescription = txtDesc Then txtDescription = txtDesc
    
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

Private Sub txtDescription_Change()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtDescription_Change"
    Const ContainerName = "frmAccountTypes"
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


    If Not txtDesc = txtDescription Then txtDesc = txtDescription
 
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

    txtFee(Index).ToolTipText = "GST Inclusive Price " & Format(Val(txtFee(Index)) * oTax(Login.TaxCode, Login.TaxCountry) + Val(txtFee(Index)), "Currency")
    
End Sub

Private Sub txtFee_DblClick(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtFee_DblClick"
    Const ContainerName = "frmAccountTypes"
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

Private Sub txtFee_GotFocus(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtFee_GotFocus"
    Const ContainerName = "frmAccountTypes"
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


    txtFee(Index).SelStart = 0
    txtFee(Index).SelLength = Len(txtFee(Index).Text)
    
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
    Const ContainerName = "frmAccountTypes"
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
    Case 8
    Case 48 To 57
    Case Asc(".")
        If InStr(txtFee(Index), ".") > 0 Then KeyAscii = 0
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


Private Sub txtHours_KeyPress(KeyAscii As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtHours_KeyPress"
    Const ContainerName = "frmAccountTypes"
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
    Case 8
    Case 48 To 57
    Case Asc(".")
        If InStr(txtHours, ".") > 0 Then KeyAscii = 0
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

Private Sub txtMB_KeyPress(Index As Integer, KeyAscii As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtMB_KeyPress"
    Const ContainerName = "frmAccountTypes"
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
    Case 8
    Case 48 To 57
    Case Asc(".")
        If InStr(txtMB(Index), ".") > 0 Then KeyAscii = 0
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

Public Function LoadFlags(lRecID As Variant)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "LoadFlags"
    Const ContainerName = "frmAccountTypes"
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
    Dim itmX As ListItem
    
    bResult = MySQL.OpenTable(ADOConn, rsload, , "select plantypes.Description, flags_plantype.* from plantypes, flags_plantype Where flags_plantype.PlanType = plantypes.RecID AND ptRecID = " & lRecID)
    
    lvPlans.ListItems.Clear
    
    If rsload.RecordCount > 0 Then
        While Not rsload.EOF And Err.Number = 0
            Set itmX = lvPlans.ListItems.Add(, "x" & rsload!RecID, rsload!Description)
            itmX.Tag = rsload!PlanType
            itmX.SubItems(1) = rsload!NumberOf
            itmX.Checked = rsload!Checked
            rsload.MoveNext
        Wend
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


Public Function SaveFlags(lRecID As Variant)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "SaveFlags"
    Const ContainerName = "frmAccountTypes"
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


    Dim itmX As ListItem
    Dim sa As Integer
    Dim rsSave As ADODB.Recordset
    
    If lvPlans.ListItems.Count > 0 Then
    
        For sa = 1 To lvPlans.ListItems.Count
            Set itmX = lvPlans.ListItems(sa)
            If Left(itmX.Key, 1) = "r" Then
                
                Call MySQL.OpenTable(ADOConn, rsSave, , "select * from flags_plantype Limit 0")
                
                rsSave.AddNew
                rsSave!ptRecID = lRecID
                rsSave!PlanType = itmX.Tag
                rsSave!NumberOf = Val(itmX.SubItems(1))
                rsSave!Checked = itmX.Checked
                
                rsSave.Update
                                
            ElseIf Left(itmX.Key, 1) = "e" Then
            
                Call MySQL.OpenTable(ADOConn, rsSave, , "select * from flags_plantype where RecID = " & Mid(itmX.Key, 2) & " Limit 1")
                
                rsSave!ptRecID = lRecID
                rsSave!PlanType = itmX.Tag
                rsSave!NumberOf = Val(itmX.SubItems(1))
                rsSave!Checked = itmX.Checked
                
                rsSave.Update
                
            End If
        Next
    
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

Public Sub LoadContracts(ptRecID As Variant)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "LoadContracts"
    Const ContainerName = "frmAccountTypes"
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


    Dim rsload As ADODB.Recordset


        Call MySQL.OpenTable(ADOConn, rsload, , "select contracttemplates.Description, contracttemplates.NoPeriods, contracttemplates.TypePeriods, contractsruntime.* from contracttemplates, contractsruntime where contractsruntime.ContractID = contracttemplates.RecID and contractsruntime.ptRecID = " & ptRecID & " and contracttemplates.bDeleted = 0")
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
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Sub
