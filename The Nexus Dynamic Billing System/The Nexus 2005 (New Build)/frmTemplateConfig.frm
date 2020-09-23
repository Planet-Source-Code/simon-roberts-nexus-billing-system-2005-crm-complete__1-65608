VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTemplateConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Templates Configuration"
   ClientHeight    =   12435
   ClientLeft      =   5430
   ClientTop       =   960
   ClientWidth     =   13605
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTemplateConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   829
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   907
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Service and Plans"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   12315
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   13395
      Begin VB.CommandButton Command2 
         Caption         =   "Link Template"
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
         Height          =   345
         Left            =   240
         TabIndex        =   6
         Top             =   7560
         Width           =   1785
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2130
         TabIndex        =   3
         Top             =   7950
         Width           =   1785
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   4
         Top             =   7950
         Width           =   1785
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2130
         TabIndex        =   2
         Top             =   7560
         Width           =   1785
      End
      Begin MSComctlLib.TreeView tvServiceTypes 
         Height          =   7185
         Left            =   150
         TabIndex        =   1
         Top             =   270
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   12674
         _Version        =   393217
         Indentation     =   441
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         HotTracking     =   -1  'True
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
      End
      Begin MSComctlLib.ListView lvAccounts 
         Height          =   3795
         Left            =   240
         TabIndex        =   5
         Top             =   8430
         Width           =   13020
         _ExtentX        =   22966
         _ExtentY        =   6694
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   16777215
         BackColor       =   0
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "!Description!^^"
            Text            =   "Description"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "!VendorID!^select vname as nResult from vendors where RecID = ^"
            Text            =   "Vendor"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "!VendorPartID!^^"
            Text            =   "Vendor Code"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "!SubPartID!^^"
            Text            =   "Sub Part Number"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "!PeriodFee!^Currency^"
            Text            =   "Cycle Fee (ex GST)"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "!iTax+PeriodFee!^Currency^"
            Text            =   "Cycle Fee (inc GST)"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "!MBPerPeriod!^unlimitedMB^"
            Text            =   "Cycle Data"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Tag             =   "!HoursPerPeriod!^unlimitedHR^"
            Text            =   "Cycle Time"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Tag             =   "!SessionTimeout!^###,###,###,### sec^"
            Text            =   "Session Time"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Object.Tag             =   "!IdleTimeout!^###,###,##,### sec^"
            Text            =   "Idle"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame frmts 
         BackColor       =   &H0092BA5A&
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
         Height          =   7635
         Index           =   0
         Left            =   4230
         TabIndex        =   7
         Top             =   690
         Width           =   9015
         Begin VB.PictureBox picSet 
            Appearance      =   0  'Flat
            BackColor       =   &H0092BA5A&
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
            ForeColor       =   &H80000008&
            Height          =   3975
            Index           =   2
            Left            =   120
            ScaleHeight     =   3975
            ScaleWidth      =   8775
            TabIndex        =   107
            Tag             =   "CONSULT, TRAINING, DESIGN"
            Top             =   1110
            Visible         =   0   'False
            Width           =   8775
            Begin VB.Frame Frame10 
               BackColor       =   &H009191F4&
               Caption         =   "Requirement or Conditions of Service"
               Height          =   2145
               Left            =   2820
               TabIndex        =   124
               Top             =   1740
               Width           =   6015
               Begin MSComctlLib.ListView lvREQ 
                  Height          =   1815
                  Left            =   90
                  TabIndex        =   125
                  Top             =   240
                  Width           =   5835
                  _ExtentX        =   10292
                  _ExtentY        =   3201
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   -1  'True
                  HideSelection   =   -1  'True
                  FullRowSelect   =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   9540084
                  Appearance      =   0
                  NumItems        =   1
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "Requirement"
                     Object.Width           =   10583
                  EndProperty
               End
            End
            Begin VB.Frame Frame9 
               BackColor       =   &H00A26F74&
               Caption         =   "Location"
               ForeColor       =   &H00FFFFFF&
               Height          =   705
               Left            =   2790
               TabIndex        =   122
               Top             =   960
               Width           =   3405
               Begin VB.TextBox txtLoc 
                  Alignment       =   2  'Center
                  BackColor       =   &H00A26F74&
                  ForeColor       =   &H00FFFFFF&
                  Height          =   345
                  Left            =   120
                  TabIndex        =   123
                  Top             =   240
                  Width           =   3165
               End
            End
            Begin VB.Frame Frame8 
               BackColor       =   &H00AF94D3&
               Caption         =   "Rate of Charge"
               Height          =   2955
               Left            =   60
               TabIndex        =   113
               Top             =   930
               Width           =   2655
               Begin VB.OptionButton optRates 
                  BackColor       =   &H00AF94D3&
                  Caption         =   "A year per block fee"
                  Enabled         =   0   'False
                  Height          =   255
                  Index           =   6
                  Left            =   150
                  TabIndex        =   121
                  Top             =   2610
                  Width           =   2385
               End
               Begin VB.OptionButton optRates 
                  BackColor       =   &H00AF94D3&
                  Caption         =   "A quarter per block fee"
                  Enabled         =   0   'False
                  Height          =   255
                  Index           =   5
                  Left            =   150
                  TabIndex        =   120
                  Top             =   2280
                  Width           =   2385
               End
               Begin VB.OptionButton optRates 
                  BackColor       =   &H00AF94D3&
                  Caption         =   "A month per block fee"
                  Enabled         =   0   'False
                  Height          =   255
                  Index           =   4
                  Left            =   150
                  TabIndex        =   119
                  Top             =   1950
                  Width           =   2385
               End
               Begin VB.OptionButton optRates 
                  BackColor       =   &H00AF94D3&
                  Caption         =   "One week per block fee"
                  Enabled         =   0   'False
                  Height          =   255
                  Index           =   3
                  Left            =   150
                  TabIndex        =   118
                  Top             =   1620
                  Width           =   2385
               End
               Begin VB.OptionButton optRates 
                  BackColor       =   &H00AF94D3&
                  Caption         =   "One day per block fee"
                  Enabled         =   0   'False
                  Height          =   255
                  Index           =   2
                  Left            =   150
                  TabIndex        =   117
                  Top             =   1290
                  Width           =   2385
               End
               Begin VB.OptionButton optRates 
                  BackColor       =   &H00AF94D3&
                  Caption         =   "60 min block per fee"
                  Enabled         =   0   'False
                  Height          =   255
                  Index           =   1
                  Left            =   150
                  TabIndex        =   116
                  Top             =   960
                  Width           =   2385
               End
               Begin VB.OptionButton optRates 
                  BackColor       =   &H00AF94D3&
                  Caption         =   "30 min block per fee"
                  Enabled         =   0   'False
                  Height          =   255
                  Index           =   0
                  Left            =   150
                  TabIndex        =   115
                  Top             =   630
                  Width           =   2385
               End
               Begin VB.CheckBox chkRates 
                  BackColor       =   &H00AF94D3&
                  Caption         =   "Charge by selected rate"
                  Height          =   225
                  Left            =   150
                  TabIndex        =   114
                  Top             =   300
                  Width           =   2385
               End
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00E7EB52&
               Caption         =   "Vendor"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Index           =   2
               Left            =   60
               TabIndex        =   111
               Top             =   120
               Width           =   6135
               Begin VB.ComboBox cmbVendors 
                  BackColor       =   &H00F2F2C6&
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
                  Index           =   2
                  Left            =   120
                  Style           =   2  'Dropdown List
                  TabIndex        =   112
                  Top             =   210
                  Width           =   5895
               End
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H0075726F&
               Caption         =   "Data Quota (MB's)"
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
               Height          =   735
               Index           =   2
               Left            =   6270
               TabIndex        =   109
               Top             =   120
               Width           =   2565
               Begin VB.TextBox txtQuota 
                  Alignment       =   2  'Center
                  BackColor       =   &H0075726F&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   345
                  Index           =   2
                  Left            =   120
                  TabIndex        =   110
                  Text            =   "20"
                  Top             =   270
                  Width           =   2325
               End
            End
            Begin VB.CommandButton Command1 
               BackColor       =   &H00A9F5AE&
               Caption         =   "Product Text"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Index           =   2
               Left            =   6300
               MaskColor       =   &H0092BA5A&
               Style           =   1  'Graphical
               TabIndex        =   108
               Top             =   960
               Width           =   2535
            End
         End
         Begin VB.PictureBox picSet 
            Appearance      =   0  'Flat
            BackColor       =   &H0092BA5A&
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
            ForeColor       =   &H80000008&
            Height          =   1125
            Index           =   1
            Left            =   2910
            ScaleHeight     =   1125
            ScaleWidth      =   1515
            TabIndex        =   101
            Tag             =   "ALIAS, HOST, COLO, GATEWAY, WWW, POP3, FTP, DOMAIN"
            Top             =   990
            Visible         =   0   'False
            Width           =   1515
            Begin VB.CommandButton Command1 
               BackColor       =   &H00A9F5AE&
               Caption         =   "Product Text"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Index           =   1
               Left            =   6300
               MaskColor       =   &H0092BA5A&
               Style           =   1  'Graphical
               TabIndex        =   106
               Top             =   960
               Width           =   2535
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H0075726F&
               Caption         =   "Data Quota (MB's)"
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
               Height          =   735
               Index           =   1
               Left            =   6270
               TabIndex        =   104
               Top             =   120
               Width           =   2565
               Begin VB.TextBox txtQuota 
                  Alignment       =   2  'Center
                  BackColor       =   &H0075726F&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   345
                  Index           =   1
                  Left            =   120
                  TabIndex        =   105
                  Text            =   "20"
                  Top             =   270
                  Width           =   2325
               End
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00E7EB52&
               Caption         =   "Vendor"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Index           =   1
               Left            =   60
               TabIndex        =   102
               Top             =   120
               Width           =   6135
               Begin VB.ComboBox cmbVendors 
                  BackColor       =   &H00F2F2C6&
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
                  Index           =   1
                  Left            =   120
                  Style           =   2  'Dropdown List
                  TabIndex        =   103
                  Top             =   210
                  Width           =   5895
               End
            End
         End
         Begin VB.PictureBox picSet 
            Appearance      =   0  'Flat
            BackColor       =   &H0092BA5A&
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
            ForeColor       =   &H80000008&
            Height          =   855
            Index           =   3
            Left            =   60
            ScaleHeight     =   855
            ScaleWidth      =   1245
            TabIndex        =   126
            Tag             =   "SALES"
            Top             =   1020
            Visible         =   0   'False
            Width           =   1245
            Begin VB.Frame Frame14 
               BackColor       =   &H0084E8E8&
               Caption         =   "Item Constraints and Definitions"
               Height          =   2865
               Left            =   2610
               TabIndex        =   149
               Top             =   960
               Width           =   3585
               Begin VB.CheckBox chkCont 
                  BackColor       =   &H0084E8E8&
                  Caption         =   "Tools"
                  Height          =   225
                  Index           =   17
                  Left            =   1830
                  TabIndex        =   167
                  Top             =   2490
                  Width           =   1635
               End
               Begin VB.CheckBox chkCont 
                  BackColor       =   &H0084E8E8&
                  Caption         =   "Adult material"
                  Height          =   225
                  Index           =   16
                  Left            =   150
                  TabIndex        =   166
                  Top             =   2490
                  Width           =   1635
               End
               Begin VB.CheckBox chkCont 
                  BackColor       =   &H0084E8E8&
                  Caption         =   "Hardware"
                  Height          =   225
                  Index           =   15
                  Left            =   1830
                  TabIndex        =   165
                  Top             =   2220
                  Width           =   1635
               End
               Begin VB.CheckBox chkCont 
                  BackColor       =   &H0084E8E8&
                  Caption         =   "Membership"
                  Height          =   225
                  Index           =   14
                  Left            =   150
                  TabIndex        =   164
                  Top             =   2220
                  Width           =   1635
               End
               Begin VB.CheckBox chkCont 
                  BackColor       =   &H0084E8E8&
                  Caption         =   "Software"
                  Height          =   225
                  Index           =   13
                  Left            =   1830
                  TabIndex        =   163
                  Top             =   1950
                  Width           =   1635
               End
               Begin VB.CheckBox chkCont 
                  BackColor       =   &H0084E8E8&
                  Caption         =   "Printed media"
                  Height          =   225
                  Index           =   12
                  Left            =   150
                  TabIndex        =   162
                  Top             =   1950
                  Width           =   1635
               End
               Begin VB.CheckBox chkCont 
                  BackColor       =   &H0084E8E8&
                  Caption         =   "Food Stuffs"
                  Height          =   225
                  Index           =   11
                  Left            =   1830
                  TabIndex        =   161
                  Top             =   1680
                  Width           =   1635
               End
               Begin VB.CheckBox chkCont 
                  BackColor       =   &H0084E8E8&
                  Caption         =   "Comsumable"
                  Height          =   225
                  Index           =   10
                  Left            =   150
                  TabIndex        =   160
                  Top             =   1680
                  Width           =   1635
               End
               Begin VB.CheckBox chkCont 
                  BackColor       =   &H0084E8E8&
                  Caption         =   "Component"
                  Height          =   225
                  Index           =   9
                  Left            =   1830
                  TabIndex        =   159
                  Top             =   1410
                  Width           =   1635
               End
               Begin VB.CheckBox chkCont 
                  BackColor       =   &H0084E8E8&
                  Caption         =   "Toy"
                  Height          =   225
                  Index           =   8
                  Left            =   150
                  TabIndex        =   158
                  Top             =   1410
                  Width           =   1635
               End
               Begin VB.CheckBox chkCont 
                  BackColor       =   &H0084E8E8&
                  Caption         =   "Air freight"
                  Height          =   225
                  Index           =   7
                  Left            =   1830
                  TabIndex        =   157
                  Top             =   1140
                  Width           =   1635
               End
               Begin VB.CheckBox chkCont 
                  BackColor       =   &H0084E8E8&
                  Caption         =   "Road freight"
                  Height          =   225
                  Index           =   6
                  Left            =   150
                  TabIndex        =   156
                  Top             =   1140
                  Width           =   1635
               End
               Begin VB.CheckBox chkCont 
                  BackColor       =   &H0084E8E8&
                  Caption         =   "Chemical"
                  Height          =   225
                  Index           =   5
                  Left            =   1830
                  TabIndex        =   155
                  Top             =   870
                  Width           =   1635
               End
               Begin VB.CheckBox chkCont 
                  BackColor       =   &H0084E8E8&
                  Caption         =   "Hazardous"
                  Height          =   225
                  Index           =   4
                  Left            =   150
                  TabIndex        =   154
                  Top             =   870
                  Width           =   1635
               End
               Begin VB.CheckBox chkCont 
                  BackColor       =   &H0084E8E8&
                  Caption         =   "Accessory"
                  Height          =   225
                  Index           =   3
                  Left            =   1830
                  TabIndex        =   153
                  Top             =   600
                  Width           =   1635
               End
               Begin VB.CheckBox chkCont 
                  BackColor       =   &H0084E8E8&
                  Caption         =   "Botanical"
                  Height          =   225
                  Index           =   2
                  Left            =   150
                  TabIndex        =   152
                  Top             =   600
                  Width           =   1635
               End
               Begin VB.CheckBox chkCont 
                  BackColor       =   &H0084E8E8&
                  Caption         =   "Electrical"
                  Height          =   225
                  Index           =   1
                  Left            =   1830
                  TabIndex        =   151
                  Top             =   330
                  Width           =   1635
               End
               Begin VB.CheckBox chkCont 
                  BackColor       =   &H0084E8E8&
                  Caption         =   "Fragile"
                  Height          =   225
                  Index           =   0
                  Left            =   150
                  TabIndex        =   150
                  Top             =   330
                  Width           =   1635
               End
            End
            Begin VB.Frame Frame13 
               BackColor       =   &H00BA3F3F&
               Caption         =   "Units per package"
               ForeColor       =   &H00FFFFFF&
               Height          =   765
               Left            =   120
               TabIndex        =   147
               Top             =   3060
               Width           =   2415
               Begin VB.TextBox txtUnits 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00BA3F3F&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   405
                  Left            =   90
                  TabIndex        =   148
                  Text            =   "1"
                  Top             =   270
                  Width           =   2265
               End
            End
            Begin VB.Frame Frame12 
               BackColor       =   &H0066B0DD&
               Caption         =   "Packaging"
               Height          =   1875
               Left            =   90
               TabIndex        =   141
               Top             =   960
               Width           =   2445
               Begin VB.ComboBox cmbPack 
                  BackColor       =   &H0066B0DD&
                  Height          =   345
                  Left            =   390
                  TabIndex        =   168
                  Text            =   "No Packaging"
                  Top             =   240
                  Width           =   1965
               End
               Begin VB.OptionButton optPack 
                  BackColor       =   &H0066B0DD&
                  Caption         =   "Plastic Box"
                  Height          =   255
                  Index           =   4
                  Left            =   150
                  TabIndex        =   146
                  Top             =   1530
                  Width           =   2175
               End
               Begin VB.OptionButton optPack 
                  BackColor       =   &H0066B0DD&
                  Caption         =   "Metal Box"
                  Height          =   255
                  Index           =   3
                  Left            =   150
                  TabIndex        =   145
                  Top             =   1230
                  Width           =   2175
               End
               Begin VB.OptionButton optPack 
                  BackColor       =   &H0066B0DD&
                  Caption         =   "Cardboard Box"
                  Height          =   255
                  Index           =   2
                  Left            =   150
                  TabIndex        =   144
                  Top             =   960
                  Width           =   2175
               End
               Begin VB.OptionButton optPack 
                  BackColor       =   &H0066B0DD&
                  Caption         =   "Blister Pack"
                  Height          =   255
                  Index           =   1
                  Left            =   150
                  TabIndex        =   143
                  Top             =   660
                  Width           =   2175
               End
               Begin VB.OptionButton optPack 
                  BackColor       =   &H0066B0DD&
                  Height          =   255
                  Index           =   0
                  Left            =   150
                  TabIndex        =   142
                  Top             =   300
                  Width           =   2175
               End
            End
            Begin VB.Frame Frame11 
               BackColor       =   &H00A9C9AE&
               Caption         =   "Dimension's && Weights"
               Height          =   1665
               Index           =   0
               Left            =   6300
               TabIndex        =   132
               Top             =   2160
               Width           =   2505
               Begin VB.TextBox txtDem 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00A9C9AE&
                  Height          =   255
                  Index           =   3
                  Left            =   900
                  TabIndex        =   140
                  Top             =   1260
                  Width           =   1485
               End
               Begin VB.TextBox txtDem 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00A9C9AE&
                  Height          =   255
                  Index           =   2
                  Left            =   900
                  TabIndex        =   138
                  Top             =   960
                  Width           =   1485
               End
               Begin VB.TextBox txtDem 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00A9C9AE&
                  Height          =   255
                  Index           =   1
                  Left            =   900
                  TabIndex        =   136
                  Top             =   660
                  Width           =   1485
               End
               Begin VB.TextBox txtDem 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00A9C9AE&
                  Height          =   255
                  Index           =   0
                  Left            =   900
                  TabIndex        =   134
                  Top             =   360
                  Width           =   1485
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Weight:"
                  Height          =   225
                  Index           =   3
                  Left            =   180
                  TabIndex        =   139
                  Top             =   1260
                  Width           =   615
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Depth:"
                  Height          =   225
                  Index           =   2
                  Left            =   180
                  TabIndex        =   137
                  Top             =   960
                  Width           =   540
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Length:"
                  Height          =   225
                  Index           =   1
                  Left            =   180
                  TabIndex        =   135
                  Top             =   660
                  Width           =   615
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Height:"
                  Height          =   225
                  Index           =   0
                  Left            =   180
                  TabIndex        =   133
                  Top             =   360
                  Width           =   585
               End
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00E7EB52&
               Caption         =   "Vendor"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Index           =   3
               Left            =   60
               TabIndex        =   130
               Top             =   120
               Width           =   6135
               Begin VB.ComboBox cmbVendors 
                  BackColor       =   &H00F2F2C6&
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
                  Index           =   3
                  Left            =   120
                  Style           =   2  'Dropdown List
                  TabIndex        =   131
                  Top             =   210
                  Width           =   5895
               End
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H0075726F&
               Caption         =   "Data Quota (MB's)"
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
               Height          =   735
               Index           =   3
               Left            =   6270
               TabIndex        =   128
               Top             =   120
               Width           =   2565
               Begin VB.TextBox txtQuota 
                  Alignment       =   2  'Center
                  BackColor       =   &H0075726F&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   345
                  Index           =   3
                  Left            =   120
                  TabIndex        =   129
                  Text            =   "20"
                  Top             =   270
                  Width           =   2325
               End
            End
            Begin VB.CommandButton Command1 
               BackColor       =   &H00A9F5AE&
               Caption         =   "Product Text"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Index           =   3
               Left            =   6300
               MaskColor       =   &H0092BA5A&
               Style           =   1  'Graphical
               TabIndex        =   127
               Top             =   960
               Width           =   2535
            End
         End
         Begin VB.Frame frameExtras 
            BackColor       =   &H0092BA5A&
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
            Height          =   2475
            Index           =   0
            Left            =   150
            TabIndex        =   49
            Top             =   5010
            Width           =   8775
            Begin VB.CommandButton cmdAddPlan 
               Caption         =   "+"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   8190
               TabIndex        =   53
               Top             =   210
               Width           =   405
            End
            Begin VB.TextBox txtNumOf 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00A9F5AE&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   0
               Left            =   7170
               TabIndex        =   52
               Text            =   "1"
               Top             =   210
               Width           =   975
            End
            Begin VB.ComboBox cmbAllPLans 
               BackColor       =   &H00A9F5AE&
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
               Index           =   0
               Left            =   1920
               Style           =   2  'Dropdown List
               TabIndex        =   51
               Top             =   240
               Width           =   5175
            End
            Begin VB.CommandButton cmdPlansConfigure 
               Caption         =   "Plans Configuration"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Index           =   0
               Left            =   150
               TabIndex        =   50
               Top             =   1950
               Width           =   1665
            End
            Begin MSComctlLib.ListView lvPlans 
               Height          =   1665
               Index           =   0
               Left            =   1920
               TabIndex        =   54
               Top             =   660
               Width           =   6735
               _ExtentX        =   11880
               _ExtentY        =   2937
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               Checkboxes      =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               HotTracking     =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   11138478
               BorderStyle     =   1
               Appearance      =   1
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
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Plan Type to Include:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   150
               TabIndex        =   55
               Top             =   300
               Width           =   1560
            End
         End
         Begin VB.TextBox txtDesc 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   5340
            MaxLength       =   128
            TabIndex        =   12
            Top             =   270
            Width           =   3525
         End
         Begin VB.TextBox txtFee 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Index           =   0
            Left            =   5340
            TabIndex        =   11
            Top             =   660
            Width           =   1365
         End
         Begin VB.TextBox txtFee 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Index           =   3
            Left            =   7770
            TabIndex        =   10
            Top             =   660
            Width           =   1095
         End
         Begin VB.TextBox txtPartID 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   1770
            MaxLength       =   30
            TabIndex        =   9
            Top             =   270
            Width           =   1485
         End
         Begin VB.TextBox txtSubPartID 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   1770
            MaxLength       =   30
            TabIndex        =   8
            Top             =   660
            Width           =   1485
         End
         Begin VB.PictureBox picSet 
            Appearance      =   0  'Flat
            BackColor       =   &H0092BA5A&
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
            ForeColor       =   &H80000008&
            Height          =   795
            Index           =   0
            Left            =   180
            ScaleHeight     =   795
            ScaleWidth      =   1395
            TabIndex        =   67
            Tag             =   "DIALUP, ADSL, SHDSL"
            Top             =   120
            Visible         =   0   'False
            Width           =   1395
            Begin VB.CheckBox chkHidden 
               BackColor       =   &H0092BA5A&
               Caption         =   "Hidden from VISP's"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   705
               Left            =   5970
               TabIndex        =   100
               Top             =   3210
               Visible         =   0   'False
               Width           =   2835
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H0075726F&
               Caption         =   "Data Quota (MB's)"
               ForeColor       =   &H00FFFFFF&
               Height          =   735
               Index           =   0
               Left            =   3300
               TabIndex        =   98
               Top             =   3180
               Width           =   2565
               Begin VB.TextBox txtQuota 
                  Alignment       =   2  'Center
                  BackColor       =   &H0075726F&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   345
                  Index           =   0
                  Left            =   120
                  TabIndex        =   99
                  Text            =   "20"
                  Top             =   270
                  Width           =   2325
               End
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00E7EB52&
               Caption         =   "Vendor"
               Height          =   705
               Index           =   0
               Left            =   3300
               TabIndex        =   96
               Top             =   90
               Width           =   5565
               Begin VB.ComboBox cmbVendors 
                  BackColor       =   &H00F2F2C6&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   390
                  Index           =   0
                  Left            =   120
                  Style           =   2  'Dropdown List
                  TabIndex        =   97
                  Top             =   210
                  Width           =   5295
               End
            End
            Begin VB.Frame frameLimit 
               BackColor       =   &H0031FB64&
               Caption         =   "Radius Session Timeouts"
               Height          =   1275
               Index           =   2
               Left            =   3300
               TabIndex        =   89
               Top             =   1800
               Width           =   3435
               Begin VB.TextBox txtIdleTimeout 
                  BackColor       =   &H00A9F5AE&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   1650
                  TabIndex        =   91
                  Text            =   "600"
                  Top             =   750
                  Width           =   1335
               End
               Begin VB.TextBox txtSessionTimeout 
                  BackColor       =   &H00A9F5AE&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   1650
                  TabIndex        =   90
                  Text            =   "10800"
                  Top             =   300
                  Width           =   1335
               End
               Begin MSComCtl2.UpDown UpDown2 
                  Height          =   360
                  Left            =   2985
                  TabIndex        =   92
                  Top             =   750
                  Width           =   240
                  _ExtentX        =   423
                  _ExtentY        =   635
                  _Version        =   393216
                  Value           =   600
                  BuddyControl    =   "lvAccounts"
                  BuddyDispid     =   196673
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
                  Left            =   2985
                  TabIndex        =   93
                  Top             =   300
                  Width           =   240
                  _ExtentX        =   423
                  _ExtentY        =   635
                  _Version        =   393216
                  Value           =   10800
                  BuddyControl    =   "tvServiceTypes"
                  BuddyDispid     =   196672
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
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   8
                  Left            =   480
                  TabIndex        =   95
                  Top             =   780
                  Width           =   1110
               End
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Session Timeout:"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   7
                  Left            =   90
                  TabIndex        =   94
                  Top             =   360
                  Width           =   1515
               End
            End
            Begin VB.Frame frameLimit 
               BackColor       =   &H0039F2F2&
               Caption         =   "Time Limits"
               ForeColor       =   &H00000000&
               Height          =   1605
               Index           =   1
               Left            =   90
               TabIndex        =   81
               Top             =   2310
               Width           =   3105
               Begin VB.TextBox txtFee 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H0075CDE3&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   5
                  Left            =   1590
                  TabIndex        =   85
                  Top             =   900
                  Width           =   1425
               End
               Begin VB.CheckBox chkLimit 
                  BackColor       =   &H0039F2F2&
                  Caption         =   "Set Time Limits on this template"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   0
                  Left            =   120
                  TabIndex        =   84
                  Top             =   210
                  Width           =   2835
               End
               Begin VB.TextBox txtHours 
                  Appearance      =   0  'Flat
                  BackColor       =   &H0075CDE3&
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Left            =   1590
                  TabIndex        =   83
                  Top             =   570
                  Width           =   1425
               End
               Begin VB.TextBox txtFee 
                  Appearance      =   0  'Flat
                  BackColor       =   &H0075CDE3&
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   2
                  Left            =   1590
                  TabIndex        =   82
                  Top             =   1230
                  Width           =   1425
               End
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackColor       =   &H0039F2F2&
                  Caption         =   "Cost Price:"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   210
                  Index           =   11
                  Left            =   750
                  TabIndex        =   88
                  Top             =   930
                  Width           =   780
               End
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackColor       =   &H0039F2F2&
                  Caption         =   "Hours Per Cycle:"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   210
                  Index           =   6
                  Left            =   330
                  TabIndex        =   87
                  Top             =   600
                  Width           =   1215
               End
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackColor       =   &H0039F2F2&
                  Caption         =   "Fee Per Extra Hour:"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   210
                  Index           =   5
                  Left            =   120
                  TabIndex        =   86
                  Top             =   1230
                  Width           =   1410
               End
            End
            Begin VB.Frame frameLimit 
               BackColor       =   &H007B7EF0&
               Caption         =   "Data Limits"
               ForeColor       =   &H00FFFFFF&
               Height          =   2145
               Index           =   0
               Left            =   90
               TabIndex        =   71
               Top             =   90
               Width           =   3105
               Begin VB.TextBox txtFee 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00A3A3FE&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   4
                  Left            =   1530
                  TabIndex        =   76
                  Top             =   990
                  Width           =   1425
               End
               Begin VB.TextBox txtFee 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00A3A3FE&
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   1
                  Left            =   1530
                  TabIndex        =   75
                  Top             =   1650
                  Width           =   1425
               End
               Begin VB.TextBox txtMB 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00A3A3FE&
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   0
                  Left            =   1530
                  TabIndex        =   74
                  Top             =   660
                  Width           =   1425
               End
               Begin VB.TextBox txtMB 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00A3A3FE&
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   1
                  Left            =   1530
                  TabIndex        =   73
                  Top             =   1320
                  Width           =   1425
               End
               Begin VB.CheckBox chkLimit 
                  BackColor       =   &H007B7EF0&
                  Caption         =   "Set Data Limits on this template"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   1
                  Left            =   180
                  TabIndex        =   72
                  Top             =   270
                  Width           =   2775
               End
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Cost Price:"
                  Height          =   225
                  Index           =   10
                  Left            =   570
                  TabIndex        =   80
                  Top             =   990
                  Width           =   900
               End
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Fee Per Block:"
                  Height          =   225
                  Index           =   4
                  Left            =   240
                  TabIndex        =   79
                  Top             =   1650
                  Width           =   1185
               End
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "MB's Per Block:"
                  Height          =   225
                  Index           =   3
                  Left            =   180
                  TabIndex        =   78
                  Top             =   660
                  Width           =   1260
               End
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "MB's Per Month:"
                  Height          =   225
                  Index           =   2
                  Left            =   150
                  TabIndex        =   77
                  Top             =   1320
                  Width           =   1305
               End
            End
            Begin VB.Frame Frame5 
               BackColor       =   &H00B14BAA&
               Caption         =   "Radius Pool to use"
               ForeColor       =   &H0084E8E8&
               Height          =   765
               Left            =   3300
               TabIndex        =   69
               Top             =   900
               Width           =   5565
               Begin VB.ComboBox cmdRadius 
                  BackColor       =   &H00EAB9E2&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   390
                  Left            =   120
                  Style           =   2  'Dropdown List
                  TabIndex        =   70
                  Top             =   240
                  Width           =   5295
               End
            End
            Begin VB.CommandButton Command1 
               BackColor       =   &H00A9F5AE&
               Caption         =   "Product Text"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1335
               Index           =   0
               Left            =   6810
               MaskColor       =   &H0092BA5A&
               Style           =   1  'Graphical
               TabIndex        =   68
               Top             =   1770
               Width           =   2025
            End
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Template Description:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   3390
            TabIndex        =   17
            Top             =   300
            Width           =   1860
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reseller Block Fee:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   3615
            TabIndex        =   16
            Top             =   690
            Width           =   1620
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cost Price:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   9
            Left            =   6780
            TabIndex        =   15
            Top             =   690
            Width           =   930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendor Code:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   510
            TabIndex        =   14
            Top             =   300
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendor Sub Code:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   150
            TabIndex        =   13
            Top             =   690
            Width           =   1515
         End
      End
      Begin MSComctlLib.TabStrip ts 
         Height          =   8115
         Left            =   4140
         TabIndex        =   18
         Top             =   270
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   14314
         MultiRow        =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Primary Definitions"
               Object.ToolTipText     =   $"frmTemplateConfig.frx":0442
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Contracts"
               Object.ToolTipText     =   "Here is where you configure the different contracts available for this template."
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Categories"
               Object.ToolTipText     =   "Here is the categories of the template"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame frmts 
         BackColor       =   &H0084E8E8&
         Caption         =   "Templates Category Tree"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7665
         Index           =   2
         Left            =   4230
         TabIndex        =   63
         Top             =   660
         Width           =   9075
         Begin VB.CommandButton cmdMakeCategory 
            BackColor       =   &H0084E8E8&
            Caption         =   "Create &Category"
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
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   7080
            Width           =   3105
         End
         Begin MSComctlLib.TreeView tvCat 
            Height          =   6195
            Left            =   150
            TabIndex        =   65
            Top             =   780
            Width           =   8805
            _ExtentX        =   15531
            _ExtentY        =   10927
            _Version        =   393217
            Indentation     =   617
            LabelEdit       =   1
            LineStyle       =   1
            Sorted          =   -1  'True
            Style           =   7
            FullRowSelect   =   -1  'True
            ImageList       =   "ilSML"
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
         End
         Begin MSComctlLib.ImageList ilSML 
            Left            =   0
            Top             =   420
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   147
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":04D0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":0922
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":0D74
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":11C6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1618
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1A6A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1EBC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":230E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":2760
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":2BB2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":3004
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":3456
                  Key             =   ""
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":38A8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":3BC2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":4014
                  Key             =   ""
               EndProperty
               BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":4466
                  Key             =   ""
               EndProperty
               BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":48B8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":4D0A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":515C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":55AE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":5A00
                  Key             =   ""
               EndProperty
               BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":5E52
                  Key             =   ""
               EndProperty
               BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":62A4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":66F6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":6B48
                  Key             =   ""
               EndProperty
               BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":6F9A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":73EC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":783E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":7C90
                  Key             =   ""
               EndProperty
               BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":80E2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":8534
                  Key             =   ""
               EndProperty
               BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":8986
                  Key             =   ""
               EndProperty
               BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":8DD8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":922A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":967C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":9ACE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":9F20
                  Key             =   ""
               EndProperty
               BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":A372
                  Key             =   ""
               EndProperty
               BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":A7C4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":AC16
                  Key             =   ""
               EndProperty
               BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":B068
                  Key             =   ""
               EndProperty
               BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":B4BA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":B90C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":BD5E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":C1B0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":C602
                  Key             =   ""
               EndProperty
               BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":CA54
                  Key             =   ""
               EndProperty
               BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":CEA6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":D2F8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":D74A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":DB9C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":DFEE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":E440
                  Key             =   ""
               EndProperty
               BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":E892
                  Key             =   ""
               EndProperty
               BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":ECE4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":F136
                  Key             =   ""
               EndProperty
               BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":F588
                  Key             =   ""
               EndProperty
               BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":F9DA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":FE2C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1027E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":106D0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":10B22
                  Key             =   ""
               EndProperty
               BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":10F74
                  Key             =   ""
               EndProperty
               BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":113C6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":11818
                  Key             =   ""
               EndProperty
               BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":11C6A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":120BC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1250E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":12960
                  Key             =   ""
               EndProperty
               BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":12DB2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":13204
                  Key             =   ""
               EndProperty
               BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":13656
                  Key             =   ""
               EndProperty
               BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":13AA8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":141FA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1464C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":14A9E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":14EF0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":15342
                  Key             =   ""
               EndProperty
               BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":15794
                  Key             =   ""
               EndProperty
               BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":15AAE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":15DC8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":160E2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":163FC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":16716
                  Key             =   ""
               EndProperty
               BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":16A30
                  Key             =   ""
               EndProperty
               BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":16D4A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":17064
                  Key             =   ""
               EndProperty
               BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1737E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":17698
                  Key             =   ""
               EndProperty
               BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":179B2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":17CCC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":17FE6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":18300
                  Key             =   ""
               EndProperty
               BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1861A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":18934
                  Key             =   ""
               EndProperty
               BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":18C4E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":18F68
                  Key             =   ""
               EndProperty
               BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":19282
                  Key             =   ""
               EndProperty
               BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1959C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":198B6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":19BD0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":19EEA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1A204
                  Key             =   ""
               EndProperty
               BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1A51E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1A970
                  Key             =   ""
               EndProperty
               BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1ADC2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1B214
                  Key             =   ""
               EndProperty
               BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1B666
                  Key             =   ""
               EndProperty
               BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1BAB8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1BF0A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage111 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1C35C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage112 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1C7AE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage113 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1CC00
                  Key             =   ""
               EndProperty
               BeginProperty ListImage114 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1D052
                  Key             =   ""
               EndProperty
               BeginProperty ListImage115 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1D4A4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage116 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1D8F6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage117 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1DD48
                  Key             =   ""
               EndProperty
               BeginProperty ListImage118 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1E19A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage119 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1E5EC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage120 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1EA3E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage121 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1EE90
                  Key             =   ""
               EndProperty
               BeginProperty ListImage122 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1F2E2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage123 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1F734
                  Key             =   ""
               EndProperty
               BeginProperty ListImage124 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1FB86
                  Key             =   ""
               EndProperty
               BeginProperty ListImage125 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":1FFD8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage126 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":2042A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage127 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":2087C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage128 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":20CCE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage129 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":21120
                  Key             =   ""
               EndProperty
               BeginProperty ListImage130 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":21572
                  Key             =   ""
               EndProperty
               BeginProperty ListImage131 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":219C4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage132 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":21E16
                  Key             =   ""
               EndProperty
               BeginProperty ListImage133 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":22268
                  Key             =   ""
               EndProperty
               BeginProperty ListImage134 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":226BA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage135 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":22B0C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage136 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":22F5E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage137 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":233B0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage138 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":23802
                  Key             =   ""
               EndProperty
               BeginProperty ListImage139 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":23C54
                  Key             =   ""
               EndProperty
               BeginProperty ListImage140 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":240A6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage141 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":244F8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage142 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":2494A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage143 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":24D9C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage144 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":251EE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage145 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":25640
                  Key             =   ""
               EndProperty
               BeginProperty ListImage146 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":25A92
                  Key             =   ""
               EndProperty
               BeginProperty ListImage147 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTemplateConfig.frx":25EE4
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Label lblCat 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   210
            TabIndex        =   66
            Top             =   360
            Width           =   8745
         End
      End
      Begin VB.Frame frmts 
         BackColor       =   &H0039F2F2&
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
         Height          =   7695
         Index           =   1
         Left            =   4230
         TabIndex        =   19
         Top             =   660
         Visible         =   0   'False
         Width           =   9075
         Begin VB.Frame frameExtras 
            BackColor       =   &H0092BA5A&
            Caption         =   "Things that are included with this contract"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2475
            Index           =   1
            Left            =   90
            TabIndex        =   56
            Top             =   5100
            Width           =   8775
            Begin VB.CommandButton cmdPlansConfigure 
               Caption         =   "Plans Configuration"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Index           =   1
               Left            =   150
               TabIndex        =   60
               Top             =   1950
               Width           =   1665
            End
            Begin VB.ComboBox cmbAllPLans 
               BackColor       =   &H00A9F5AE&
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
               Index           =   1
               Left            =   1920
               Style           =   2  'Dropdown List
               TabIndex        =   59
               Top             =   240
               Width           =   5175
            End
            Begin VB.TextBox txtNumOf 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00A9F5AE&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   1
               Left            =   7170
               TabIndex        =   58
               Text            =   "1"
               Top             =   210
               Width           =   975
            End
            Begin VB.CommandButton cmdAddPlan 
               Caption         =   "+"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   8220
               TabIndex        =   57
               Top             =   240
               Width           =   405
            End
            Begin MSComctlLib.ListView lvPlans 
               Height          =   1665
               Index           =   1
               Left            =   1920
               TabIndex        =   61
               Top             =   660
               Width           =   6735
               _ExtentX        =   11880
               _ExtentY        =   2937
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               Checkboxes      =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               HotTracking     =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   11138478
               BorderStyle     =   1
               Appearance      =   1
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
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Plan Type to Include:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   1
               Left            =   150
               TabIndex        =   62
               Top             =   300
               Width           =   1560
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H007B7EF0&
            Caption         =   "Contract Details"
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
            Height          =   2325
            Left            =   60
            TabIndex        =   20
            Top             =   60
            Width           =   8925
            Begin VB.Frame Frame6 
               BackColor       =   &H00A3A3FE&
               Caption         =   "Measurement of Contract Interval"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1965
               Left            =   90
               TabIndex        =   21
               Top             =   240
               Width           =   3045
               Begin VB.OptionButton Option1 
                  BackColor       =   &H00A3A3FE&
                  Caption         =   "Seconds"
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
                  Index           =   0
                  Left            =   150
                  TabIndex        =   22
                  Tag             =   "s"
                  Top             =   300
                  Width           =   1245
               End
               Begin VB.OptionButton Option1 
                  BackColor       =   &H00A3A3FE&
                  Caption         =   "Minutes"
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
                  Left            =   150
                  TabIndex        =   23
                  Tag             =   "n"
                  Top             =   600
                  Width           =   1245
               End
               Begin VB.OptionButton Option1 
                  BackColor       =   &H00A3A3FE&
                  Caption         =   "Hours"
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
                  Index           =   2
                  Left            =   150
                  TabIndex        =   24
                  Tag             =   "h"
                  Top             =   900
                  Width           =   1245
               End
               Begin VB.OptionButton Option1 
                  BackColor       =   &H00A3A3FE&
                  Caption         =   "Days"
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
                  Index           =   3
                  Left            =   150
                  TabIndex        =   25
                  Tag             =   "d"
                  Top             =   1200
                  Width           =   1245
               End
               Begin VB.OptionButton Option1 
                  BackColor       =   &H00A3A3FE&
                  Caption         =   "Week days"
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
                  Index           =   4
                  Left            =   1440
                  TabIndex        =   26
                  Tag             =   "w"
                  Top             =   300
                  Width           =   1245
               End
               Begin VB.OptionButton Option1 
                  BackColor       =   &H00A3A3FE&
                  Caption         =   "Weeks"
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
                  Index           =   5
                  Left            =   1440
                  TabIndex        =   27
                  Tag             =   "ww"
                  Top             =   600
                  Value           =   -1  'True
                  Width           =   1245
               End
               Begin VB.OptionButton Option1 
                  BackColor       =   &H00A3A3FE&
                  Caption         =   "Months"
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
                  Index           =   6
                  Left            =   1440
                  TabIndex        =   28
                  Tag             =   "m"
                  Top             =   900
                  Width           =   1245
               End
               Begin VB.OptionButton Option1 
                  BackColor       =   &H00A3A3FE&
                  Caption         =   "Quarters"
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
                  Index           =   7
                  Left            =   1440
                  TabIndex        =   29
                  Tag             =   "q"
                  Top             =   1200
                  Width           =   1245
               End
               Begin VB.TextBox txtNoIntervals 
                  BackColor       =   &H00E0DFFF&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   1470
                  TabIndex        =   30
                  Tag             =   "NoPeriods"
                  Text            =   "52"
                  Top             =   1500
                  Width           =   1080
               End
               Begin MSComCtl2.UpDown UpDown1 
                  Height          =   360
                  Left            =   2550
                  TabIndex        =   31
                  Top             =   1500
                  Width           =   255
                  _ExtentX        =   450
                  _ExtentY        =   635
                  _Version        =   393216
                  Value           =   52
                  OrigLeft        =   2880
                  OrigTop         =   750
                  OrigRight       =   3120
                  OrigBottom      =   1110
                  Max             =   999999999
                  Wrap            =   -1  'True
                  Enabled         =   -1  'True
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Contract Length:"
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
                  Index           =   1
                  Left            =   150
                  TabIndex        =   48
                  Top             =   1590
                  Width           =   1185
               End
            End
            Begin VB.Frame Frame7 
               BackColor       =   &H00A3A3FE&
               Caption         =   "Contract Description"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1965
               Left            =   3180
               TabIndex        =   38
               Top             =   240
               Width           =   5655
               Begin VB.TextBox txtField 
                  BackColor       =   &H00E0DFFF&
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
                  Left            =   1260
                  TabIndex        =   32
                  Tag             =   "Description"
                  Top             =   240
                  Width           =   4275
               End
               Begin VB.TextBox txtFee 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0DFFF&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Index           =   6
                  Left            =   4020
                  TabIndex        =   34
                  Tag             =   "Termination"
                  Top             =   600
                  Width           =   1545
               End
               Begin VB.TextBox txtFee 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0DFFF&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Index           =   7
                  Left            =   3990
                  TabIndex        =   36
                  Tag             =   "PeriodFee"
                  Top             =   1050
                  Width           =   1545
               End
               Begin VB.TextBox txtFee 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0DFFF&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Index           =   8
                  Left            =   1260
                  TabIndex        =   33
                  Tag             =   "JoiningFee"
                  Top             =   600
                  Width           =   1395
               End
               Begin VB.TextBox txtFee 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0DFFF&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Index           =   9
                  Left            =   1260
                  TabIndex        =   35
                  Tag             =   "FeePerBlock"
                  Top             =   1050
                  Width           =   1395
               End
               Begin VB.TextBox txtFee 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0DFFF&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Index           =   10
                  Left            =   1260
                  TabIndex        =   37
                  Tag             =   "FeePerHour"
                  Top             =   1500
                  Width           =   1395
               End
               Begin VB.CommandButton cmdContract 
                  BackColor       =   &H00A3A3FE&
                  Caption         =   "Sa&ve"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Index           =   0
                  Left            =   4050
                  Style           =   1  'Graphical
                  TabIndex        =   40
                  Top             =   1530
                  Width           =   1515
               End
               Begin VB.CommandButton cmdContract 
                  BackColor       =   &H00A3A3FE&
                  Caption         =   "&New"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Index           =   1
                  Left            =   2730
                  Style           =   1  'Graphical
                  TabIndex        =   39
                  Top             =   1530
                  Width           =   1275
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Description:"
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
                  Left            =   330
                  TabIndex        =   47
                  Top             =   300
                  Width           =   840
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Termination Fee:"
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
                  Index           =   2
                  Left            =   2730
                  TabIndex        =   46
                  Top             =   690
                  Width           =   1185
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Billing Cycle Fee:"
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
                  Index           =   3
                  Left            =   2730
                  TabIndex        =   45
                  Top             =   1140
                  Width           =   1200
               End
               Begin VB.Label Label3 
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
                  Height          =   195
                  Index           =   4
                  Left            =   330
                  TabIndex        =   44
                  Top             =   720
                  Width           =   855
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Extra Per MB:"
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
                  Index           =   5
                  Left            =   210
                  TabIndex        =   43
                  Top             =   1140
                  Width           =   975
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Extra Per Hour:"
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
                  Index           =   6
                  Left            =   120
                  TabIndex        =   42
                  Top             =   1560
                  Width           =   1080
               End
            End
         End
         Begin MSComctlLib.ListView lvContracts 
            Height          =   2595
            Left            =   60
            TabIndex        =   41
            Tag             =   "0"
            Top             =   2430
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   4577
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
   End
   Begin VB.Menu mnuHide 
      Caption         =   "mnuHidden"
      Visible         =   0   'False
      Begin VB.Menu lvPlans_lvContracts 
         Caption         =   "ContractsListview"
         Begin VB.Menu lvContracts_Delete 
            Caption         =   "Delete"
         End
         Begin VB.Menu lvContracts_Import 
            Caption         =   "Import Contracts from other service"
         End
      End
   End
End
Attribute VB_Name = "frmTemplateConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const iLeft = 60
Const iTop = 1020
Const iWidth = 8925
Const iHeight = 3945

Dim rsPlanTemp As ADODB.Recordset
Dim aConn As ADODB.Connection

Private Sub chkLimit_Click(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "chkLimit_Click"
    Const ContainerName = "frmTemplateConfig"
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

Private Sub cmbVendors_Click(Index As Integer)


    Dim Xnt As Integer
    
    For Xnt = cmbVendors.LBound To cmbVendors.UBound
    
        If Xnt <> Index Then
            cmbVendors(Xnt).ListIndex = cmbVendors(Index).ListIndex
        End If
    
    Next
    
End Sub

Private Sub cmdContract_Click(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdContract_Click"
    Const ContainerName = "frmTemplateConfig"
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
    Case 1
        txtField(0).Text = txtNoIntervals.Text
        Dim kx As Byte
        For kx = 0 To 7
            If Option1(kx).Value = True Then
                txtField(0).Text = txtField(0).Text + " " + Option1(kx).Caption
            End If
        Next
        
        txtFee(6) = "0.00"
        txtFee(7) = txtFee(0)
        txtFee(8) = "0.00"
        txtFee(9) = txtFee(1)
        txtFee(10) = txtFee(2)
        lvPlans(1).ListItems.Clear
        
        lvContracts.Tag = 0
    Case 0
    
        Dim SQL As String
        Dim itmX As ListItem
        
        Select Case lvContracts.Tag
        Case 0
        
            On Error Resume Next
            Do
                Err.Clear
                lvContracts.Tag = MySQL.GetTMPRecID("contracttemplates", ADOConn)
                Call MySQL.Execute(ADOConn, "INSERT into contracttemplates (RecID, ptRecID,VirtualID) VALUES ('" & lvContracts.Tag & "','" & txtDesc.Tag & "','" & Login.lVirtualID & "')")
            Loop Until Err.Number = 0
            
            Set itmX = lvContracts.ListItems.Add(, "r" & lvContracts.Tag, txtField(0).Text)
        Case Else
        
            Set itmX = lvContracts.SelectedItem
            
        End Select
    
        SQL = "Update contracttemplates set "
        SQL = SQL + "`" & txtField(0).Tag & "` = '" & MySQL.ESC(txtField(0).Text) & "', "
        SQL = SQL + "`" & txtNoIntervals.Tag & "` = '" & txtNoIntervals.Text & "', "
        SQL = SQL + "`" & txtFee(6).Tag & "` = '" & txtFee(6).Text & "', "
        SQL = SQL + "`" & txtFee(7).Tag & "` = '" & txtFee(7).Text & "', "
        SQL = SQL + "`" & txtFee(8).Tag & "` = '" & txtFee(8).Text & "', "
        SQL = SQL + "`" & txtFee(9).Tag & "` = '" & txtFee(9).Text & "', "
        SQL = SQL + "`" & txtFee(10).Tag & "` = '" & txtFee(10).Text & "', "
        For bx = 0 To 7
            If Option1(bx).Value = True Then SQL = SQL + "`TypePeriods` = '" & Option1(bx).Tag & "' "
        Next bx
        SQL = SQL + " where RecID = " & lvContracts.Tag
        
        MySQL.Execute ADOConn, SQL
    
        itmX.Text = txtField(0).Text
        itmX.SubItems(1) = txtNoIntervals.Text
        
        For bx = 0 To 7
            If Option1(bx).Value = True Then itmX.SubItems(2) = Option1(bx).Tag
        Next bx
        
        itmX.SubItems(3) = txtFee(6)
        itmX.SubItems(5) = txtFee(7)
        itmX.SubItems(4) = txtFee(8)
        itmX.SubItems(6) = txtFee(9)
        itmX.SubItems(7) = txtFee(10)
        itmX.SubItems(8) = lvContracts.Tag
        
        SaveFlags lvContracts.Tag, 1
        
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

Private Sub cmdMakeCategory_Click()

    Dim fCat As New frmCreateCat
    
    Set fCat.tv = tvCat
    fCat.Show 1
    
    If fCat.RecID <> 0 Then
        Me.LoadCategories tvCat
    End If
    
End Sub


Private Sub Command1_Click(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Command1_Click"
    Const ContainerName = "frmTemplateConfig"
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


    Dim oPT As New frmProductText
    oPT.bUnlock = False
    oPT.Text = Command1(0).Tag
    oPT.Show 1
    
    MySQL.Execute aConn, "Update plantemplates set ProductText = '" & MySQL.ESC(oPT.Text) & "' where RecID = " & txtDesc.Tag
    Command1(0).Tag = oPT.Text
    
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

Private Sub cmdAdd_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdAdd_Click"
    Const ContainerName = "frmTemplateConfig"
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


    Command1(0).Enabled = False
    txtDesc = ""
    txtDesc.Tag = 0
    txtFee(0) = ""
    txtFee(1) = ""
    txtFee(2) = ""
    txtFee(3) = ""
    txtFee(4) = ""
    txtFee(5) = ""
    lvPlans(0).ListItems.Clear
    
    txtDesc.SetFocus
    
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
    Const ContainerName = "frmTemplateConfig"
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
Public Function LoadFlags(lRecID As Variant, Optional bContract As Boolean)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "LoadFlags"
    Const ContainerName = "frmTemplateConfig"
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
    
    Select Case bContract
    Case False
        bResult = MySQL.OpenTable(aConn, rsload, , "select plantypes.Description, flags_tempextras.* from plantypes, flags_tempextras Where flags_tempextras.PlanType = plantypes.RecID AND ptRecID = " & lRecID)
        lvPlans(0).ListItems.Clear
        
        If rsload.RecordCount > 0 Then
            rsload.MoveFirst
            While Not rsload.EOF And Err.Number = 0
                Set itmX = lvPlans(0).ListItems.Add(, "r" & rsload!RecID, rsload!Description)
                itmX.Tag = rsload!PlanType
                itmX.SubItems(1) = rsload!NumberOf
                itmX.Checked = rsload!Checked
                rsload.MoveNext
            Wend
        End If
    
    Case True
        
        bResult = MySQL.OpenTable(aConn, rsload, , "select plantypes.Description, flags_tempextras.* from plantypes, flags_tempextras Where flags_tempextras.PlanType = plantypes.RecID AND ContractID = " & lRecID)
    
        lvPlans(1).ListItems.Clear
        
        If rsload.RecordCount > 0 Then
            rsload.MoveFirst
            While Not rsload.EOF And Err.Number = 0
                Set itmX = lvPlans(1).ListItems.Add(, "r" & rsload!RecID, rsload!Description)
                itmX.Tag = rsload!PlanType
                itmX.SubItems(1) = rsload!NumberOf
                itmX.Checked = rsload!Checked
                rsload.MoveNext
            Wend
        End If
        
    End Select
    


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

Public Function SaveFlags(lRecID As Variant, Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "SaveFlags"
    Const ContainerName = "frmTemplateConfig"
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
    
        If lvPlans(Index).ListItems.Count > 0 Then
        
            For sa = 1 To lvPlans(Index).ListItems.Count
                Set itmX = lvPlans(Index).ListItems(sa)
                If itmX.Key = "" Then
                    
                    Select Case Index
                    Case 0
                        Call MySQL.Execute(ADOConn, "insert into flags_tempextras (PlanType, NumberOf, Checked, ptRecID) VALUES ('" & itmX.Tag & "','" & itmX.SubItems(1) & "','" & IIf(itmX.Checked = True, "-1", "0") & "','" & lRecID & "')")
                    Case 1
                        Call MySQL.Execute(ADOConn, "insert into flags_tempextras (PlanType, NumberOf, Checked, ContractID) VALUES ('" & itmX.Tag & "','" & itmX.SubItems(1) & "','" & IIf(itmX.Checked = True, "-1", "0") & "','" & lvContracts.Tag & "')")
                    End Select
                                    
                    itmX.Key = "saved" & Rnd * 6.55646546546465E+17
                    
                ElseIf Left(itmX.Key, 1) = "e" Then
                
                    Call MySQL.OpenTable(aConn, rsSave, , "select * from flags_tempextras where RecID = " & Mid(itmX.Key, 2) & " Limit 1")
                    
                    rsSave!ptRecID = lRecID
                    rsSave!PlanType = itmX.Tag
                    rsSave!NumberOf = Val(itmX.SubItems(1))
                    rsSave!Checked = itmX.Checked
                    Select Case Index
                    Case 0
                        rsSave!ptRecID = lRecID
                    Case 1
                        rsSave!ContractID = lvContracts.Tag
                    End Select
                    
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

Private Sub cmdPlansConfigure_Click(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdPlansConfigure_Click"
    Const ContainerName = "frmTemplateConfig"
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


    frmAccountTypes.Show 1
    
    
    cmbAllPLans(0).Clear
    cmbAllPLans(1).Clear
    bResult = MySQL.OpenTable(aConn, rsPlanTemp, , "select * from plantypes")
    If rsPlanTemp.RecordCount > 0 Then
        While Not rsPlanTemp.EOF And Err.Number = 0
            cmbAllPLans(0).AddItem rsPlanTemp!Description
            cmbAllPLans(0).ItemData(cmbAllPLans(0).ListCount - 1) = rsPlanTemp!RecID
            cmbAllPLans(1).AddItem rsPlanTemp!Description
            cmbAllPLans(1).ItemData(cmbAllPLans(0).ListCount - 1) = rsPlanTemp!RecID
            rsPlanTemp.MoveNext
        Wend
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

Private Sub UpdateBody(lRecID As Double)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "UpdateBody"
    Const ContainerName = "frmTemplateConfig"
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


    Select Case True
    Case picSet(0).Visible
        
        

    Case picSet(1).Visible
    
    
    Case picSet(2).Visible
    
    
    Case picSet(3).Visible
    
        

    End Select


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
    Const ContainerName = "frmTemplateConfig"
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


    Dim bNumeric As Boolean
    'Dim rsload As ADODB.Recordset
    Dim SQL As String
    Dim bResult As Boolean
    Dim rsload As ADODB.Recordset
    Dim rsServices As ADODB.Recordset
    Dim Xnt As Integer
    
    On Error Resume Next
    Dim ltmpRecID As Variant
   
    cmdSave.Enabled = False
    lvAccounts.Enabled = False
    
    Dim bx As Byte
    
    
    bResult = MySQL.OpenTable(aConn, rsServices, , "select BillImmediately from servicetypes where RecID = " & Mid(tvServiceTypes.SelectedItem.Key, 2))
    
    Select Case tvServiceTypes.SelectedItem.Tag
    
    Case "FTP", "POP3", "WWW", "DESIGN", "DOMAIN", "GATEWAY", "TRAINING", "SALES", "CONSULT", "COLO", "ALIAS", "HOST"
        Select Case txtDesc.Tag
        Case 0, ""
            
            On Error Resume Next
            Do
                Err.Clear
                ltmpRecID = MySQL.GetTMPRecID("plantemplates", ADOConn)
                Call MySQL.Execute(ADOConn, "INSERT INTO plantemplates (RecID,SysopID) VALUES ('" & ltmpRecID & "','" & Login.lSysopID & "')")
            Loop Until Err.Number = 0
            
            txtDesc.Tag = ltmpRecID
        
            MySQL.Execute aConn, "Update plantemplates set ServiceID = '" & Mid(tvServiceTypes.SelectedItem.Key, 2) & "' where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set Description = '" & MySQL.ESC(txtDesc) & "' where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set PeriodFee = '" & Val(txtFee(0)) & "' where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set CostPrice = '" & Val(txtFee(3)) & "' where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set BillImmediately =  '" & rsServices!BillImmediately & "' where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set MBCostPrice = '" & Val(txtFee(4)) & "' where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set PeriodCostPrice = '" & Val(txtFee(5)) & "' where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set MBPerPeriod = -1 where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set MBBlockSize = 0 where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set FeePerBlock = 0 where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set MbQuota = '" & Val(txtQuota(0).Text) & "' where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set HoursPerPeriod = -1 where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set ExtraPerHour = 0 where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set SessionTimeout = '" & MySQL.ESC(txtSessionTimeout.Text) & "' where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set IdleTimeout = '" & MySQL.ESC(txtIdleTimeout.Text) & "' where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set Hidden = '" & chkHidden.Value & "' where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set Description = '" & MySQL.ESC(txtDesc) & "' where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set VendorID = '" & cmbVendors(0).ItemData(cmbVendors(0).ListIndex) & "' where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set VendorPartID = '" & MySQL.ESC(txtPartID.Text) & "' where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set SubPartID = '" & MySQL.ESC(txtSubPartID.Text) & "' where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set Description = '" & MySQL.ESC(txtDesc) & "' where RecID = " & ltmpRecID
            MySQL.Execute aConn, "Update plantemplates set RadiusID = '" & cmdRadius.ItemData(cmdRadius.ListIndex) & "' where RecID = " & ltmpRecID
            MySQL.Execute aConn, "Update plantemplates set VirtualID = '" & Login.lVirtualID & "' where RecID = " & txtDesc.Tag
            If tvCat.SelectedItem Is Nothing Then
            Else
                MySQL.Execute aConn, "Update plantemplates set CategoryID = '" & Mid(tvCat.SelectedItem.Key, 2) & "' where RecID = " & txtDesc.Tag
            End If
            MySQL.Execute aConn, "Update plantypes set CostPrice = '" & Val(txtFee(0)) & "' where TemplateID = " & txtDesc.Tag
            
            Call MySQL.OpenTable(ADOConn, rsPlanTemp, , "select * from plantemplates where RecID = '" & txtDesc.Tag & "'")
            
            MySQL.fillLV ADOConn, rsPlanTemp, lvAccounts, True
            
        Case Else
        
            'rsPlanTemp.Filter = "RecID = " &
        
            'If rsPlanTemp.RecordCount > 0 Then
            
                MySQL.Execute aConn, "Update plantemplates set Description = '" & MySQL.ESC(txtDesc) & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set PeriodFee = '" & Val(txtFee(0)) & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set CostPrice = '" & Val(txtFee(3)) & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set BillImmediately =  '" & rsServices!BillImmediately & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set MBCostPrice = '" & Val(txtFee(4)) & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set PeriodCostPrice = '" & Val(txtFee(5)) & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set MBPerPeriod = -1 where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set MBBlockSize = 0 where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set FeePerBlock = 0 where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set MbQuota = '" & Val(txtQuota(0).Text) & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set HoursPerPeriod = -1 where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set ExtraPerHour = 0 where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set SessionTimeout = '" & MySQL.ESC(txtSessionTimeout.Text) & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set IdleTimeout = '" & MySQL.ESC(txtIdleTimeout.Text) & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set Hidden = '" & chkHidden.Value & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set Description = '" & MySQL.ESC(txtDesc) & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set VendorID = '" & cmbVendors(0).ItemData(cmbVendors(0).ListIndex) & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set VendorPartID = '" & MySQL.ESC(txtPartID.Text) & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set SubPartID = '" & MySQL.ESC(txtSubPartID.Text) & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set ServiceID = '" & Mid(tvServiceTypes.SelectedItem.Key, 2) & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set RadiusID = '" & cmdRadius.ItemData(cmdRadius.ListIndex) & "' where RecID = " & ltmpRecID
                MySQL.Execute aConn, "Update plantemplates set VirtualID = '" & Login.lVirtualID & "' where RecID = " & txtDesc.Tag
                
                
                UpdateBody txtDesc.Tag
                
                If tvCat.SelectedItem Is Nothing Then
                Else
                    MySQL.Execute aConn, "Update plantemplates set CategoryID = '" & Mid(tvCat.SelectedItem.Key, 2) & "' where RecID = " & txtDesc.Tag
                End If
                MySQL.Execute aConn, "Update plantypes set CostPrice = '" & Val(txtFee(0)) & "' where TemplateID = " & txtDesc.Tag
                
                ltmpRecID = txtDesc.Tag
                
                Call MySQL.OpenTable(ADOConn, rsPlanTemp, , "select * from plantemplates where RecID = '" & txtDesc.Tag & "'")
                MySQL.fillLV ADOConn, rsPlanTemp, lvAccounts, True, lvAccounts.ListItems("r" & ltmpRecID)
        End Select
    Case "DIALUP", "ADSL", "SHDSL"
            
        If cmdRadius.ListIndex = -1 Then
            MsgBox "You must set the radius pool that is being used!"
            Exit Sub
        End If

        Select Case txtDesc.Tag
        Case 0, ""
            
            On Error Resume Next
            Do
                Err.Clear
                ltmpRecID = MySQL.GetTMPRecID("plantemplates", ADOConn)
                Call MySQL.Execute(ADOConn, "INSERT INTO plantemplates (RecID, SysopID) VALUES ('" & ltmpRecID & "','" & Login.lSysopID & "')")
            Loop Until Err.Number = 0
            
            txtDesc.Tag = ltmpRecID
        
            MySQL.Execute aConn, "Update plantemplates set Description = '" & MySQL.ESC(txtDesc) & "' where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set PeriodFee = '" & Val(txtFee(0)) & "' where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set CostPrice = '" & Val(txtFee(3)) & "' where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set BillImmediately =  '" & rsServices!BillImmediately & "' where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set MBCostPrice = '" & Val(txtFee(4)) & "' where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set PeriodCostPrice = '" & Val(txtFee(5)) & "' where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set ServiceID = '" & Mid(tvServiceTypes.SelectedItem.Key, 2) & "' where RecID = " & txtDesc.Tag
            
            UpdateBody txtDesc.Tag
            
            If tvCat.SelectedItem Is Nothing Then
            Else
                MySQL.Execute aConn, "Update plantemplates set CategoryID = '" & Mid(tvCat.SelectedItem.Key, 2) & "' where RecID = " & txtDesc.Tag
            End If
            MySQL.Execute aConn, "Update plantypes set CostPrice = '" & Val(txtFee(0)) & "' where TemplateID = " & txtDesc.Tag
            
            If chkLimit(1).Value = 1 Then
            
                MySQL.Execute aConn, "Update plantemplates set MBPerPeriod = '" & Val(txtMB(1)) & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set MBBlockSize = '" & Val(txtMB(0)) & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set FeePerBlock = '" & Val(txtFee(1)) & "' where RecID = " & txtDesc.Tag
            
            Else
            
                MySQL.Execute aConn, "Update plantemplates set MBPerPeriod = -1 where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set MBBlockSize = 0 where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set FeePerBlock = 0 where RecID = " & txtDesc.Tag
            
            End If
            
            MySQL.Execute aConn, "Update plantemplates set MbQuota = '" & Val(txtQuota(0).Text) & "' where RecID = " & txtDesc.Tag
            
            If chkLimit(0).Value = 1 Then
            
                MySQL.Execute aConn, "Update plantemplates set HoursPerPeriod = '" & txtHours & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set ExtraPerHour = '" & Val(txtFee(2)) & "' where RecID = " & txtDesc.Tag
            
            Else
            
                MySQL.Execute aConn, "Update plantemplates set HoursPerPeriod = -1 where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set ExtraPerHour = 0 where RecID = " & txtDesc.Tag
                
            End If

            MySQL.Execute aConn, "Update plantemplates set SessionTimeout = '" & MySQL.ESC(txtSessionTimeout.Text) & "' where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set IdleTimeout = '" & MySQL.ESC(txtIdleTimeout.Text) & "' where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set Hidden = '" & chkHidden.Value & "' where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set Description = '" & MySQL.ESC(txtDesc) & "' where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set VendorID = '" & cmbVendors(0).ItemData(cmbVendors(0).ListIndex) & "' where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set VendorPartID = '" & MySQL.ESC(txtPartID.Text) & "' where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set SubPartID = '" & MySQL.ESC(txtSubPartID.Text) & "' where RecID = " & txtDesc.Tag
            MySQL.Execute aConn, "Update plantemplates set Description = '" & MySQL.ESC(txtDesc) & "' where RecID = " & ltmpRecID
            MySQL.Execute aConn, "Update plantemplates set RadiusID = '" & cmdRadius.ItemData(cmdRadius.ListIndex) & "' where RecID = " & ltmpRecID
            MySQL.Execute aConn, "Update plantemplates set VirtualID = '" & Login.lVirtualID & "' where RecID = " & txtDesc.Tag
            Call MySQL.OpenTable(ADOConn, rsPlanTemp, , "select * from plantemplates where RecID = '" & txtDesc.Tag & "'")
            
            
            MySQL.fillLV ADOConn, rsPlanTemp, lvAccounts, True
            
        Case Else
        
            If cmdRadius.ListIndex = -1 Then
                MsgBox "You must set the radius pool that is being used!"
                Exit Sub
            End If
            
                MySQL.Execute aConn, "Update plantypes set CostPrice = '" & Val(txtFee(0)) & "' where TemplateID = " & txtDesc.Tag
                
                MySQL.Execute aConn, "Update plantemplates set ServiceID = '" & Mid(tvServiceTypes.SelectedItem.Key, 2) & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set Description = '" & MySQL.ESC(txtDesc) & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set PeriodFee = '" & Val(txtFee(0)) & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set CostPrice = '" & Val(txtFee(3)) & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set BillImmediately =  '" & rsServices!BillImmediately & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set MBCostPrice = '" & Val(txtFee(4)) & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set PeriodCostPrice = '" & Val(txtFee(5)) & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set MBPerPeriod = -1 where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set MBBlockSize = 0 where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set FeePerBlock = 0 where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set MbQuota = '" & Val(txtQuota(0).Text) & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set HoursPerPeriod = -1 where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set ExtraPerHour = 0 where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set SessionTimeout = '" & MySQL.ESC(txtSessionTimeout.Text) & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set IdleTimeout = '" & MySQL.ESC(txtIdleTimeout.Text) & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set Hidden = '" & chkHidden.Value & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set Description = '" & MySQL.ESC(txtDesc) & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set VendorID = '" & cmbVendors(0).ItemData(cmbVendors(0).ListIndex) & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set VendorPartID = '" & MySQL.ESC(txtPartID.Text) & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set SubPartID = '" & MySQL.ESC(txtSubPartID.Text) & "' where RecID = " & txtDesc.Tag
                MySQL.Execute aConn, "Update plantemplates set Description = '" & MySQL.ESC(txtDesc) & "' where RecID = " & ltmpRecID
                MySQL.Execute aConn, "Update plantemplates set RadiusID = '" & cmdRadius.ItemData(cmdRadius.ListIndex) & "' where RecID = " & ltmpRecID
                    
                UpdateBody txtDesc.Tag
                
                If chkLimit(1).Value = 1 Then
                    MySQL.Execute aConn, "Update plantemplates set MBPerPeriod = '" & Val(txtMB(1)) & "' where RecID = " & txtDesc.Tag
                    MySQL.Execute aConn, "Update plantemplates set MBBlockSize = '" & Val(txtMB(0)) & "' where RecID = " & txtDesc.Tag
                    MySQL.Execute aConn, "Update plantemplates set FeePerBlock = '" & Val(txtFee(1)) & "' where RecID = " & txtDesc.Tag
                Else
                    MySQL.Execute aConn, "Update plantemplates set MBPerPeriod = -1 where RecID = " & txtDesc.Tag
                    MySQL.Execute aConn, "Update plantemplates set MBBlockSize = 0 where RecID = " & txtDesc.Tag
                    MySQL.Execute aConn, "Update plantemplates set FeePerBlock = 0 where RecID = " & txtDesc.Tag
                End If
                
                If chkLimit(0).Value = 1 Then
                    MySQL.Execute aConn, "Update plantemplates set HoursPerPeriod = '" & txtHours & "' where RecID = " & txtDesc.Tag
                    MySQL.Execute aConn, "Update plantemplates set ExtraPerHour = '" & Val(txtFee(2)) & "' where RecID = " & txtDesc.Tag
                Else
                    MySQL.Execute aConn, "Update plantemplates set HoursPerPeriod = -1 where RecID = " & txtDesc.Tag
                    MySQL.Execute aConn, "Update plantemplates set ExtraPerHour = 0 where RecID = " & txtDesc.Tag
                End If
                If tvCat.SelectedItem Is Nothing Then
                Else
                    MySQL.Execute aConn, "Update plantemplates set CategoryID = '" & Mid(tvCat.SelectedItem.Key, 2) & "' where RecID = " & txtDesc.Tag
                End If
                ltmpRecID = txtDesc.Tag
            
                MySQL.Execute aConn, "Update plantemplates set Description = '" & MySQL.ESC(txtDesc) & "' where RecID = " & ltmpRecID
                MySQL.Execute aConn, "Update plantemplates set VirtualID = '" & Login.lVirtualID & "' where RecID = " & txtDesc.Tag
                
                Call MySQL.OpenTable(ADOConn, rsPlanTemp, , "select * from plantemplates where RecID = '" & txtDesc.Tag & "'")
                
                MySQL.fillLV ADOConn, rsPlanTemp, lvAccounts, True, lvAccounts.ListItems("r" & ltmpRecID)
            
        End Select
    End Select
    
    Select Case True
    Case picSet(3).Visible
        
        Dim csql As String
                
        Dim X As Integer
        
        For X = chkCont.LBound To chkCont.UBound
            csql = csql + "`c" & MySQL.ReplaceString(chkCont(X).Caption, " ", "_") & "` = '" & chkCont(X).Value & "', "
        Next X
        
        csql = Left(csql, Len(csql) - 2)
        
        MySQL.Execute aConn, "Update plantemplates set " + csql + " where RecID = " & ltmpRecID
        MySQL.Execute aConn, "Update plantemplates set unitperpack = '" & Val(txtUnits) & "' where RecID = " & ltmpRecID
        MySQL.Execute aConn, "Update plantemplates set height = " & Val(txtDem(0).Text) & "' where RecID = " & ltmpRecID
        MySQL.Execute aConn, "Update plantemplates set length = " & Val(txtDem(1).Text) & "' where RecID = " & ltmpRecID
        MySQL.Execute aConn, "Update plantemplates set depth = " & Val(txtDem(2).Text) & "' where RecID = " & ltmpRecID
        MySQL.Execute aConn, "Update plantemplates set weight = " & Val(txtDem(3).Text) & "' where RecID = " & ltmpRecID
    
    
        For X = optPack.LBound To optPack.UBound
            Select Case X
            Case 0
                MySQL.Execute aConn, "Update plantemplates set packaging = " & MySQL.ESC(cmbPack.Text) & "' where RecID = " & ltmpRecID
            Case Else
                MySQL.Execute aConn, "Update plantemplates set packaging = " & MySQL.ESC(optPack(X).Caption) & "' where RecID = " & ltmpRecID
            End Select
        Next
        
    Case picSet(2).Visible
    
        MySQL.Execute aConn, "Update plantemplates set location = '" & Val(txtLoc.Text) & "' where RecID = " & ltmpRecID
        MySQL.Execute aConn, "Update plantemplates set chargebyrate = '" & Val(chkRates.Value) & "' where RecID = " & ltmpRecID
        
        If chkRates.Value = 1 Or chkRates.Value = 2 Then
            Dim Y As Integer
                
            Select Case True
            Case optRates(0).Value
                MySQL.Execute aConn, "Update plantemplates set ratetype = 'n' where RecID = " & ltmpRecID
                MySQL.Execute aConn, "Update plantemplates set rateinterval = '" & 30 & "' where RecID = " & ltmpRecID
            Case optRates(1).Value
                MySQL.Execute aConn, "Update plantemplates set ratetype = 'n' where RecID = " & ltmpRecID
                MySQL.Execute aConn, "Update plantemplates set rateinterval = '" & 60 & "' where RecID = " & ltmpRecID
            Case optRates(2).Value
                MySQL.Execute aConn, "Update plantemplates set ratetype = 'd' where RecID = " & ltmpRecID
                MySQL.Execute aConn, "Update plantemplates set rateinterval = '1' where RecID = " & ltmpRecID
            Case optRates(3).Value
                MySQL.Execute aConn, "Update plantemplates set ratetype = 'ww' where RecID = " & ltmpRecID
                MySQL.Execute aConn, "Update plantemplates set rateinterval = '1' where RecID = " & ltmpRecID
            Case optRates(4).Value
                MySQL.Execute aConn, "Update plantemplates set ratetype = 'm' where RecID = " & ltmpRecID
                MySQL.Execute aConn, "Update plantemplates set rateinterval = '1' where RecID = " & ltmpRecID
            Case optRates(5).Value
                MySQL.Execute aConn, "Update plantemplates set ratetype = 'q' where RecID = " & ltmpRecID
                MySQL.Execute aConn, "Update plantemplates set rateinterval = '1' where RecID = " & ltmpRecID
            Case optRates(6).Value
                MySQL.Execute aConn, "Update plantemplates set ratetype = 'y' where RecID = " & ltmpRecID
                MySQL.Execute aConn, "Update plantemplates set rateinterval = '1' where RecID = " & ltmpRecID
            End Select
        End If
        
    Case Else
    
    End Select
        
    SaveFlags ltmpRecID, 0
    
    'Call tvservicetypes_NodeClick(tvServiceTypes.selectedItem)
    
    cmdSave.Enabled = True
    For Xnt = Command1.LBound To Command1.UBound
        Command1(Xnt).Enabled = True
    Next
    rsPlanTemp.Filter = ""
    lvAccounts.Enabled = True
    frmAgent.oChar.StopAll
    
            
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

Private Sub cmdAddPlan_Click(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdAddPlan_Click"
    Const ContainerName = "frmTemplateConfig"
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
        
    If cmbAllPLans(Index).ListIndex > -1 Then
    
        Set itmX = lvPlans(Index).ListItems.Add(, , cmbAllPLans(Index).Text)
        itmX.SubItems(1) = txtNumOf(Index).Text
        itmX.Tag = cmbAllPLans(Index).ItemData(cmbAllPLans(Index).ListIndex)
        itmX.Checked = True
        
    
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

Private Sub Command3_Click(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Command3_Click"
    Const ContainerName = "frmTemplateConfig"
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
Sub LoadCategories(tv As TreeView)

    tv.NodeS.Clear
    lblCat.Caption = ""
    
    Dim db As Double
    Dim dba As Double
    
    On Error Resume Next
    
    While dba < GUI.mapCategory.Count And Err.Number = 0
    
        If db > GUI.mapCategory.Count Then db = 1 Else db = db + 1
        
        Select Case GUI.mapCategory(db).SubRecID
        Case 0
            tvCat.NodeS.Add , , "r" & GUI.mapCategory(db).RecID, GUI.mapCategory(db).Description, GUI.mapCategory(db).Icon, GUI.mapCategory(db).Icon
        
        Case Else
            tvCat.NodeS.Add "r" & GUI.mapCategory(db).SubRecID, tvwChild, "r" & GUI.mapCategory(db).RecID, GUI.mapCategory(db).Description, GUI.mapCategory(db).Icon, GUI.mapCategory(db).Icon
        
        End Select
        
        Select Case Err.Number
        Case 0
            dba = dba + 1
        Case Else
            Err.Clear
        End Select
    Wend
    
End Sub
Private Sub Form_Load()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Load"
    Const ContainerName = "frmTemplateConfig"
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


    Select Case Login.bTestBench
    Case False
        If MySQL.Connection(, , , , aConn) = False Then
            MsgBox "Unable to Connect to MySQL Server, Please check your internet connection and attempt to restart the program.", vbCritical, "MySQL Server Not Found"
            End
        End If
    Case True
        
        sServer = "localhost"
        sUID = "pa2004"
        sPWD = "p0st41"
        If MySQL.Connection(, sServer, sUID, sPWD, aConn) = False Then
            MsgBox "Unable to Connect to MySQL Test Bench Server, Please check your LAN connection and attempt to restart the program.", vbCritical, "MySQL Test Bench Not Found"
            End
        End If
    
    End Select
    
    Call Me.LoadCategories(tvCat)
    
        
    PopulateList
    Call GUI.LoadColWidths(lvAccounts, Me)
    Call GUI.LoadColWidths(lvPlans(0), Me)
    Call GUI.LoadColWidths(lvPlans(1), Me)
    Call GUI.LoadColWidths(lvContracts, Me)
    
   
        For ix = frmts.LBound To frmts.UBound
            If ix <> ts.SelectedItem.Index - 1 Then
                frmts(ix).Visible = False
            Else
                frmts(ix).Move ts.ClientLeft, ts.ClientTop, ts.ClientWidth, ts.ClientHeight
                frmts(ix).ZOrder 0
                frmts(ix).Visible = True
            End If
        Next ix
        
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
    Const ContainerName = "frmTemplateConfig"
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
    
    
    If MySQL.OpenTable(aConn, rsload, , MySQL.virtualisp("select distinct vendors.vName ,vendors.RecID from vendors", "vendors", False, IIf(Login.lLevel > 75, True, False))) = True Then
        If rsload.State = adStateOpen Then
            If rsload.RecordCount > 0 Then
                Do
                    If rsload!bFound = 1 Then
                        cmbVendors(0).AddItem rsload!vName
                        cmbVendors(0).ItemData(cmbVendors(0).ListCount - 1) = rsload!RecID
                        cmbVendors(1).AddItem rsload!vName
                        cmbVendors(1).ItemData(cmbVendors(1).ListCount - 1) = rsload!RecID
                        cmbVendors(2).AddItem rsload!vName
                        cmbVendors(2).ItemData(cmbVendors(2).ListCount - 1) = rsload!RecID
                        cmbVendors(3).AddItem rsload!vName
                        cmbVendors(3).ItemData(cmbVendors(3).ListCount - 1) = rsload!RecID
                    End If
                    rsload.MoveNext
                Loop Until rsload.EOF
            End If
        End If
    End If
    
    bResult = MySQL.OpenTable(aConn, rsload, , "select * from servicetypes")
    
    tvServiceTypes.NodeS.Clear
    
    
    If rsload.State = adStateOpen Then
    
        If rsload.RecordCount > 0 Then
            While Not rsload.EOF And Err.Number = 0
                If Not IsNull(rsload!SubofRecID) Then
                    If rsload!SubofRecID <> 0 Then
                        Set NodeX = tvServiceTypes.NodeS("k" & rsload!SubofRecID)
                        Set NodX = tvServiceTypes.NodeS.Add(NodeX.Key, tvwChild, "k" & rsload!RecID, IIf(IsNull(rsload!Description), "Description Not set", rsload!Description))
                        NodX.Tag = rsload!ServiceKey
                        'NodeX.Expanded = True
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
        
    End If
    
    bResult = MySQL.OpenTable(aConn, rsload, , "select * from radiuspools")
    If rsload.RecordCount > 0 Then
        While Not rsload.EOF And Err.Number = 0
            cmdRadius.AddItem IIf(IsNull(rsload!Description), "(NULL)", rsload!Description)
            cmdRadius.ItemData(cmdRadius.ListCount - 1) = rsload!RecID
            rsload.MoveNext
            gSleep
        Wend
    End If
    
    'Exit Function
    
    On Error Resume Next
    cmbAllPLans(Index).Clear
    bResult = MySQL.OpenTable(aConn, rsPlanTemp, , "select * from plantypes")
    If rsPlanTemp.RecordCount > 0 Then
        While Not rsPlanTemp.EOF And Err.Number = 0
            cmbAllPLans(Index).AddItem IIf(IsNull(rsPlanTemp!Description), "(NULL)", rsPlanTemp!Description)
            cmbAllPLans(Index).ItemData(cmbAllPLans(Index).ListCount - 1) = rsPlanTemp!RecID
            rsPlanTemp.MoveNext
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
    Const ContainerName = "frmTemplateConfig"
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


    Call GUI.SaveColWidths(lvAccounts, Me)
    Call GUI.SaveColWidths(lvPlans(0), Me)
    Call GUI.SaveColWidths(lvPlans(1), Me)
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
    Const ContainerName = "frmTemplateConfig"
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

        
    
    MySQL.Execute aConn, "Update plantemplates set `" & Left(lvAccounts.ColumnHeaders(1).Tag, InStr(lvAccounts.ColumnHeaders(1).Tag, ">") - 1) & "` = '" & MySQL.ESC(NewString) & "' where RecID = " & txtDesc.Tag
  

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
    Const ContainerName = "frmTemplateConfig"
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

    Dim Xnt As Integer
    
    For Xnt = Command1.LBound To Command1.UBound
    
        Command1(Xnt).Enabled = True
        
    Next
    
    bResult = MySQL.OpenTable(aConn, rsPlanTemp, , "select * from plantemplates Where RecID = " & Mid(Item.Key, 2))
        
    LoadFlags Mid(Item.Key, 2)
    
    If rsPlanTemp.RecordCount > 0 Then
    
        chkHidden.Value = rsPlanTemp!Hidden
        If IsNull(rsPlanTemp!VendorID) Then
            cmbVendors(0).ListIndex = -1
            Call cmbVendors_Click(0)
        Else
            Dim lx As Long
            For lx = 0 To cmbVendors(0).ListCount - 1
                If cmbVendors(0).ItemData(lx) = rsPlanTemp!VendorID Then
                    cmbVendors(0).ListIndex = lx
                    Call cmbVendors_Click(0)
                    Exit For
                End If
            Next
        End If
        txtDesc = IIf(IsNull(rsPlanTemp!Description), "", rsPlanTemp!Description)
        
        If rsPlanTemp!CategoryID = 0 Then
            Set tvCat.SelectedItem = Nothing
            lblCat.Caption = ""
        Else
            tvCat.NodeS("r" & rsPlanTemp!CategoryID).Selected = True
            lblCat.Caption = tvCat.NodeS("r" & rsPlanTemp!CategoryID).Text
        End If
        
        txtDesc.Tag = rsPlanTemp!RecID
        txtFee(0) = rsPlanTemp!PeriodFee
        txtFee(3) = rsPlanTemp!CostPrice
        txtSessionTimeout.Text = rsPlanTemp!SessionTimeout
        txtIdleTimeout.Text = rsPlanTemp!IdleTimeout
        txtQuota(0).Text = "" & rsPlanTemp!MBQuota
        txtPartID = IIf(IsNull(rsPlanTemp!VendorPartID), "", rsPlanTemp!VendorPartID)
        txtSubPartID = IIf(IsNull(rsPlanTemp!SubPartID), "", rsPlanTemp!SubPartID)
        Command1(0).Tag = IIf(IsNull(rsPlanTemp!ProductText), "", rsPlanTemp!ProductText)
        
        Dim X As Long
        
        For X = cmbVendors(0).ListCount - 1 To 0
            If cmbVendors(0).ItemData(X) = rsPlanTemp!VendorID Then
                cmbVendors(0).ListIndex = X
                Exit For
            End If
        Next
        
        If rsPlanTemp!MBPerPeriod <> -1 Then
            chkLimit(1).Value = 1
            Call chkLimit_Click(1)
            txtMB(1) = rsPlanTemp!MBPerPeriod
            txtMB(0) = rsPlanTemp!MBBlockSize
            txtFee(1) = rsPlanTemp!FeePerBlock
            txtFee(4) = rsPlanTemp!MBCostPrice
        Else
            chkLimit(1).Value = 0
            Call chkLimit_Click(0)
            txtMB(0) = "0"
            txtMB(1) = "0"
            txtFee(1) = "0.00"
            txtFee(4) = "0.00"
        End If
        
        
        If rsPlanTemp!HoursPerPeriod <> -1 Then
            chkLimit(0).Value = 1
            Call chkLimit_Click(0)
            txtHours = rsPlanTemp!HoursPerPeriod
            txtFee(2) = rsPlanTemp!ExtraPerHour
            txtFee(5) = rsPlanTemp!PeriodCostPrice
        Else
            chkLimit(0).Value = 0
            Call chkLimit_Click(0)
            txtHours = "0"
            txtFee(2) = "0.00"
            txtFee(5) = "0.00"
        End If
        
        If Not IsNull(rsPlanTemp!RadiusID) Then
            Dim ix As Integer
            For ix = 0 To cmdRadius.ListCount - 1
                If cmdRadius.ItemData(ix) = rsPlanTemp!RadiusID Then
                   cmdRadius.ListIndex = ix
                   Exit For
                End If
            Next
        End If
        
        Call cmdContract_Click(1)
        
        Dim rsload As ADODB.Recordset
        
        
        Call MySQL.OpenTable(ADOConn, rsload, , MySQL.virtualisp("select * from contracttemplates where ptRecID = " & rsPlanTemp!RecID & " and bDeleted = 0", "contracttemplates", False, Login.bMaster))
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
                    rsload.MoveNext
                Wend
            End If
        End If
        
    End If
    
    
    Select Case True
    Case picSet(3).Visible
        
        Dim csql As String
                
        Dim Xh As Integer
        
        For Xh = chkCont.LBound To chkCont.UBound
            
            
            chkCont(X).Value = Val(rsPlanTemp("c" & MySQL.ReplaceString(chkCont(X).Caption, " ", "_")))
            
        Next Xh
        
        
        
        txtUnits = IIf(IsNull(rsPlanTemp!unitperpack), "", rsPlanTemp!unitperpack)
        txtDem(0).Text = IIf(IsNull(rsPlanTemp!Height), "", rsPlanTemp!unitperpack)
        txtDem(1).Text = IIf(IsNull(rsPlanTemp!Length), "", rsPlanTemp!unitperpack)
        txtDem(2).Text = IIf(IsNull(rsPlanTemp!Depth), "", rsPlanTemp!unitperpack)
        txtDem(3).Text = IIf(IsNull(rsPlanTemp!Weight), "", rsPlanTemp!unitperpack)
        
            
        For Xh = optPack.UBound To optPack.LBound Step -1
            optPack(Xh).Value = False
            If Not IsNull(rsPlanTemp!packaging) Then
                If optPack(Xh).Caption = rsPlanTemp!packaging Then
                    optPack(Xh).Value = True
                Else
                    optPack(Xh).Value = False
                    If X = 0 Then
                        optPack(Xh).Value = True
                        cmbPack.Text = rsPlanTemp!packaging
                    End If
                End If
            End If
        Next
            
    Case picSet(2).Visible
    
        txtLoc.Text = IIf(IsNull(rsPlanTemp!Location), "", rsPlanTemp!Location)
        chkRates.Value = Val(IIf(IsNull(rsPlanTemp!chargebyrate), 0, rsPlanTemp!chargebyrate))
        
        
        If chkRates.Value = 1 Or chkRates.Value = 2 Then
            Dim Y As Integer
                
            Select Case IIf(IsNull(rsPlanTemp!ratetype), "", rsPlanTemp!ratetype)
            
            Case "n"
                If IIf(IsNull(rsPlanTemp!rateinterval), "", rsPlanTemp!rateinterval) = 30 Then
                    optRates(0).Value = True
                Else
                    optRates(1).Value = True
                End If
            Case "d"
                optRates(2).Value = True
            Case "ww"
                optRates(3).Value = True
            Case "m"
                optRates(4).Value = True
            Case "q"
                optRates(5).Value = True
            Case "y"
                optRates(6).Value = True
            End Select
        End If
        
    Case Else
    
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

Private Sub lvAccounts_KeyUp(KeyCode As Integer, Shift As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvAccounts_KeyUp"
    Const ContainerName = "frmTemplateConfig"
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

Private Sub lvContracts_AfterLabelEdit(Cancel As Integer, NewString As String)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvContracts_AfterLabelEdit"
    Const ContainerName = "frmTemplateConfig"
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


    MySQL.Execute ADOConn, "Update contracttemplates set `" & lvContracts.ColumnHeaders(1).Tag & "` = '" & MySQL.ESC(NewString) & "' where RecID = " & lvContracts.Tag
    
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

Private Sub lvContracts_Delete_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvContracts_Delete_Click"
    Const ContainerName = "frmTemplateConfig"
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


    If Not lvPlans(Index).SelectedItem Is Nothing Then
        MySQL.Execute ADOConn, "delete from flags_tempextras where RecID = " & Mid(lvPlans(Index).SelectedItem.Key, 2)
        lvPlans(Index).ListItems.Remove lvPlans(Index).SelectedItem.Index
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

Private Sub lvContracts_Import_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvContracts_Import_Click"
    Const ContainerName = "frmTemplateConfig"
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


    Dim frmIMP As New frmImpContracts
    
    frmIMP.ptRecID = Val(txtDesc.Tag)
    frmIMP.Show 1
    
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

Private Sub lvContracts_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvContracts_ItemClick"
    Const ContainerName = "frmTemplateConfig"
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


    txtField(0).Text = Item.Text
    txtNoIntervals.Text = Item.SubItems(1)
    UpDown1.Value = Val(Item.SubItems(1))
    
    Dim bx As Byte
    
    For bx = 0 To 7
        If Option1(bx).Tag = Item.SubItems(2) Then Option1(bx).Value = True
    Next bx
    
    txtFee(6) = Item.SubItems(3)
    txtFee(7) = Item.SubItems(5)
    txtFee(8) = Item.SubItems(4)
    txtFee(9) = Item.SubItems(6)
    txtFee(10) = Item.SubItems(7)
    
    lvContracts.Tag = Item.SubItems(8)
    
    LoadFlags lvContracts.Tag, True
    
    
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

Private Sub lvContracts_KeyUp(KeyCode As Integer, Shift As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvContracts_KeyUp"
    Const ContainerName = "frmTemplateConfig"
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
    Case vbKeyDelete
        If Not lvPlans(Index).SelectedItem Is Nothing Then
            MySQL.Execute ADOConn, "delete from flags_tempextras where ContractID = " & Mid(lvContracts.SelectedItem.Key, 2)
            MySQL.Execute ADOConn, "update contracttemplates set bDeleted = -1 where RecID = " & Mid(lvContracts.SelectedItem.Key, 2)
            lvContracts.ListItems.Remove lvContracts.SelectedItem.Index
        End If
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

Private Sub lvContracts_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvContracts_MouseDown"
    Const ContainerName = "frmTemplateConfig"
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


    
    If Button = 2 Then PopupMenu lvPlans_lvContracts
    
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

Private Sub lvPlans_ItemCheck(Index As Integer, ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvPlans_ItemCheck"
    Const ContainerName = "frmTemplateConfig"
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

Private Sub lvPlans_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvPlans_KeyUp"
    Const ContainerName = "frmTemplateConfig"
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
    Case vbKeyDelete
        If Not lvPlans(Index).SelectedItem Is Nothing Then
            Select Case Index
            Case 0
                MySQL.Execute ADOConn, "delete from flags_tempextras where RecID = " & Mid(lvPlans(0).SelectedItem.Key, 2)
            Case 1
            
                MySQL.Execute ADOConn, "delete from flags_tempextras where RecID = " & Mid(lvPlans(1).SelectedItem.Key, 2)
                'MySQL.Execute ADOConn, "delete from contracttemplates where RecID = " & Mid(lvPlans(1).selectedItem.Key, 2)
            End Select
            lvPlans(Index).ListItems.Remove lvPlans(Index).SelectedItem.Index
        End If
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

Private Sub ts_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "ts_Click"
    Const ContainerName = "frmTemplateConfig"
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


    Select Case txtDesc.Tag
    Case "", 0
        
        If ts.SelectedItem.Index <> 1 Then ts.Tabs(1).Selected = True
        
        For ix = frmts.LBound To frmts.UBound
            If ix <> ts.SelectedItem.Index - 1 Then
                frmts(ix).Visible = False
            Else
                frmts(ix).Move ts.ClientLeft, ts.ClientTop, ts.ClientWidth, ts.ClientHeight
                frmts(ix).ZOrder 0
                frmts(ix).Visible = True
            End If
        Next ix
        
    Case Else
    
        
        
        For ix = frmts.LBound To frmts.UBound
            If ix <> ts.SelectedItem.Index - 1 Then
                frmts(ix).Visible = False
            Else
                frmts(ix).Move ts.ClientLeft, ts.ClientTop, ts.ClientWidth, ts.ClientHeight
                frmts(ix).ZOrder 0
                frmts(ix).Visible = True
            End If
        Next ix
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

Private Sub tvCat_NodeClick(ByVal Node As MSComctlLib.Node)

    lblCat.Caption = Node.Text
    
End Sub

Private Sub tvservicetypes_NodeClick(ByVal Node As MSComctlLib.Node)

                        

    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "tvservicetypes_NodeClick"
    Const ContainerName = "frmTemplateConfig"
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
    Dim bResult As Boolean
    Dim Xnt As Integer
    
    Frame1(0).Caption = "Services and Plans: " & Node.Text
    
    For Xnt = Command1.LBound To Command1.UBound
        Command1(Xnt).Enabled = False
    Next
    
    bResult = MySQL.OpenTable(aConn, rsPlanTemp, , MySQL.virtualisp("select * from plantemplates Where ServiceID = " & Mid(Node.Key, 2), "plantemplates", False, Login.bMaster))
    
    'rsPlanTemp.Filter
    
    Select Case tvServiceTypes.SelectedItem.Tag
    Case "DIALUP", "ADSL", "SHDSL"
        cmdRadius.Enabled = True
        frameLimit(0).Enabled = True
        frameLimit(1).Enabled = True
        frameLimit(2).Enabled = True
    Case "FTP", "POP3", "WWW", "DESIGN", "DOMAIN", "GATEWAY", "TRAINING", "SALES", "CONSULT", "COLO", "HOST"
        cmdRadius.Enabled = False
        frameLimit(0).Enabled = False
        frameLimit(1).Enabled = False
        frameLimit(2).Enabled = False
    End Select
    
    
    
    
    For Xnt = picSet.LBound To picSet.UBound
        If InStr(picSet(Xnt).Tag, tvServiceTypes.SelectedItem.Tag) > 0 Then
            picSet(Xnt).Move iLeft, iTop, iWidth, iHeight
            picSet(Xnt).ZOrder 0
            picSet(Xnt).Visible = True
        Else
            picSet(Xnt).Visible = False
        End If
    Next
        
        
    
    Dim bNumeric As Boolean
    
    Dim SQL As String
    
    lvPlans(Index).ListItems.Clear
    MySQL.fillLV ADOConn, rsPlanTemp, lvAccounts, False, , 1
    
    

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
    Const ContainerName = "frmTemplateConfig"
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

Private Sub txtFee_KeyPress(Index As Integer, KeyAscii As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtFee_KeyPress"
    Const ContainerName = "frmTemplateConfig"
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
    Const ContainerName = "frmTemplateConfig"
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

Private Sub txtMB_KeyPress(Index As Integer, KeyAscii As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtMB_KeyPress"
    Const ContainerName = "frmTemplateConfig"
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


Private Sub txtNumOf_KeyPress(Index As Integer, KeyAscii As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtNumOf_KeyPress"
    Const ContainerName = "frmTemplateConfig"
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

Private Sub txtQuota_Change(Index As Integer)

    Dim Xnt As Integer
    
    For Xnt = txtQuota.LBound To txtQuota.UBound
        If Xnt <> Index Then
            txtQuota(Xnt).Text = txtQuota(Index).Text
        End If
    Next
    
End Sub
