VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCommissions 
   Caption         =   "Forecast Commissions"
   ClientHeight    =   10065
   ClientLeft      =   1575
   ClientTop       =   2220
   ClientWidth     =   15900
   Icon            =   "frmCommissions_new.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10065
   ScaleWidth      =   15900
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cd 
      Left            =   9150
      Top             =   2580
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picTSContainer 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   8955
      Index           =   0
      Left            =   150
      ScaleHeight     =   8955
      ScaleWidth      =   15315
      TabIndex        =   2
      Top             =   1080
      Width           =   15315
      Begin VB.Frame frmAssign 
         BackColor       =   &H0086D28D&
         Caption         =   "Rates Class Assignment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4545
         Left            =   120
         TabIndex        =   35
         Top             =   4290
         Width           =   6105
         Begin MSComctlLib.ImageList ilSysops 
            Left            =   3450
            Top             =   450
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
                  Picture         =   "frmCommissions_new.frx":030A
                  Key             =   "k100"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCommissions_new.frx":0624
                  Key             =   "k080_099"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCommissions_new.frx":093E
                  Key             =   "k060_079"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCommissions_new.frx":1218
                  Key             =   "k040_059"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCommissions_new.frx":166A
                  Key             =   "k001_020"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCommissions_new.frx":1ABC
                  Key             =   "k020_039"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCommissions_new.frx":1F0E
                  Key             =   "k000"
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView lvSysops 
            Height          =   3225
            Index           =   2
            Left            =   150
            TabIndex        =   40
            Top             =   1170
            Width           =   3105
            _ExtentX        =   5477
            _ExtentY        =   5689
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393217
            SmallIcons      =   "ilSysops"
            ForeColor       =   5670752
            BackColor       =   -2147483643
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
               Text            =   "Username"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Level"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Description"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "ViSP"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H0086D28D&
            Caption         =   "Selection text"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   150
            TabIndex        =   37
            Top             =   240
            Width           =   5835
            Begin VB.CheckBox chkPercentile 
               BackColor       =   &H0086D28D&
               Caption         =   "Calculate by Percenitile Only."
               Enabled         =   0   'False
               Height          =   405
               Left            =   4200
               TabIndex        =   39
               Top             =   270
               Width           =   1545
            End
            Begin VB.Label lblTxt 
               Alignment       =   2  'Center
               BackColor       =   &H0086D28D&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   510
               Left            =   135
               TabIndex        =   38
               Top             =   210
               Width           =   3975
            End
         End
         Begin VB.ListBox lstRates 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00568760&
            Height          =   3060
            Index           =   1
            Left            =   3330
            TabIndex        =   36
            Top             =   1170
            Width           =   2655
         End
      End
      Begin VB.Frame frmScales 
         BackColor       =   &H000040C0&
         Caption         =   "Rates and Scales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0043D143&
         Height          =   8745
         Left            =   6330
         TabIndex        =   11
         Top             =   120
         Width           =   9135
         Begin VB.CommandButton cmdSave 
            BackColor       =   &H00EFEF98&
            Caption         =   "&Refresh Records on server with this matrix."
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
            Left            =   120
            MaskColor       =   &H00EFEF98&
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   1260
            Width           =   4185
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H000040C0&
            Caption         =   "Plans and Services - Commission Scales"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   7005
            Left            =   120
            TabIndex        =   14
            Top             =   1650
            Width           =   8865
            Begin VB.Frame Frame7 
               BackColor       =   &H00EC7A71&
               Caption         =   "Scales"
               ForeColor       =   &H00FFFFFF&
               Height          =   2115
               Left            =   180
               TabIndex        =   56
               Top             =   4800
               Width           =   8505
               Begin MSComctlLib.ListView lvScale 
                  Height          =   1635
                  Left            =   150
                  TabIndex        =   57
                  Top             =   300
                  Width           =   8205
                  _ExtentX        =   14473
                  _ExtentY        =   2884
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   -1  'True
                  HideSelection   =   -1  'True
                  FullRowSelect   =   -1  'True
                  GridLines       =   -1  'True
                  HotTracking     =   -1  'True
                  _Version        =   393217
                  ForeColor       =   15497841
                  BackColor       =   -2147483643
                  BorderStyle     =   1
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
                  NumItems        =   5
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "Min"
                     Object.Width           =   2540
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Text            =   "Max"
                     Object.Width           =   2540
                  EndProperty
                  BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   1
                     SubItemIndex    =   2
                     Text            =   "Commission Per Unit"
                     Object.Width           =   2540
                  EndProperty
                  BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   1
                     SubItemIndex    =   3
                     Text            =   "Percentile"
                     Object.Width           =   2540
                  EndProperty
                  BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   1
                     SubItemIndex    =   4
                     Text            =   "Total Block Commission"
                     Object.Width           =   2540
                  EndProperty
               End
            End
            Begin VB.Frame frameCalc 
               BackColor       =   &H8000000C&
               Caption         =   "Rates"
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
               Height          =   1755
               Left            =   180
               TabIndex        =   16
               Top             =   3030
               Width           =   8505
               Begin VB.CommandButton cmdBonuses 
                  BackColor       =   &H0080FF80&
                  Caption         =   "&Bonuses"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   465
                  Left            =   7050
                  MaskColor       =   &H0080FF80&
                  Style           =   1  'Graphical
                  TabIndex        =   55
                  Top             =   1170
                  Width           =   1335
               End
               Begin VB.CommandButton cmdUpdate 
                  BackColor       =   &H0080FF80&
                  Caption         =   "&Update"
                  Enabled         =   0   'False
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
                  Left            =   7050
                  MaskColor       =   &H0080FF80&
                  Style           =   1  'Graphical
                  TabIndex        =   33
                  Top             =   720
                  Width           =   1335
               End
               Begin VB.CommandButton cmdAdd 
                  BackColor       =   &H0080FF80&
                  Caption         =   "&Add"
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
                  Left            =   7050
                  MaskColor       =   &H0080FF80&
                  Style           =   1  'Graphical
                  TabIndex        =   32
                  Top             =   300
                  Width           =   1365
               End
               Begin VB.Frame Frame8 
                  BackColor       =   &H00FFC0C0&
                  Caption         =   "Margin"
                  Height          =   645
                  Index           =   6
                  Left            =   2940
                  TabIndex        =   29
                  Top             =   990
                  Width           =   2385
                  Begin VB.TextBox txtMargin 
                     Alignment       =   2  'Center
                     BackColor       =   &H00FFC0C0&
                     BorderStyle     =   0  'None
                     Height          =   285
                     Left            =   90
                     Locked          =   -1  'True
                     TabIndex        =   30
                     Top             =   270
                     Width           =   2205
                  End
               End
               Begin VB.Frame Frame8 
                  BackColor       =   &H00FFC0C0&
                  Caption         =   "Cost"
                  Height          =   645
                  Index           =   5
                  Left            =   1530
                  TabIndex        =   27
                  Top             =   990
                  Width           =   1395
                  Begin VB.TextBox txtCost 
                     Alignment       =   2  'Center
                     BackColor       =   &H00FFC0C0&
                     BorderStyle     =   0  'None
                     Height          =   285
                     Left            =   90
                     Locked          =   -1  'True
                     TabIndex        =   28
                     Top             =   270
                     Width           =   1185
                  End
               End
               Begin VB.Frame Frame8 
                  BackColor       =   &H00FFC0C0&
                  Caption         =   "RRP"
                  Height          =   645
                  Index           =   4
                  Left            =   120
                  TabIndex        =   25
                  Top             =   990
                  Width           =   1395
                  Begin VB.TextBox txtRRP 
                     Alignment       =   2  'Center
                     BackColor       =   &H00FFC0C0&
                     BorderStyle     =   0  'None
                     Height          =   285
                     Left            =   90
                     Locked          =   -1  'True
                     TabIndex        =   26
                     Top             =   270
                     Width           =   1185
                  End
               End
               Begin VB.Frame Frame8 
                  BackColor       =   &H00FFC0C0&
                  Caption         =   "(%) && Totals"
                  Height          =   1335
                  Index           =   3
                  Left            =   5340
                  TabIndex        =   23
                  Top             =   300
                  Width           =   1635
                  Begin VB.TextBox txtPer 
                     Alignment       =   2  'Center
                     BackColor       =   &H00FFC0C0&
                     BorderStyle     =   0  'None
                     Height          =   285
                     Index           =   1
                     Left            =   90
                     Locked          =   -1  'True
                     TabIndex        =   31
                     Top             =   630
                     Width           =   1485
                  End
                  Begin VB.TextBox txtPer 
                     Alignment       =   2  'Center
                     BackColor       =   &H00FFC0C0&
                     BorderStyle     =   0  'None
                     Height          =   285
                     Index           =   0
                     Left            =   90
                     Locked          =   -1  'True
                     TabIndex        =   24
                     Top             =   300
                     Width           =   1485
                  End
               End
               Begin VB.Frame Frame8 
                  BackColor       =   &H000080FF&
                  Caption         =   "Commission Rate Per Unit"
                  Height          =   645
                  Index           =   2
                  Left            =   2940
                  TabIndex        =   21
                  Top             =   330
                  Width           =   2385
                  Begin VB.TextBox txtQuanity 
                     Alignment       =   2  'Center
                     BackColor       =   &H000080FF&
                     BorderStyle     =   0  'None
                     Height          =   285
                     Index           =   2
                     Left            =   90
                     TabIndex        =   22
                     Top             =   270
                     Width           =   2205
                  End
               End
               Begin VB.Frame Frame8 
                  BackColor       =   &H0080C0FF&
                  Caption         =   "Max Quanity"
                  Height          =   645
                  Index           =   1
                  Left            =   1530
                  TabIndex        =   19
                  Top             =   330
                  Width           =   1395
                  Begin VB.TextBox txtQuanity 
                     Alignment       =   2  'Center
                     BackColor       =   &H0080C0FF&
                     BorderStyle     =   0  'None
                     Height          =   285
                     Index           =   1
                     Left            =   90
                     TabIndex        =   20
                     Top             =   270
                     Width           =   1215
                  End
               End
               Begin VB.Frame Frame8 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "Min Quanity"
                  Height          =   645
                  Index           =   0
                  Left            =   120
                  TabIndex        =   17
                  Top             =   330
                  Width           =   1395
                  Begin VB.TextBox txtQuanity 
                     Alignment       =   2  'Center
                     BackColor       =   &H00C0E0FF&
                     BorderStyle     =   0  'None
                     Height          =   285
                     Index           =   0
                     Left            =   90
                     TabIndex        =   18
                     Top             =   270
                     Width           =   1215
                  End
               End
            End
            Begin MSComctlLib.ListView lvPlans 
               Height          =   2625
               Left            =   150
               TabIndex        =   15
               Top             =   330
               Width           =   8565
               _ExtentX        =   15108
               _ExtentY        =   4630
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   16777215
               BackColor       =   8421504
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
               NumItems        =   8
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Description"
                  Object.Width           =   9878
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "Cost"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Cost (Inc Tax)"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Vendor"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "ViSP"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   5
                  Text            =   "RRP"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   6
                  Text            =   "RRP (Inc tax)"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   7
                  Text            =   "Margin"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Commission Class Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   885
            Left            =   120
            TabIndex        =   12
            Top             =   300
            Width           =   8835
            Begin VB.CommandButton cmdCreate 
               BackColor       =   &H00EFEF98&
               Caption         =   "&Create"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   7650
               MaskColor       =   &H00EFEF98&
               Style           =   1  'Graphical
               TabIndex        =   34
               Top             =   300
               Width           =   1035
            End
            Begin VB.ComboBox cmbClass 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Left            =   150
               TabIndex        =   13
               Top             =   300
               Width           =   7425
            End
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H0047ADE4&
         Caption         =   "Select Date to Calculate Between"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   6105
         Begin VB.Frame Frame3 
            BackColor       =   &H0047ADE4&
            Caption         =   "End Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2895
            Index           =   1
            Left            =   3120
            TabIndex        =   9
            Top             =   240
            Width           =   2895
            Begin MSComCtl2.MonthView MonthView 
               Height          =   2370
               Index           =   1
               Left            =   90
               TabIndex        =   10
               Top             =   330
               Width           =   2700
               _ExtentX        =   4763
               _ExtentY        =   4180
               _Version        =   393216
               ForeColor       =   -2147483630
               BackColor       =   -2147483633
               Appearance      =   1
               StartOfWeek     =   20578306
               CurrentDate     =   37682
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H0047ADE4&
            Caption         =   "Start Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2895
            Index           =   0
            Left            =   150
            TabIndex        =   7
            Top             =   240
            Width           =   2895
            Begin MSComCtl2.MonthView MonthView 
               Height          =   2370
               Index           =   0
               Left            =   90
               TabIndex        =   8
               Top             =   360
               Width           =   2700
               _ExtentX        =   4763
               _ExtentY        =   4180
               _Version        =   393216
               ForeColor       =   -2147483630
               BackColor       =   12632256
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
               StartOfWeek     =   20578306
               CurrentDate     =   37682
            End
         End
         Begin VB.CommandButton cmdCalc 
            BackColor       =   &H0080FFFF&
            Caption         =   "<<< Calculate from Date Selected >>>"
            Default         =   -1  'True
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   3270
            Width           =   5805
         End
      End
   End
   Begin MSComctlLib.TabStrip ts 
      Height          =   9255
      Left            =   90
      TabIndex        =   1
      Top             =   780
      Width           =   15585
      _ExtentX        =   27490
      _ExtentY        =   16325
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Totals"
            Object.ToolTipText     =   "Here is where you select the date to calculate from as well as displaying totals."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sysop Commissions"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picTSContainer 
      BorderStyle     =   0  'None
      Height          =   5415
      Index           =   1
      Left            =   9690
      ScaleHeight     =   5415
      ScaleWidth      =   4155
      TabIndex        =   3
      Top             =   2730
      Width           =   4155
      Begin VB.Frame Frame6 
         Height          =   1005
         Left            =   120
         TabIndex        =   41
         Top             =   4410
         Width           =   8835
         Begin VB.CommandButton cmbExport 
            Caption         =   "Export to CSV"
            Height          =   405
            Left            =   7500
            TabIndex        =   42
            Top             =   390
            Width           =   1245
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Total Commission Paid:"
            Height          =   240
            Index           =   5
            Left            =   3180
            TabIndex        =   54
            Top             =   690
            Width           =   1695
         End
         Begin VB.Label lblComm 
            AutoSize        =   -1  'True
            Caption         =   "$ 0.00"
            Height          =   240
            Left            =   4950
            TabIndex        =   53
            Top             =   690
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Total Paid:"
            Height          =   240
            Index           =   4
            Left            =   90
            TabIndex        =   52
            Top             =   690
            Width           =   795
         End
         Begin VB.Label lblPaid 
            AutoSize        =   -1  'True
            Caption         =   "$ 0.00"
            Height          =   240
            Left            =   1860
            TabIndex        =   51
            Top             =   690
            Width           =   465
         End
         Begin VB.Label lblCredit 
            AutoSize        =   -1  'True
            Caption         =   "$ 0.00"
            Height          =   240
            Left            =   4950
            TabIndex        =   50
            Top             =   450
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Total Credited:"
            Height          =   240
            Index           =   3
            Left            =   3180
            TabIndex        =   49
            Top             =   450
            Width           =   1140
         End
         Begin VB.Label lblDebit 
            AutoSize        =   -1  'True
            Caption         =   "$ 0.00"
            Height          =   240
            Left            =   4950
            TabIndex        =   48
            Top             =   210
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Total Debit"
            Height          =   240
            Index           =   2
            Left            =   3180
            TabIndex        =   47
            Top             =   210
            Width           =   825
         End
         Begin VB.Label lblCost 
            AutoSize        =   -1  'True
            Caption         =   "$ 0.00"
            Height          =   240
            Left            =   1860
            TabIndex        =   46
            Top             =   450
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Total Cost:"
            Height          =   240
            Index           =   1
            Left            =   90
            TabIndex        =   45
            Top             =   450
            Width           =   825
         End
         Begin VB.Label lblMargin 
            AutoSize        =   -1  'True
            Caption         =   "$ 0.00"
            Height          =   240
            Left            =   1860
            TabIndex        =   44
            Top             =   210
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Total Margin (Profit):"
            Height          =   240
            Index           =   0
            Left            =   90
            TabIndex        =   43
            Top             =   210
            Width           =   1590
         End
      End
      Begin MSComctlLib.ListView lvSysops 
         Height          =   4245
         Index           =   0
         Left            =   90
         TabIndex        =   4
         Top             =   90
         Width           =   8865
         _ExtentX        =   15637
         _ExtentY        =   7488
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Username"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "ViSP"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Total Debit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Total Cost"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Total Amount Paid"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Total Amount Credited"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Total Commissions"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Income Tax"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Super Annuation"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Remaining Commission"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Commission Forcaster"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0968F&
      Height          =   360
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   3300
   End
   Begin VB.Image Image1 
      Height          =   690
      Left            =   90
      Picture         =   "frmCommissions_new.frx":2360
      Stretch         =   -1  'True
      Top             =   30
      Width           =   690
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   15900
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "frmCommissions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oComm As New clsCommission
Dim itxM As ListItem


Private Sub cmbClass_Change()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmbClass_Change"
    Const ContainerName = "frmCommissions"
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


    lvScale.ListItems.Clear
    
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

Private Sub cmbClass_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmbClass_Click"
    Const ContainerName = "frmCommissions"
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


    lvScale.ListItems.Clear
    
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

Private Sub cmbClass_GotFocus()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmbClass_GotFocus"
    Const ContainerName = "frmCommissions"
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


    Static bPlay As Boolean
    Dim oXY As POINTAPI, L As Long
    
    If bPlay = False And picTSContainer(0).Visible = True Then
            
        L = GetCursorPos(oXY)
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "Type in the new name for the commision class and press create otherwise use the down arrow to select which one you are working with."
        bPlay = True
    
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

Private Sub cmbExport_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmbExport_Click"
    Const ContainerName = "frmCommissions"
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


    cd.Filter = "CSV FILE (*.CSV)|*.CSV"
    cd.FilterIndex = 1
    cd.ShowSave
    
    If cd.Filename <> "" Then
    
        Dim Header As String
        
        Header = Header + vbCrLf + Chr$(34) + "Total Amount Paid" + Chr$(34) + "," + Chr$(34) + lblPaid.Caption + Chr$(34) + String(lvSysops(0).ColumnHeaders.Count - 2, ",")
        Header = Header + vbCrLf + Chr$(34) + "Total Cost " + Chr$(34) + "," + Chr$(34) + lblCost.Caption + Chr$(34) + String(lvSysops(0).ColumnHeaders.Count - 2, ",")
        Header = Header + vbCrLf + Chr$(34) + "Total Margin" + Chr$(34) + "," + Chr$(34) + lblMargin.Caption + Chr$(34) + String(lvSysops(0).ColumnHeaders.Count - 2, ",")
        Header = Header + vbCrLf + Chr$(34) + "Total Amount Debit" + Chr$(34) + "," + Chr$(34) + lblDebit.Caption + Chr$(34) + String(lvSysops(0).ColumnHeaders.Count - 2, ",")
        Header = Header + vbCrLf + Chr$(34) + "Total Amount Credited" + Chr$(34) + "," + Chr$(34) + lblCredit.Caption + Chr$(34) + String(lvSysops(0).ColumnHeaders.Count - 2, ",")
        Header = Header + vbCrLf + Chr$(34) + "Total Commission Paid" + Chr$(34) + "," + Chr$(34) + lblComm.Caption + Chr$(34) + String(lvSysops(0).ColumnHeaders.Count - 2, ",")
        Header = Header + vbCrLf + String(lvSysops(0).ColumnHeaders.Count, ",")
        Header = Header + vbCrLf + String(lvSysops(0).ColumnHeaders.Count, ",")
            
        Header = Header + vbCrLf + Chr$(34) + "from" + Chr$(34) + "," + Chr$(34) + Format(MonthView(0).Value, "dddd dd-mm-yyyy") + Chr$(34) + String(lvSysops(0).ColumnHeaders.Count - 2, ",")
        Header = Header + vbCrLf + Chr$(34) + "To" + Chr$(34) + "," + Chr$(34) + Format(MonthView(1).Value, "dddd dd-mm-yyyy") + Chr$(34) + String(lvSysops(0).ColumnHeaders.Count - 2, ",")
        Header = Header + vbCrLf + String(lvSysops(0).ColumnHeaders.Count, ",")
        Header = Header + vbCrLf + String(lvSysops(0).ColumnHeaders.Count, ",")
            
        Call GUI.LV2CSV(cd.Filename, lvSysops(0), Header)
        
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

Private Sub cmdAdd_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdAdd_Click"
    Const ContainerName = "frmCommissions"
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
    
    Dim oRate As clsRates
    
    If cmbClass.ListIndex = -1 Then
        MsgBox "You must select a commission class", vbCritical, "Class not selected"
        Exit Sub
    End If
        
        
    If lvPlans.SelectedItem Is Nothing Then
        MsgBox "You must select a plan for the commission to be calculated for."
        Exit Sub
    End If
    
    If Val(txtQuanity(0)) > 0 And Val(txtQuanity(1)) > Val(txtQuanity(0)) Then
    
    Else
        MsgBox "Your quanities must be greater than zero and the Maximum quanity must be greater than the Minimum."
        Exit Sub
    End If
    
    Set oRate = oComm.colRates.Add(lvPlans.SelectedItem.Key & "_" & oComm.colRates.Count + 1 & "_" & cmbClass.ItemData(cmbClass.ListIndex), , , , , , , lvPlans.SelectedItem.Key & "_" & oComm.colRates.Count + 1)
    oRate.ClassID = cmbClass.ItemData(cmbClass.ListIndex)
    oRate.ptRecID = Val(Mid(lvPlans.SelectedItem.Key, 2))
    oRate.Min = Val(txtQuanity(0))
    oRate.Max = Val(txtQuanity(1))
    oRate.Rate = Val(txtQuanity(2))
    oRate.Margin = 0
    oRate.Flag = 1
    oRate.Percentile = (Val(txtQuanity(2)) * (Val(txtQuanity(1)) - Val(txtQuanity(0)))) / ((Val(txtQuanity(1)) - Val(txtQuanity(0))) * Val(txtMargin.Tag))
    
    Dim itmX As ListItem
    
    Set itmX = lvScale.ListItems.Add(, oRate.Key, oRate.Min)
    itmX.SubItems(1) = oRate.Max
    itmX.SubItems(2) = Format(oRate.Rate, "Currency")
    itmX.SubItems(3) = oRate.Percentile
    itmX.SubItems(4) = Format((oRate.Max - oRate.Min) * oRate.Rate, "currency")
    itmX.Tag = oRate.ptRecID


    
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

Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdAdd_MouseMove"
    Const ContainerName = "frmCommissions"
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



    Static bPlay As Boolean
    Dim oXY As POINTAPI, L As Long
    
    If bPlay = False And picTSContainer(0).Visible = True Then
            
        L = GetCursorPos(oXY)
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "After you have selected a plan or service to add in a commission scale node fill in the details contain withing the Rates frame and click on this green Add button."
        bPlay = True
    
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

Private Sub cmdBonuses_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdBonuses_Click"
    Const ContainerName = "frmCommissions"
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


    Dim fSpiff As New frmSpiff
    fSpiff.ptRecID = cmdBonuses.Tag
    
    fSpiff.Show 1

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

Private Sub cmdBonuses_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdBonuses_MouseMove"
    Const ContainerName = "frmCommissions"
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


    Static bPlay As Boolean
    Dim oXY As POINTAPI, L As Long
    
    If bPlay = False And picTSContainer(0).Visible = True Then
            
        L = GetCursorPos(oXY)
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "IF you would like to add bonues to this group of scales click on this button and add in a bonus."
        bPlay = True
    
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

Private Sub cmdCalc_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdCalc_Click"
    Const ContainerName = "frmCommissions"
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


    
    Dim SQL3 As String
    
        'If bDebug = True Then On Error GoTo 0 Else On Error GoTo ErrorOccur
    
    If lvSysops(0).ListItems.Count > 0 Then
    
        ts.Tabs(2).Selected = True
        ts.Refresh
        gSleep
    
        Dim rsNum As adodb.Recordset
        Dim rsComm As adodb.Recordset
        Dim rsCount As adodb.Recordset
        Dim rsClass As adodb.Recordset
        
        Dim lClassID As Long
        Dim bPercent As Boolean
        
        Dim cTotalCost As Currency
        Dim CTotalDebit As Currency
        Dim cTotalCredit As Currency
        Dim cTotalPaid As Currency
        
        Dim cTTLCost As Currency
        Dim CTTLDebit As Currency
        Dim cTTLCredit As Currency
        Dim cTTLPaid As Currency
        Dim cTTLMArgin As Currency
        Dim cTTLCOMM As Currency
        
        Dim cComm As Currency
        
        Dim lNumOf As Long
        
        Dim lx As Long
        Dim lY As Long
        Dim lZ As Long
        
        For lx = lvSysops(0).ColumnHeaders.Count To 10 Step -1
            lvSysops(0).ColumnHeaders.Remove lx
        Next
        
        
        Dim rsBonus_Units As adodb.Recordset
        Dim rsBonus_awarded As adodb.Recordset
        
        Call MySQL.OpenTable(directConn, rsBonus_Units, , "select distinct * from bonus_units")
        
        If rsBonus_Units.State = adStateOpen Then
            If rsBonus_Units.RecordCount > 0 Then
                rsBonus_Units.MoveFirst
                Do
                    Call lvSysops(0).ColumnHeaders.Add(, "b" & rsBonus_Units!RecID, rsBonus_Units!UnitName & "(bonus)")
                    rsBonus_Units.MoveNext
                Loop Until rsBonus_Units.EOF Or Err.Number <> 0
            End If
        End If
        
        For lx = 1 To lvSysops(0).ListItems.Count
        
            If MySQL.OpenTable(directConn, rsClass, , "select wrkPercent, CommClass, IncomeTax, SuperRate from sysops where RecID = " + Mid(lvSysops(0).ListItems(lx).Key, 2)) = True Then
                If rsClass.BOF And rsClass.EOF Then
                
                    
                    
                Else
                    lClassID = rsClass!CommClass
                    bPercent = Val(IIf(IsNull(rsClass!wrkPercent), 0, rsClass!wrkPercent))
                    
                    If lClassID <> 0 Then
                        For lY = 1 To lvPlans.ListItems.Count
                            If MySQL.OpenTable(directConn, rsCount, , "select count(*) as NumPlans from acci_services where SysopID = " & Mid(lvSysops(0).ListItems(lx).Key, 2) & " and ptRecID = " & Mid(lvPlans.ListItems(lY).Key, 2)) = True Then
                                lNumOf = rsCount!NumPlans
                            Else
                                lNumOf = 0
                            End If
                        
                            If lNumOf > 0 Then
                                If MySQL.OpenTable(directConn, rsNum, , "select invoiceout.*, acci_services.ptRecID as ptRecID from invoiceout, acci_services where (invoiceout.Created >= " + Format(MonthView(0).Value, "yyyymmdd000001") + " and invoiceout.Created <= " + Format(MonthView(1).Value, "yyyymmdd240000") + ") and invoiceout.PlanServiceID = acci_services.RecID and acci_services.SysopID = " & Mid(lvSysops(0).ListItems(lx).Key, 2) & " and acci_services.ptRecID = " & Mid(lvPlans.ListItems(lY).Key, 2)) = True Then
                                    If rsNum.BOF And rsNum.EOF Then
                                                                    
                                        
                                    Else
                                        If rsNum.RecordCount > 0 Then
                                            While Not rsNum.EOF And Err.Number = 0
                                                cTotalCost = cTotalCost + CCur(lvPlans.ListItems("r" & rsNum!ptRecID).SubItems(1))
                                                CTotalDebit = CTotalDebit + rsNum!TotalDue
                                                cTotalCredit = cTotalCredit + rsNum!AmountRefunded + rsNum!GSTRefunded
                                                cTotalPaid = cTotalPaid + rsNum!AmountPaid
                                                rsNum.MoveNext
                                            Wend
                                            
                                            If MySQL.OpenTable(directConn, rsCount, , "select CommPerUnit, Percentile from CommRates Where ClassID = " & lClassID & " and ptRecID = " & Mid(lvPlans.ListItems(lY).Key, 2) & " and Min >= " & lNumOf & " and Max =< " & lNumOf) = True Then
                                                If rsCount.EOF And rsCount.BOF Then
                                                
                                                Else
                                                    If bPercent = False Then
                                                        cComm = cComm + (rsCount!CommPerUnit * rsNum.RecordCount)
                                                    ElseIf bPercent = True Then
                                                        cComm = cComm + ((CTotalDebit - cTotalCredit) - cTotalCost) * rsCount!Percentile
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    Else
                    
                        SQL = ""
                        SQL = SQL + " IN (" & Mid(lvPlans.ListItems(1).Key, 2) & ""
                        
                        If lvPlans.ListItems.Count >= 2 Then
                            For lY = 2 To lvPlans.ListItems.Count
                                SQL = SQL + "', " & Mid(lvPlans.ListItems(lY).Key, 2)
                            Next
                        End If
                        
                        SQL = SQL + ") as bFound"
                        
                        If MySQL.OpenTable(directConn, rsNum, , "select invoiceout.*, acci_services.ptRecID as ptRecID, acci_services.ptRecID " & SQL & " from invoiceout, acci_services where (invoiceout.Created >= " + Format(MonthView(0).Value, "yyyymmdd000001") + " and invoiceout.Created <= " + Format(MonthView(1).Value, "yyyymmdd240000") + ") and invoiceout.PlanServiceID = acci_services.RecID and acci_services.SysopID = " & Mid(lvSysops(0).ListItems(lx).Key, 2)) = True Then
                            If rsNum.BOF And rsNum.EOF Then
                                                            
                                
                            Else
                                If rsNum.RecordCount > 0 Then
                                    
                                    On Error Resume Next
                                    
                                    While Not rsNum.EOF And Err.Number = 0
                                        If rsNum!bFound = 1 Then
                                            cTotalCost = cTotalCost + CCur(lvPlans.ListItems("r" & rsNum!ptRecID).SubItems(1))
                                            CTotalDebit = CTotalDebit + IIf(IsNull(rsNum!TotalDue), 0, rsNum!TotalDue)
                                            cTotalCredit = cTotalCredit + IIf(IsNull(rsNum!AmountRefunded), 0, rsNum!AmountRefunded) + IIf(IsNull(rsNum!GSTRefunded), 0, rsNum!GSTRefunded)
                                            cTotalPaid = cTotalPaid + rsNum!AmountPaid
                                        End If
                                        rsNum.MoveNext
                                    Wend
                                    
                                End If
                            End If
                        End If
                        
                    End If
                End If
             
                lvSysops(0).ListItems(lx).SubItems(2) = Format(CTotalDebit, "Currency")
                lvSysops(0).ListItems(lx).SubItems(3) = Format(cTotalCost, "Currency")
                lvSysops(0).ListItems(lx).SubItems(4) = Format(cTotalPaid, "Currency")
                lvSysops(0).ListItems(lx).SubItems(5) = Format(cTotalCredit, "Currency")
                
                    lvSysops(0).ListItems(lx).SubItems(6) = Format(cComm, "Currency")
                    lvSysops(0).ListItems(lx).SubItems(7) = Format(cComm * (IIf(IsNull(rsClass!IncomeTax), 0, rsClass!IncomeTax) / 100), "Currency")
                    lvSysops(0).ListItems(lx).SubItems(8) = Format(cComm * (IIf(IsNull(rsClass!SuperRate), 0, rsClass!SuperRate) / 100), "Currency")
                    lvSysops(0).ListItems(lx).SubItems(9) = Format(cComm - ((cComm * (IIf(IsNull(rsClass!IncomeTax), 0, rsClass!IncomeTax) / 100)) + (cComm * (IIf(IsNull(rsClass!SuperRate), 0, rsClass!SuperRate) / 100))), "Currency")
                
                rsBonus_Units.MoveFirst
            
                For lZ = 10 To lvSysops(0).ColumnHeaders.Count - 1
                 
                     Call MySQL.OpenTable(directConn, rsBonus_awarded, , "select sum(`" & rsBonus_Units!FieldName & "`) as nResult from bonus_matrix where (bonus_matrix.Created >= " + Format(MonthView(0).Value, "yyyymmdd000001") + " and bonus_matrix.Created <= " + Format(MonthView(1).Value, "yyyymmdd240000") + ") and AwardTo = '" & tosysop & "' and SysopID = '" & Mid(lvSysops(0).ListItems(lx).Key, 2) + "'")
                
                     lvSysops(0).ListItems(lx).SubItems(lZ) = "" & Format(rsBonus_awarded!nResult, rsBonus_Units!Format)
                     rsBonus_Units.MoveNext
                 
                Next lZ
                
                
                
                cTTLCost = cTTLCost + cTotalCost
                CTTLDebit = CTTLDebit + CTotalDebit
                cTTLPaid = cTTLPaid + cTotalPaid
                cTTLCredit = cTTLCredit = cTotalCredit
                cTTLCOMM = cTTLCOMM + cComm
                cTTLMArgin = cTTLMArgin + (CTotalDebit - cTotalCost - cTotalCredit / 2)
                
                CTotalDebit = 0
                cTotalCost = 0
                cTotalPaid = 0
                cTotalCredit = 0
                cComm = 0
                
                lvSysops(0).Refresh
                
            
            End If
        Next
        
        lvSysops(0).Refresh
        gSleep
    Else
    
    End If
    
    
    lblCost.Caption = Format(cTTLCost, "Currency")
    lblMargin.Caption = Format(cTTLMArgin, "Currency")
    lblPaid.Caption = Format(cTTLPaid, "Currency")
    lblDebit.Caption = Format(CTTLDebit, "Currency")
    lblComm.Caption = Format(cTTLCOMM, "Currency")
    lblCredit.Caption = Format(cTTLCredit, "Currency")
    
    Screen.MousePointer = vbDefault
    
    

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

Private Sub Command2_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Command2_Click"
    Const ContainerName = "frmCommissions"
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

Private Sub cmdCalc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdCalc_MouseMove"
    Const ContainerName = "frmCommissions"
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



    Static bPlay As Boolean
    Dim oXY As POINTAPI, L As Long
    
    If bPlay = False And picTSContainer(0).Visible = True Then
            
        L = GetCursorPos(oXY)
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "After you have selected the start and end time for the query on the master database, click this button and the forecaster will display how much commission the sysops in the list have recieved."
        bPlay = True
    
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

Private Sub cmdCreate_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdCreate_Click"
    Const ContainerName = "frmCommissions"
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
    
    cmdCreate.Enabled = False
    
    If Trim(cmbClass.Text) <> "" Then
        On Error GoTo ErrorOccur
        Dim iNum As Long
        Do
            Err.Clear
            iNum = MySQL.GetTMPRecID("commnames", directConn)
            MySQL.Execute directConn, "Insert into commnames (RecID, CommName, VirtualID) Values(" & iNum & ",'" & MySQL.ESC(cmbClass.Text) & "'," & Login.lVirtualID & ")"
            If Err.Number = 0 Then Exit Do
        Loop
            
        cmbClass.AddItem cmbClass.Text
        cmbClass.ItemData(cmbClass.ListCount - 1) = iNum
        cmbClass.ListIndex = cmbClass.ListCount - 1

        lstRates(1).AddItem cmbClass.Text
        lstRates(1).ItemData(lstRates(1).ListCount - 1) = iNum

    End If
    

    cmdCreate.Enabled = True
    
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

Private Sub cmdCreate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdCreate_MouseMove"
    Const ContainerName = "frmCommissions"
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



    Static bPlay As Boolean
    Dim oXY As POINTAPI, L As Long
    
    If bPlay = False And picTSContainer(0).Visible = True Then
            
        L = GetCursorPos(oXY)
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "After you have typed in the new name for the class of commission in the combo box to the left of this create button. This will be the name for this type of commission. It will add it to the list and allow you to select which type of commission you are working with."
        bPlay = True
    
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

Private Sub cmdSave_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdSave_Click"
    Const ContainerName = "frmCommissions"
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
    
    Dim oRate As clsRates
    
    If oComm.colRates.Count > 0 Then
        Screen.MousePointer = vbHourglass
        
        Dim ix As Long
        For ix = 1 To oComm.colRates.Count
            Set oRate = oComm.colRates(ix)
            
            Select Case oRate.Flag
            Case 1
                MySQL.Execute directConn, "Insert Into commrates (ClassID, ptRecID, Min, Max, CommPerUnit, Percentile, VirtualID, sKey) " + _
                                        "VALUES ('" & oRate.ClassID & "','" & oRate.ptRecID & "','" & oRate.Min & "','" & oRate.Max & "','" & oRate.Rate & "','" & oRate.Percentile & "','" & Login.lVirtualID & "','" & oRate.Key & "')"
                oRate.Flag = 0
            Case 2
                MySQL.Execute directConn, "Update commrates Set " + _
                                        "ClassID = " & oRate.ClassID & "," & _
                                        "ptRecID = " & oRate.ptRecID & "," & _
                                        "Min = " & oRate.Min & "," & _
                                        "Max = " & oRate.Max & "," & _
                                        "CommPerUnit = " & oRate.Rate & "," & _
                                        "Percentile = " & oRate.Percentile & " " & _
                                        "WHERE sKey = '" & oRate.Key & "'"
                oRate.Flag = 0
                
            Case Else
                
            End Select
        Next
    
        Screen.MousePointer = vbDefault
        
    End If
    
    
'Exit Sub

  
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

Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdSave_MouseMove"
    Const ContainerName = "frmCommissions"
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


    Static bPlay As Boolean
    Dim oXY As POINTAPI, L As Long
    
    If bPlay = False And picTSContainer(0).Visible = True Then
            
        L = GetCursorPos(oXY)
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "Once you have finished adjusting rates and scales click on this button and the details will be saved to the database."
        bPlay = True
    
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

Private Sub cmdUpdate_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdUpdate_Click"
    Const ContainerName = "frmCommissions"
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

    Dim oRate As clsRates
    
    Set oRate = oComm.colRates(cmdUpdate.Tag)
    oRate.ClassID = cmbClass.ItemData(cmbClass.ListIndex)
    oRate.ptRecID = Val(Mid(lvPlans.SelectedItem.Key, 2))
    oRate.Min = Val(txtQuanity(0))
    oRate.Max = Val(txtQuanity(1))
    oRate.Rate = Val(txtQuanity(2))
    If oRate.Flag = 1 Then oRate.Flag = 1 Else oRate.Flag = 2
    oRate.Percentile = (Val(txtQuanity(2)) * (Val(txtQuanity(1)) - Val(txtQuanity(0)))) / ((Val(txtQuanity(1)) - Val(txtQuanity(0))) * Val(txtMargin.Tag))
    
    Dim itmX As ListItem
    
    Set itmX = lvScale.ListItems(cmdUpdate.Tag)
    itmX.Text = oRate.Min
    itmX.SubItems(1) = oRate.Max
    itmX.SubItems(2) = Format(oRate.Rate, "Currency")
    itmX.SubItems(3) = oRate.Percentile
    itmX.SubItems(4) = Format((oRate.Max - oRate.Min) * oRate.Rate, "currency")
    itmX.Tag = oRate.ptRecID



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

Private Sub cmdUpdate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdUpdate_MouseMove"
    Const ContainerName = "frmCommissions"
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


Static bPlay As Boolean
    Dim oXY As POINTAPI, L As Long
    
    If bPlay = False And picTSContainer(0).Visible = True Then
            
        L = GetCursorPos(oXY)
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "After you have selected a scales node and changed the details in this box, click on the update button."
        bPlay = True
    
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

Private Sub Form_Load()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Load"
    Const ContainerName = "frmCommissions"
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

    
    frmScales.Enabled = IIf(Login.bMaster = True Or Login.bVISPPrimary = True, True, False)
    frmAssign.Enabled = IIf(Login.bMaster = True Or Login.bVISPPrimary = True, True, False)
    
    'If Login.bVISPFiscal = False Then
    '    lvsysops(0).ColumnHeaders.Remove 10
    '    lvsysops(0).ColumnHeaders.Remove 9
    '    lvsysops(0).ColumnHeaders.Remove 8
    '    lvsysops(0).ColumnHeaders.Remove 7
    '    lvsysops(0).ColumnHeaders.Remove 6
    '    lvsysops(1).ColumnHeaders.Remove 3
    'End If
    
   ' Call GUI.LoadColWidths(lvsysops(0), Me)
    'Call GUI.LoadColWidths(lvsysops(1), Me)
    Call GUI.LoadColWidths(lvSysops(2), Me)
    Call GUI.LoadColWidths(lvPlans, Me)
    Call GUI.LoadColWidths(lvScale, Me)

    Call ts_Click
     
    MonthView(0).Value = Format(GetSetting(App.ProductName, "Commissions", "StartDate", Format(DateAdd("m", -1, sysnow), "dd/mm/yyyy")), "ddddd")
    
    MonthView(1).Value = Format(GetSetting(App.ProductName, "Commissions", "StopDate", Format(sysnow, "dd/mm/yyyy")), "ddddd")
    PopulateSysop
    PopulateComm
    
    

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

Private Sub Form_Resize()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Resize"
    Const ContainerName = "frmCommissions"
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

    
    If Me.WindowState <> vbMinimized Then
    
        If Me.ScaleHeight <> 0 Then
        
            ts.Move ts.Left, ts.Top, Me.ScaleWidth - ts.Left * 2, Me.ScaleHeight - ts.Top - ts.Left
            Line1.X1 = 0
            Line1.X2 = Me.ScaleWidth
        End If
        
        Call ts_Click
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

Private Sub Form_Unload(Cancel As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Unload"
    Const ContainerName = "frmCommissions"
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


    'Call GUI.SaveColWidths(lvsysops(0), Me)
    'Call GUI.SaveColWidths(lvsysops(1), Me)
    Call GUI.SaveColWidths(lvSysops(2), Me)
    Call GUI.SaveColWidths(lvPlans, Me)
    Call GUI.SaveColWidths(lvScale, Me)
    
    SaveSetting App.ProductName, "Commissions", "StartDate", MonthView(0).Value
    SaveSetting App.ProductName, "Commissions", "StopDate", MonthView(1).Value
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

Private Sub lblTxt_Change()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lblTxt_Change"
    Const ContainerName = "frmCommissions"
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


    MySQL.Execute directConn, "UPDATE sysops Set CommRate = " & itxM.Tag & ", wrkPercent = " & IIf(itxM.Checked = True, -1, 0) & " where RecID = " & Mid(itxM.Key, 2)
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

Private Sub lstRates_Click(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lstRates_Click"
    Const ContainerName = "frmCommissions"
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


    Dim txt As String
    
        
    If Not itxM Is Nothing Then
        If itxM.Tag <> 0 Then
        End If
        
        If lstRates(1).ListIndex <> -1 Then
            itxM.Tag = lstRates(1).ItemData(lstRates(1).ListIndex)
        End If
                    
        txt = itxM.Text & " is assigned "
        
        If lstRates(1).ListIndex <> -1 Then
            txt = txt + lstRates(1).List(lstRates(1).ListIndex)
        Else
            txt = txt + "nothing so far"
        End If
        
        lblTxt = txt
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

Private Sub lstRates_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lstRates_MouseMove"
    Const ContainerName = "frmCommissions"
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



    Static bPlay As Boolean
    Dim oXY As POINTAPI, L As Long
    
    If bPlay = False And picTSContainer(0).Visible = True Then
            
        L = GetCursorPos(oXY)
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "This is a list of commision classes, select the sysop that you want to assign a commission class to then select the commission class in this listview."
        bPlay = True
    
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

Private Sub lvPlans_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    GUI.ColumnSort ColumnHeader, lvPlans
    
End Sub

Private Sub lvPlans_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvPlans_ItemClick"
    Const ContainerName = "frmCommissions"
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

   
    txtQuanity(2) = ""
    txtPer(0) = ""
    txtPer(1) = ""
'    txtPer(2) = ""
    
    txtQuanity(0).Text = "1"
    txtQuanity(1).Text = "2"
    txtQuanity(2).Text = "0"
    txtRRP.Tag = CCur(Item.SubItems(6))
    txtRRP.Text = Item.SubItems(6)
    txtCost.Tag = CCur(Item.SubItems(1))
    txtCost.Text = ""
    txtMargin.Tag = CCur(Item.SubItems(7))
    txtMargin.Text = ""
    frameCalc.Tag = Item.Key
    frameCalc.Enabled = True
    frameCalc.Caption = "Rates: " & Item.Text
    
    lvScale.ListItems.Clear
    cmdAdd.Enabled = True
    cmdUpdate.Enabled = False
    cmdBonuses.Enabled = True
    cmdBonuses.Tag = Val(Mid(Item.Key, 2))
    
    If oComm.colRates.Count > 0 Then
        Dim ix As Long
        Dim itmX As ListItem
        
        
        
        For ix = 1 To oComm.colRates.Count
            If oComm.colRates(ix).ptRecID = Val(Mid(Item.Key, 2)) And oComm.colRates(ix).ClassID = cmbClass.ItemData(cmbClass.ListIndex) Then
                Set itmX = lvScale.ListItems.Add(, oComm.colRates(ix).Key)
                itmX.Text = oComm.colRates(ix).Min
                itmX.SubItems(1) = oComm.colRates(ix).Max
                itmX.SubItems(2) = Format(oComm.colRates(ix).Rate, "Currency")
                itmX.SubItems(3) = oComm.colRates(ix).Percentile
                itmX.SubItems(4) = Format((oComm.colRates(ix).Max - oComm.colRates(ix).Min) * oComm.colRates(ix).Rate, "currency")
            End If
        Next
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

Private Sub lvPlans_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvPlans_MouseMove"
    Const ContainerName = "frmCommissions"
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


    Static bPlay As Boolean
    Dim oXY As POINTAPI, L As Long
    
    If bPlay = False And picTSContainer(0).Visible = True Then
            
        L = GetCursorPos(oXY)
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "This is the list of services and plan that you can apply individual rate scales against."
        bPlay = True
    
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

Private Sub lvScale_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    GUI.ColumnSort ColumnHeader, lvScale
    
End Sub

Private Sub lvScale_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvScale_ItemClick"
    Const ContainerName = "frmCommissions"
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


    cmdAdd.Enabled = False
    cmdUpdate.Enabled = True
    cmdUpdate.Tag = Item.Key
    
    Dim ix As Long
    For ix = 1 To lvPlans.ListItems.Count
        If lvPlans.ListItems(ix).Key = "r" & oComm.colRates(Item.Key).ptRecID Then
            lvPlans.ListItems(ix).Selected = True
            Exit For
        End If
    Next
    
    txtQuanity(0) = Item.Text
    txtQuanity(1) = Item.SubItems(1)
    txtQuanity(2) = CCur(Item.SubItems(2))
    
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

Private Sub lvScale_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvScale_MouseMove"
    Const ContainerName = "frmCommissions"
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


    Static bPlay As Boolean
    Dim oXY As POINTAPI, L As Long
    
    If bPlay = False And picTSContainer(0).Visible = True Then
            
        L = GetCursorPos(oXY)
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "Here is the list of scales that have been applied to the individual item that you have selected. To change a scale you have already added; select it then change the details and press the green update button."
        bPlay = True
    
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

Private Sub lvSysops_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    GUI.ColumnSort ColumnHeader, lvSysops(Index)
    
End Sub

Private Sub lvsysops_ItemCheck(Index As Integer, ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvsysops_ItemCheck"
    Const ContainerName = "frmCommissions"
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
    Case 1
    Case 2
        
        Set itxM = Item
        Set lvSysops(2).SelectedItem = Item
        chkPercentile.Value = -Item.Checked
        
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

Private Sub lvsysops_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvsysops_ItemClick"
    Const ContainerName = "frmCommissions"
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
    Case 1
    Case 2
        
        Set itxM = Item
        
        Dim txt As String
        
            
        If Not lvSysops(2).SelectedItem Is Nothing Then
            If lvSysops(2).SelectedItem.Tag <> 0 Then
                Dim ix As Integer
                
                For ix = 0 To lstRates(1).ListCount - 1
                    If lstRates(1).ItemData(ix) = lvSysops(2).SelectedItem.Tag Then
                        lstRates(1).ListIndex = ix
                        Exit For
                    End If
                Next
            Else
                lstRates(1).ListIndex = -1
            End If
            
            If lstRates(1).ListIndex <> -1 Then
                lvSysops(2).SelectedItem.Tag = lstRates(1).ItemData(lstRates(1).ListIndex)
            End If
            
        End If
        
               
        If lvSysops(2).SelectedItem Is Nothing Then Exit Sub
        
        txt = lvSysops(2).SelectedItem.Text & " is assigned "
        
        If lstRates(1).ListIndex <> -1 Then
            txt = txt + lstRates(1).List(lstRates(1).ListIndex)
        Else
            txt = txt + "nothing so far"
        End If
        
        lblTxt = txt
        chkPercentile.Value = -Item.Checked
        
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

Private Sub lvsysops_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvsysops_MouseDown"
    Const ContainerName = "frmCommissions"
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
    Case 1
    
        If Button = 2 And Login.bVISPPrimary = True Then
            If lvSysops(1).SelectedItem Is Nothing Then
            
            Else
                PopupMenu mnuPopups
            End If
        End If
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

Private Sub lvSysops_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvSysops_MouseMove"
    Const ContainerName = "frmCommissions"
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


    Static bPlay As Boolean
    Dim oXY As POINTAPI, L As Long
    
    If bPlay = False And picTSContainer(0).Visible = True Then
            
        L = GetCursorPos(oXY)
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "Here is the list of sysops in your organisation that you can adjust and set a commission classes to, if the sysop has a check next to his/her's name then it is calculated on percentiles, this sometime can generate more money but if sales are being marked down it will generate less money."
        bPlay = True
    
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

Private Sub picTSContainer_Resize(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "picTSContainer_Resize"
    Const ContainerName = "frmCommissions"
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
    
    Case 1
        lvSysops(0).Move lvSysops(0).Left, lvSysops(0).Top, picTSContainer(Index).ScaleWidth - lvSysops(0).Left * 2, picTSContainer(Index).ScaleHeight - lvSysops(0).Top * 3 - Frame6.Height
        Frame6.Move lvSysops(0).Left, lvSysops(0).Top * 2 + lvSysops(0).Height, lvSysops(0).Width
        cmbExport.Left = Frame6.Width - cmbExport.Width - 60
        
    Case 2
        lvSysops(1).Move lvSysops(1).Left, lvSysops(1).Top, picTSContainer(Index).ScaleWidth - lvSysops(1).Left * 2, picTSContainer(Index).ScaleHeight - lvSysops(1).Top * 2
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

Private Sub ts_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "ts_Click"
    Const ContainerName = "frmCommissions"
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


    Dim X As Integer
    
    For X = picTSContainer.LBound To picTSContainer.UBound
        If ts.SelectedItem.Index - 1 <> X Then picTSContainer(X).Visible = False
    Next
    
    If ts.SelectedItem.Index - 1 <= picTSContainer.UBound Then
        picTSContainer(ts.SelectedItem.Index - 1).Move ts.ClientLeft, ts.ClientTop, ts.ClientWidth, ts.ClientHeight
        picTSContainer(ts.SelectedItem.Index - 1).Visible = True
        picTSContainer(ts.SelectedItem.Index - 1).ZOrder 0
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

Public Function PopulateSysop()


    '*[ Error Checking Variables ]**********************************************************************************
    
    Const RoutineName = "PopulateSysop"
    Const ContainerName = "frmCommissions"
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

    Dim bResult As Boolean
    Dim rsSysops As adodb.Recordset
    
    bResult = MySQL.OpenTable(directConn, rsSysops, , MySQL.virtualisp("select distinct * ,virtualisp.description as ViSPDesc from sysops ", "sysops", True) & " group by sysops.RecID")
    
    If rsSysops.RecordCount > 0 Then
        Dim itmX As ListItem
        While Not rsSysops.EOF And Err.Number = 0
        
            
            Set itmX = lvSysops(0).ListItems.Add(, "s" & rsSysops!RecID, rsSysops!Username)
            itmX.Tag = "" & rsSysops!CommRate & "_" & rsSysops!PerVISP
            itmX.SubItems(1) = IIf(IsNull(rsSysops!ViSPDesc), "", rsSysops!ViSPDesc)
            
            Set itmX = lvSysops(2).ListItems.Add(, "r" & rsSysops!RecID, rsSysops!Username)
            itmX.SubItems(1) = rsSysops!SecurityLevel
            itmX.SubItems(2) = rsSysops!Description
            itmX.SubItems(3) = MySQL.CodeME(rsSysops!ViSPDesc)
            
            itmX.Tag = Val(rsSysops!CommRate)
            itmX.Checked = IIf(IIf(IsNull(rsSysops!wrkPercent), 0, Val(wrkPercent)) <> 0, True, False)
            
            Dim imgX As Byte
            Dim imgMin As Byte
            Dim imgMax As Byte
        
            For imgX = 1 To ilSysops.ListImages.Count
                If InStr(ilSysops.ListImages(imgX).Key, "_") > 0 Then
                    imgMin = CByte(Mid(ilSysops.ListImages(imgX).Key, 2, 3))
                    imgMax = CByte(Mid(ilSysops.ListImages(imgX).Key, 6, 3))
                Else
                    imgMin = CByte(Mid(ilSysops.ListImages(imgX).Key, 2, 3))
                    imgMax = imgMin
                End If
                If CByte(rsSysops!SecurityLevel) >= imgMin And CByte(rsSysops!SecurityLevel) <= imgMax Then
                    itmX.SmallIcon = imgX
                End If
            Next imgX
            
            rsSysops.MoveNext
            
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

Public Sub PopulateComm()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "PopulateComm"
    Const ContainerName = "frmCommissions"
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

    Dim rsload As adodb.Recordset
    Dim itmX As ListItem
    
    
    If MySQL.OpenTable(directConn, rsload, , MySQL.virtualisp("select distinct vendors.vName ,virtualisp.Description as vispDesc,plantypes.Description ,plantypes.PeriodFee ,plantypes.RecID ,plantemplates.PeriodFee as Cost from plantypes, plantemplates, vendors where plantypes.TemplateID = plantemplates.RecID and plantypes.VendorID = vendors.RecID and plantypes.VirtualID = virtualisp.RecID", "plantypes", False, False)) = True Then
        If rsload.RecordCount > 0 Then
            While Not rsload.EOF And Err.Number = 0
            
                Set itmX = lvPlans.ListItems.Add(, "r" & rsload!RecID, rsload!Description)
                itmX.SubItems(1) = Format(rsload!Cost, "Currency")
                itmX.SubItems(2) = Format(rsload!Cost + rsload!Cost * oTax(Login.TaxCode, Login.TaxCountry), "Currency")
                itmX.SubItems(3) = MySQL.CodeME(rsload!vName)
                itmX.SubItems(4) = MySQL.CodeME(rsload!ViSPDesc)
                itmX.SubItems(5) = Format(rsload!PeriodFee, "Currency")
                itmX.SubItems(6) = Format(rsload!PeriodFee * oTax(Login.TaxCode, Login.TaxCountry) + rsload!PeriodFee, "Currency")
                itmX.Tag = rsload!PeriodFee
                itmX.SubItems(7) = Format((rsload!PeriodFee) - rsload!Cost, "Currency")
                rsload.MoveNext
            Wend
        
        End If
    
    End If
    
    If MySQL.OpenTable(directConn, rsload, , MySQL.virtualisp("select distinct commnames.* from commnames,", "commnames", False, False)) = True Then
        
        If rsload.RecordCount > 0 Then
            While Not rsload.EOF And Err.Number = 0
                cmbClass.AddItem rsload!CommName
                cmbClass.ItemData(cmbClass.ListCount - 1) = rsload!RecID
                lstRates(1).AddItem rsload!CommName
                lstRates(1).ItemData(lstRates(1).ListCount - 1) = IIf(IsNull(rsload!RecID), 0, rsload!RecID)
                rsload.MoveNext
            Wend
            cmbClass.ListIndex = 0
        End If
    End If
    
    If MySQL.OpenTable(directConn, rsload, , MySQL.virtualisp("select distinct commrates.* from commrates,", "commrates", False, False)) = True Then
        oComm.colRates.Clear
        If rsload.RecordCount > 0 Then
            While Not rsload.EOF And Err.Number = 0
                oComm.colRates.Add rsload!sKey, rsload!ClassID, rsload!ptRecID, rsload!Min, rsload!Max, rsload!CommPerUnit, rsload!Percentile, rsload!sKey
                rsload.MoveNext
            Wend
        End If
    End If
    
    
    

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

Public Sub DoCalc()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "DoCalc"
    Const ContainerName = "frmCommissions"
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


    On Error GoTo ErrorOccur
    
    Dim iRange As Long
    
    iRange = Val(txtQuanity(1)) - Val(txtQuanity(0))
    txtCost = Format(iRange * Val(txtCost.Tag), "Currency")
    txtMargin = Format((iRange * Val(txtMargin.Tag)), "Currency")
    
    If Val(txtMargin.Tag) = 0 Then
        Exit Sub
    End If
    
    If txtQuanity(2) <> "" Then
        txtPer(0) = Format((Val(txtQuanity(2)) * iRange) / (iRange * Val(txtMargin.Tag)), "###,###.####%")
        txtPer(1) = Format((Val(txtQuanity(2)) * iRange), "Currency")
    End If
    
       
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

Private Sub txtCost_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtCost_MouseMove"
    Const ContainerName = "frmCommissions"
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


    Static bPlay As Boolean
    Dim oXY As POINTAPI, L As Long
    
    If bPlay = False And picTSContainer(0).Visible = True Then
            
        L = GetCursorPos(oXY)
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "The cost price of the service or plan selected, this is a static field it cannot be changed."
        bPlay = True
    
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

Private Sub txtMargin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtMargin_MouseMove"
    Const ContainerName = "frmCommissions"
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


    Static bPlay As Boolean
    Dim oXY As POINTAPI, L As Long
    
    If bPlay = False And picTSContainer(0).Visible = True Then
            
        L = GetCursorPos(oXY)
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "The marins in the sale of these items based on the default pricing of the service or plan selected, this is a static field it cannot be changed."
        bPlay = True
    
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

Private Sub txtPer_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtPer_MouseMove"
    Const ContainerName = "frmCommissions"
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



    Static bPlay1 As Boolean
    Static bPlay2 As Boolean
    Static bPlay3 As Boolean
    Dim oXY As POINTAPI, L As Long
    
    If picTSContainer(0).Visible = True Then

        Select Case Index
        Case 0
            If bPlay1 = True Then Exit Sub
            
        L = GetCursorPos(oXY)
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y
        bPlay1 = True
            If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "This is the total commission percentage of that will be award to the sysop for the money generated in upkeep of the system's database."
            
        Case 1
                        
               If bPlay2 = True Then Exit Sub
               
        L = GetCursorPos(oXY)
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y
        bPlay2 = True
            If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "This is the total amount awarded in the calculations of commission from the data collected during upkeep."
            
        Case 2
            
            If bPlay3 = True Then Exit Sub
                    
        L = GetCursorPos(oXY)
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y
        bPlay3 = True
            If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "How much per item of service or plan to award to the sysop on this commision class, from the minimum to maximum count range."
            
        End Select
        bPlay = True
    
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

Private Sub txtQuanity_Change(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtQuanity_Change"
    Const ContainerName = "frmCommissions"
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


    DoCalc
    
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

Private Sub txtQuanity_KeyPress(Index As Integer, KeyAscii As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtQuanity_KeyPress"
    Const ContainerName = "frmCommissions"
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
    Case 8
    Case 48 To 57
    Case Asc(".")
        If Index = 2 Then
            If InStr(txtQuanity(Index).Text, ".") > 0 Then KeyAscii = 0
        Else
            KeyAscii = 0
        End If
        
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

Private Sub txtQuanity_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtQuanity_MouseMove"
    Const ContainerName = "frmCommissions"
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


    Static bPlay1 As Boolean
    Static bPlay2 As Boolean
    Static bPlay3 As Boolean
    Dim oXY As POINTAPI, L As Long
    
    If picTSContainer(0).Visible = True Then
            
        Select Case Index
        Case 0
        
            If bPlay1 = True Then Exit Sub
            
            bPlay1 = True
            L = GetCursorPos(oXY)
            If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
            If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y
            If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "The lowest count to start from when calculating commissionable amounts."
            
        Case 1
            
            If bPlay2 = True Then Exit Sub
            bPlay2 = True
            L = GetCursorPos(oXY)
            If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
            If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y
            
            If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "The highest count to end from when calculating commissionable amounts."
            
        Case 2
            
            If bPlay3 = True Then Exit Sub
            bPlay3 = True
            L = GetCursorPos(oXY)
            If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
            If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y
            If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "How much per item of service or plan to award to the sysop on this commision class, from the minimum to maximum count range."
            
        End Select
       
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

Private Sub txtRRP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtRRP_MouseMove"
    Const ContainerName = "frmCommissions"
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


    Static bPlay As Boolean
    Dim oXY As POINTAPI, L As Long
    
    If bPlay = False And picTSContainer(0).Visible = True Then
            
        L = GetCursorPos(oXY)
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "The recommended retail price of the service or plan selected, this is a static field it cannot be changed."
        bPlay = True
    
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
