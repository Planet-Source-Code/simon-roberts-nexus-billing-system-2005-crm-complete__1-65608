VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVISP 
   BackColor       =   &H00000000&
   Caption         =   "Virtual ISP Configuration"
   ClientHeight    =   10170
   ClientLeft      =   1515
   ClientTop       =   2685
   ClientWidth     =   12120
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmResellers.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10170
   ScaleWidth      =   12120
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4620
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResellers.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResellers.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResellers.frx":0BAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResellers.frx":1000
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResellers.frx":1452
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResellers.frx":190C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResellers.frx":1D5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResellers.frx":21B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResellers.frx":2602
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResellers.frx":2A54
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picResize 
      BackColor       =   &H8000000D&
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
      Height          =   8175
      Left            =   3870
      MousePointer    =   9  'Size W E
      ScaleHeight     =   8175
      ScaleWidth      =   75
      TabIndex        =   2
      Top             =   900
      Width           =   75
   End
   Begin VB.Timer mTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   10320
      Top             =   30
   End
   Begin VB.PictureBox picListview 
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
      Height          =   7815
      Left            =   60
      ScaleHeight     =   7815
      ScaleWidth      =   3825
      TabIndex        =   0
      Top             =   870
      Width           =   3825
      Begin MSComctlLib.ProgressBar pb 
         Height          =   225
         Left            =   60
         TabIndex        =   58
         Top             =   7530
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
      End
      Begin MSComctlLib.ListView lvVISP 
         Height          =   6885
         Left            =   60
         TabIndex        =   1
         Top             =   600
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   12144
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   16777215
         BackColor       =   8421504
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "!Description!"
            Text            =   "Description"
            Object.Width           =   4586
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "!ABN!"
            Text            =   "ABN"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "!ACN!"
            Text            =   "ACN"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "!NoSub!"
            Text            =   "No. Customers"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "!Subscribed!"
            Text            =   "Subscribed"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "!LogoURL!"
            Text            =   "Logo URL"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Object.Tag             =   "!RecID!^select count(*) as nResult from sysops where VirtualID =^"
            Text            =   "Next Cycle Date"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Object.Tag             =   "!RecID!^select count(*) as nResult from plantypes where VirtualID = ^"
            Text            =   "Roaming Intranet Hostname"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.Label lblArticles 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   60
         TabIndex        =   60
         Top             =   330
         Width           =   60
      End
      Begin VB.Label lblArticles 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   60
         TabIndex        =   59
         Top             =   30
         Width           =   60
      End
   End
   Begin VB.PictureBox picTab 
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
      Height          =   9015
      Left            =   3960
      ScaleHeight     =   9015
      ScaleWidth      =   8115
      TabIndex        =   3
      Top             =   840
      Width           =   8115
      Begin VB.PictureBox tsContainer 
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
         Height          =   6255
         Index           =   3
         Left            =   -570
         ScaleHeight     =   6255
         ScaleWidth      =   6765
         TabIndex        =   34
         Top             =   4200
         Visible         =   0   'False
         Width           =   6765
         Begin VB.CommandButton cmdAddAddress 
            Caption         =   "&Add Address"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   90
            TabIndex        =   35
            Top             =   2910
            Width           =   2385
         End
         Begin MSComctlLib.ListView lvAddresses 
            Height          =   2685
            Left            =   60
            TabIndex        =   36
            Top             =   60
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   4736
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            HoverSelection  =   -1  'True
            PictureAlignment=   1
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Contact Name"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Street 1"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Street 2"
               Object.Width           =   2470
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Suburb"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "State"
               Object.Width           =   1589
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Postcode"
               Object.Width           =   1588
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Country"
               Object.Width           =   2540
            EndProperty
            Picture         =   "frmResellers.frx":2EA6
         End
      End
      Begin VB.PictureBox tsContainer 
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
         Height          =   1305
         Index           =   5
         Left            =   5100
         ScaleHeight     =   1305
         ScaleWidth      =   2595
         TabIndex        =   43
         Top             =   7260
         Visible         =   0   'False
         Width           =   2595
         Begin VB.Frame fraFileDB 
            Caption         =   "File Database Resource"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6465
            Left            =   120
            TabIndex        =   53
            Top             =   1830
            Width           =   7545
            Begin VB.Frame Frame9 
               Caption         =   "Proxy Server (server[:port])"
               Height          =   795
               Left            =   4080
               TabIndex        =   87
               Top             =   1530
               Width           =   3225
               Begin VB.TextBox txtftpProxy 
                  Height          =   360
                  Left            =   120
                  TabIndex        =   88
                  Top             =   270
                  Width           =   2955
               End
            End
            Begin VB.OptionButton optFileDBMode 
               Caption         =   "HTTP Services"
               Enabled         =   0   'False
               Height          =   375
               Index           =   2
               Left            =   5070
               Style           =   1  'Graphical
               TabIndex        =   86
               Top             =   270
               Width           =   2235
            End
            Begin VB.Frame frmFileDBCosts 
               Caption         =   "Data Housing Costs and Charges (Prices are Ex GST)"
               Height          =   1905
               Left            =   240
               TabIndex        =   75
               Top             =   4380
               Width           =   7065
               Begin VB.Frame Frame8 
                  Caption         =   "Statistics - Files/Folders"
                  Height          =   1035
                  Left            =   150
                  TabIndex        =   83
                  Top             =   750
                  Width           =   3015
                  Begin VB.Label lblFTPStats 
                     Alignment       =   2  'Center
                     Caption         =   "You have a total of 0 Folders and 0 Files."
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   495
                     Left            =   120
                     TabIndex        =   84
                     Top             =   390
                     Width           =   2805
                  End
               End
               Begin VB.Frame Frame7 
                  BackColor       =   &H00404040&
                  Caption         =   "Current Cycle Rates/Fees"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   1485
                  Left            =   3240
                  TabIndex        =   78
                  Top             =   300
                  Width           =   3705
                  Begin VB.Label lblFileDB_CycleIntervals 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "Charged on a Monthly Bases, at 1 month intervals."
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   465
                     Left            =   120
                     TabIndex        =   81
                     Top             =   960
                     Width           =   3465
                  End
                  Begin VB.Label lblFileDB_TotalCost 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "$ 0.00"
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
                     Height          =   375
                     Left            =   120
                     TabIndex        =   80
                     Top             =   600
                     Width           =   3465
                  End
                  Begin VB.Label lblFileDB_TotalMBs 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "0 Mb's"
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
                     Height          =   375
                     Left            =   120
                     TabIndex        =   79
                     Top             =   270
                     Width           =   3465
                  End
               End
               Begin VB.TextBox txtFileDB_CostPerMB 
                  Height          =   360
                  Left            =   1590
                  TabIndex        =   76
                  Top             =   270
                  Width           =   1575
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  Caption         =   "Cost Per MB: $"
                  Height          =   240
                  Left            =   150
                  TabIndex        =   77
                  Top             =   390
                  Width           =   1350
               End
            End
            Begin VB.Frame Frame6 
               Caption         =   "Authentication Settings with Remote Server"
               Height          =   2025
               Left            =   240
               TabIndex        =   65
               Top             =   2340
               Width           =   7065
               Begin VB.CommandButton Command1 
                  Caption         =   "Archive to Local Machine"
                  Enabled         =   0   'False
                  Height          =   345
                  Left            =   4080
                  TabIndex        =   85
                  Top             =   1560
                  Width           =   2865
               End
               Begin VB.CommandButton cmdCALCDataFee 
                  Caption         =   "&Calculate Current Data Fees"
                  Enabled         =   0   'False
                  Height          =   345
                  Left            =   4080
                  TabIndex        =   82
                  Top             =   1170
                  Width           =   2865
               End
               Begin VB.CommandButton cmdTestWAN 
                  Caption         =   "&Test Connection"
                  Height          =   345
                  Left            =   4080
                  TabIndex        =   74
                  Top             =   780
                  Width           =   2865
               End
               Begin VB.CheckBox chkPingAlive 
                  Caption         =   "Keep link alive by pinging ocassionally"
                  Height          =   285
                  Left            =   240
                  TabIndex        =   73
                  Top             =   1590
                  Value           =   1  'Checked
                  Width           =   3735
               End
               Begin VB.CheckBox chkIEProxy 
                  Caption         =   "Use IE Proxy Settings for FTP."
                  Height          =   285
                  Left            =   240
                  TabIndex        =   72
                  Top             =   1260
                  Value           =   1  'Checked
                  Width           =   3435
               End
               Begin VB.TextBox txtFileDB_Port 
                  Alignment       =   2  'Center
                  Height          =   360
                  IMEMode         =   3  'DISABLE
                  Left            =   1290
                  TabIndex        =   70
                  Text            =   "21"
                  Top             =   780
                  Width           =   2325
               End
               Begin VB.TextBox txtFileDB_Password 
                  Alignment       =   2  'Center
                  Height          =   360
                  IMEMode         =   3  'DISABLE
                  Left            =   4830
                  PasswordChar    =   "»"
                  TabIndex        =   68
                  Top             =   330
                  Width           =   2085
               End
               Begin VB.TextBox txtFileDB_Username 
                  Alignment       =   2  'Center
                  Height          =   360
                  Left            =   1290
                  TabIndex        =   66
                  Top             =   330
                  Width           =   2325
               End
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  Caption         =   "Port:"
                  Height          =   240
                  Left            =   210
                  TabIndex        =   71
                  Top             =   840
                  Width           =   420
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  Caption         =   "Password:"
                  Height          =   240
                  Left            =   3870
                  TabIndex        =   69
                  Top             =   390
                  Width           =   915
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Username:"
                  Height          =   240
                  Left            =   210
                  TabIndex        =   67
                  Top             =   390
                  Width           =   945
               End
            End
            Begin VB.OptionButton optFileDBMode 
               Caption         =   "FTP Services"
               Height          =   375
               Index           =   1
               Left            =   2760
               Style           =   1  'Graphical
               TabIndex        =   63
               Top             =   270
               Width           =   2235
            End
            Begin VB.OptionButton optFileDBMode 
               Caption         =   "File && Printer Sharing"
               Height          =   375
               Index           =   0
               Left            =   240
               Style           =   1  'Graphical
               TabIndex        =   62
               Top             =   270
               Value           =   -1  'True
               Width           =   2445
            End
            Begin VB.Frame Frame3 
               Caption         =   "HTTP URL for Web Services"
               Height          =   795
               Left            =   240
               TabIndex        =   61
               Top             =   1500
               Width           =   3735
               Begin VB.TextBox txtFileDB_URL 
                  Height          =   360
                  Left            =   120
                  TabIndex        =   64
                  Text            =   "http://"
                  Top             =   270
                  Width           =   3465
               End
            End
            Begin VB.Frame Frame5 
               Caption         =   "Base UNC/FTP Path (without trailing slashes)"
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
               Left            =   3750
               TabIndex        =   56
               Top             =   720
               Width           =   3555
               Begin VB.TextBox txtVISPFTP 
                  Alignment       =   2  'Center
                  Height          =   345
                  Index           =   1
                  Left            =   150
                  MaxLength       =   255
                  TabIndex        =   57
                  Top             =   270
                  Width           =   3255
               End
            End
            Begin VB.Frame Frame4 
               Caption         =   "Server Hostname/IP"
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
               Left            =   240
               TabIndex        =   54
               Top             =   720
               Width           =   3345
               Begin VB.TextBox txtVISPFTP 
                  Alignment       =   2  'Center
                  Height          =   345
                  Index           =   0
                  Left            =   150
                  MaxLength       =   255
                  TabIndex        =   55
                  Top             =   270
                  Width           =   2985
               End
            End
         End
         Begin VB.Frame frmTax 
            Caption         =   "Tax Code and Tax Settings"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1665
            Left            =   120
            TabIndex        =   44
            Top             =   90
            Width           =   7545
            Begin VB.OptionButton optTax 
               Caption         =   "This ViSP Pay's Tax"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Index           =   0
               Left            =   210
               Style           =   1  'Graphical
               TabIndex        =   51
               Top             =   270
               Width           =   2235
            End
            Begin VB.Frame frmTaxCode 
               Caption         =   "Tax Code && Country Code"
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
               Left            =   2790
               TabIndex        =   48
               Top             =   180
               Width           =   4605
               Begin VB.TextBox txtTax 
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
                  Height          =   315
                  Index           =   0
                  Left            =   150
                  TabIndex        =   50
                  Text            =   "GST"
                  Top             =   240
                  Width           =   2085
               End
               Begin VB.TextBox txtTax 
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
                  Height          =   315
                  Index           =   1
                  Left            =   2310
                  TabIndex        =   49
                  Text            =   "AUS0001"
                  Top             =   240
                  Width           =   2175
               End
            End
            Begin VB.OptionButton optTax 
               Caption         =   "This ViSP is from overseas or is Tax Exempt in Australia"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Index           =   1
               Left            =   210
               Style           =   1  'Graphical
               TabIndex        =   47
               Top             =   960
               Width           =   2235
            End
            Begin VB.Frame frmTaxCode 
               Caption         =   "Tax Exemption Number"
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
               Index           =   1
               Left            =   2790
               TabIndex        =   45
               Top             =   900
               Width           =   4575
               Begin VB.TextBox txtTax 
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
                  Height          =   315
                  Index           =   2
                  Left            =   150
                  TabIndex        =   46
                  Top             =   240
                  Width           =   4275
               End
            End
         End
      End
      Begin VB.PictureBox tsContainer 
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
         Height          =   5715
         Index           =   1
         Left            =   1260
         ScaleHeight     =   5715
         ScaleWidth      =   4275
         TabIndex        =   31
         Top             =   1500
         Visible         =   0   'False
         Width           =   4275
         Begin VB.CommandButton cmdAddPhone 
            Caption         =   "&Add Phone"
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
            Left            =   540
            TabIndex        =   32
            Top             =   5340
            Width           =   2145
         End
         Begin MSComctlLib.ListView lvPhone 
            Height          =   4995
            Left            =   540
            TabIndex        =   33
            Top             =   150
            Width           =   6225
            _ExtentX        =   10980
            _ExtentY        =   8811
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            HoverSelection  =   -1  'True
            PictureAlignment=   1
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Contact Name"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Phone Number"
               Object.Width           =   7057
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Extension"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Position"
               Object.Width           =   2540
            EndProperty
            Picture         =   "frmResellers.frx":3B29
         End
      End
      Begin VB.PictureBox tsContainer 
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
         Height          =   6825
         Index           =   2
         Left            =   750
         ScaleHeight     =   6825
         ScaleWidth      =   7695
         TabIndex        =   28
         Top             =   1890
         Visible         =   0   'False
         Width           =   7695
         Begin VB.CommandButton cmdAddEmail 
            Caption         =   "&Add e-Mail"
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
            Left            =   60
            TabIndex        =   29
            Top             =   1950
            Width           =   2115
         End
         Begin MSComctlLib.ListView lvEmail 
            Height          =   1785
            Left            =   60
            TabIndex        =   30
            Top             =   90
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   3149
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            HoverSelection  =   -1  'True
            PictureAlignment=   1
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Contact Name"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "eMail Address"
               Object.Width           =   8819
            EndProperty
            Picture         =   "frmResellers.frx":4208
         End
      End
      Begin VB.PictureBox tsContainer 
         BackColor       =   &H00004080&
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
         Height          =   8235
         Index           =   0
         Left            =   120
         ScaleHeight     =   8235
         ScaleWidth      =   7785
         TabIndex        =   4
         Top             =   870
         Width           =   7785
         Begin VB.Timer tmrIcons 
            Interval        =   250
            Left            =   2730
            Top             =   120
         End
         Begin VB.Frame frame 
            BackColor       =   &H0084E8E8&
            Caption         =   "Joining Fee (Fee or barter value of sign up to ViSP Networks)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   915
            Index           =   6
            Left            =   150
            TabIndex        =   27
            ToolTipText     =   $"frmResellers.frx":492B
            Top             =   7740
            Width           =   7545
            Begin VB.TextBox txtField 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404080&
               Height          =   450
               Index           =   8
               Left            =   150
               MaxLength       =   50
               TabIndex        =   21
               Top             =   300
               Width           =   7275
            End
         End
         Begin VB.CommandButton cmdSave 
            BackColor       =   &H0084E8E8&
            Caption         =   "&Save Reseller"
            Enabled         =   0   'False
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
            Left            =   5220
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   90
            Width           =   2505
         End
         Begin VB.CommandButton cmdSubscribe 
            Caption         =   "&Add More Subscribed Accounts"
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
            Left            =   5130
            TabIndex        =   18
            Top             =   7620
            Width           =   2535
         End
         Begin VB.CommandButton cmdCreate 
            BackColor       =   &H0084E8E8&
            Caption         =   "&Create Reseller"
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
            TabIndex        =   17
            Top             =   90
            Width           =   2535
         End
         Begin VB.Frame frame 
            BackColor       =   &H00004080&
            Caption         =   "Primary Sysop Account"
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
            Height          =   2775
            Index           =   4
            Left            =   150
            TabIndex        =   13
            Top             =   4020
            Width           =   7545
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
               Height          =   360
               IMEMode         =   3  'DISABLE
               Index           =   6
               Left            =   150
               MaxLength       =   50
               PasswordChar    =   "š"
               TabIndex        =   16
               ToolTipText     =   "Enter the ViSP First Sysop account, there primary account password"
               Top             =   2100
               Width           =   7275
            End
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
               Height          =   390
               Index           =   5
               Left            =   150
               MaxLength       =   50
               TabIndex        =   15
               ToolTipText     =   "Username"
               Top             =   1380
               Width           =   7275
            End
            Begin VB.TextBox txtField 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   705
               Index           =   4
               Left            =   150
               MaxLength       =   255
               MultiLine       =   -1  'True
               TabIndex        =   14
               ToolTipText     =   "Description"
               Top             =   300
               Width           =   7275
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
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
               Height          =   240
               Index           =   2
               Left            =   150
               TabIndex        =   26
               Top             =   2430
               Width           =   7230
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
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
               Height          =   240
               Index           =   1
               Left            =   150
               TabIndex        =   25
               Top             =   1740
               Width           =   7230
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Description"
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
               Height          =   240
               Index           =   0
               Left            =   150
               TabIndex        =   24
               Top             =   1020
               Width           =   7230
            End
         End
         Begin VB.Frame frame 
            BackColor       =   &H00004080&
            Caption         =   "ACN/RBN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   825
            Index           =   3
            Left            =   180
            TabIndex        =   11
            Top             =   3180
            Width           =   7515
            Begin VB.TextBox txtField 
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
               Height          =   420
               Index           =   3
               Left            =   150
               MaxLength       =   50
               TabIndex        =   12
               Top             =   270
               Width           =   7275
            End
         End
         Begin VB.Frame frame 
            BackColor       =   &H00004080&
            Caption         =   "ABN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   825
            Index           =   2
            Left            =   180
            TabIndex        =   9
            Top             =   2340
            Width           =   7515
            Begin VB.TextBox txtField 
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
               Height          =   420
               Index           =   2
               Left            =   150
               MaxLength       =   50
               TabIndex        =   10
               Top             =   300
               Width           =   7275
            End
         End
         Begin VB.Frame frame 
            BackColor       =   &H00004080&
            Caption         =   "Domain (This realm is wtihout www. [i.e. projectalpha.com.au])"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   825
            Index           =   1
            Left            =   180
            TabIndex        =   7
            Top             =   1500
            Width           =   7515
            Begin VB.TextBox txtField 
               Alignment       =   2  'Center
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
               Index           =   1
               Left            =   150
               MaxLength       =   255
               TabIndex        =   8
               Top             =   300
               Width           =   7275
            End
         End
         Begin VB.Frame frame 
            BackColor       =   &H00004080&
            Caption         =   "Company Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   825
            Index           =   0
            Left            =   180
            TabIndex        =   5
            Top             =   660
            Width           =   7515
            Begin VB.TextBox txtField 
               Alignment       =   2  'Center
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
               Index           =   0
               Left            =   150
               MaxLength       =   100
               TabIndex        =   6
               Top             =   300
               Width           =   7275
            End
         End
         Begin VB.Frame frame 
            BackColor       =   &H00004080&
            Caption         =   "Number of User Subscribed"
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
            Height          =   825
            Index           =   5
            Left            =   150
            TabIndex        =   19
            Top             =   6840
            Width           =   7545
            Begin VB.TextBox txtField 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00BA3F3F&
               Height          =   420
               Index           =   7
               Left            =   150
               MaxLength       =   50
               TabIndex        =   20
               Text            =   "100"
               Top             =   270
               Width           =   7275
            End
         End
      End
      Begin MSComctlLib.TabStrip ts 
         Height          =   7725
         Left            =   90
         TabIndex        =   37
         Top             =   210
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   13626
         MultiRow        =   -1  'True
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   7
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Company Details"
               Object.ToolTipText     =   "Here is where you set the domain, RBN, ABN, ACN and other relevant business details. "
               ImageVarType    =   2
               ImageIndex      =   4
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Phone Contacts"
               Object.ToolTipText     =   $"frmResellers.frx":49C6
               ImageVarType    =   2
               ImageIndex      =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "e-Mail"
               Object.ToolTipText     =   "This is where you can set all the email address associated with this reseller."
               ImageVarType    =   2
               ImageIndex      =   3
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Address Information"
               Object.ToolTipText     =   "This is where you will store all the Business Addresses and Postal Addresses for this reseller."
               ImageVarType    =   2
               ImageIndex      =   1
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Web Description/Profile"
               Object.ToolTipText     =   $"frmResellers.frx":4A7C
               ImageVarType    =   2
               ImageIndex      =   6
            EndProperty
            BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Constraints"
               Object.ToolTipText     =   "Here is where you can adjust configuration settings for this visp."
               ImageVarType    =   2
               ImageIndex      =   7
            EndProperty
            BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Reseller File Database"
               Object.ToolTipText     =   $"frmResellers.frx":4B1E
               ImageVarType    =   2
               ImageIndex      =   10
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.PictureBox tsContainer 
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
         Height          =   7905
         Index           =   4
         Left            =   180
         ScaleHeight     =   7905
         ScaleWidth      =   7785
         TabIndex        =   38
         Top             =   720
         Visible         =   0   'False
         Width           =   7785
         Begin VB.Frame Frame2 
            Caption         =   "Logo / Emblem URL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1125
            Left            =   150
            TabIndex        =   41
            Top             =   6180
            Width           =   7545
            Begin VB.TextBox txtLogoURL 
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "Terminal"
                  Size            =   12
                  Charset         =   255
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0043D143&
               Height          =   555
               Left            =   210
               MultiLine       =   -1  'True
               TabIndex        =   42
               Text            =   "frmResellers.frx":4C07
               Top             =   510
               Width           =   7125
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Web Description / Profile / Introduction (i.e. Raw HTML)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6045
            Left            =   150
            TabIndex        =   39
            Top             =   60
            Width           =   7545
            Begin VB.TextBox txtBriefDesc 
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0043D143&
               Height          =   5355
               Left            =   270
               MultiLine       =   -1  'True
               ScrollBars      =   3  'Both
               TabIndex        =   40
               Top             =   420
               Width           =   6975
            End
         End
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reseller Accounts and Reseller Access"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   1
      Left            =   90
      TabIndex        =   52
      Top             =   570
      Width           =   4470
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0084E8E8&
      BorderWidth     =   2
      Index           =   2
      X1              =   7260
      X2              =   7410
      Y1              =   8010
      Y2              =   8010
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0084E8E8&
      BorderWidth     =   2
      Index           =   1
      X1              =   7290
      X2              =   7440
      Y1              =   7980
      Y2              =   7980
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Virtual ISP"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   510
      Index           =   0
      Left            =   60
      TabIndex        =   22
      Top             =   60
      Width           =   2520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0968F&
      BorderWidth     =   2
      X1              =   -240
      X2              =   12120
      Y1              =   540
      Y2              =   540
   End
End
Attribute VB_Name = "frmVISP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bCreateNew As Boolean


Public SESSIONCHAR As String

Dim bEditMade As Boolean
Dim bCreate As Boolean

Dim lRecID As Variant

Dim mButtonDown As Boolean
Dim iLastPoint As POINTAPI
Dim LastMovement As POINTAPI

Function SaveInformation()

    '*[ Error Checking Variables ]**********************************************************************************
    
    Const RoutineName = "SaveInformation"
    Const ContainerName = "frmVISP"
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


    If cmdSave.Enabled = False Then Exit Function
    
    If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Play "Write"
    
    'If bCreate = False Then Exit Function
    Dim oSeller As clsReseller
    
    If bCreateNew = True Then
        Set oSeller = New clsReseller
        oSeller.NextCycle = Format(DateAdd("m", 1, sysnow), "yyyy-mm-dd ttttt")
        oSeller.PreviousCycle = Format(sysnow, "yyyy-mm-dd ttttt")
        oSeller.CreatedBy_SysopID = Login.lSysopID
    Else
        Set oSeller = oReseller.CurrentReseller
    End If
    
    oSeller.Description = txtField(0).Text
    oSeller.Realm = txtField(1).Text
    oSeller.ABN = txtField(2).Text
    oSeller.ACN = txtField(3).Text
    oSeller.Subscribed = Val(txtField(7).Text)
    oSeller.JoiningFee = Val(txtField(8).Text)
    oSeller.VirtualID = Login.lVirtualID
    oSeller.BriefDesc = MySQL.ESC(txtBriefDesc.Text)
    oSeller.LogoURL = MySQL.ESC(txtLogoURL.Text)
    oSeller.cTaxCode = MySQL.ESC(txtTax(0).Text)
    oSeller.cTaxCountry = MySQL.ESC(txtTax(1).Text)
    oSeller.cTaxExemptCode = MySQL.ESC(txtTax(2).Text)
    oSeller.ftpHostName = MySQL.ESC(txtVISPFTP(0).Text)
    oSeller.ftpBasePath = MySQL.ESC(txtVISPFTP(1).Text)
    oSeller.ftpCostPerMB = Val(txtFileDB_CostPerMB.Text)
    Dim fg As Byte
    For fg = optFileDBMode.LBound To optFileDBMode.UBound
        If optFileDBMode(fg).Value = True Then
            oSeller.ftpFileDBMode = fg
            Exit For
        End If
    Next
    
    If Len(oSeller.ftpGroupingFolder) > 0 Then
        Dim jMD5 As New MD5
        oSeller.ftpGroupingFolder = jMD5.GetCheckSumFromString(oSeller.Description)
    End If
    
    oSeller.ftpIEProxy = chkIEProxy.Value
    oSeller.ftpPingAlive = chkPingAlive.Value
    oSeller.ftpPort = txtFileDB_Port.Text
    oSeller.ftpProxy = txtftpProxy.Text
    oSeller.ftpURLPath = txtFileDB_URL.Text
    oSeller.ftpPassword = txtFileDB_Password.Tag
    
    Dim lx As Long
    For lx = optTax.LBound To optTax.UBound
        If optTax(lx).Value = True Then
            oSeller.bTaxMode = lx
            Exit For
        End If
    Next
    
    If Val(txtField(4).Tag) <> 0 Then
        MySQL.Execute directConn, "UPDATE sysops SET VirtualID=" & lRecID & ",Password=encode(" & "'" & txtField(6).Text & "','" + odb.colSalts.ReturnSalt(PWSalt) + "'), Username=" & "'" & txtField(5).Text & " ', Description='" + MySQL.ESC(txtField(4).Text) + "' where RecID = " & txtField(4).Tag, True
    Else
        oSeller.SysopID = CloneSysop(directConn, Login.lSysopID, txtField(5).Text, txtField(6).Text, txtField(4).Text)
        txtField(4).Tag = oSeller.SysopID
    End If
    
    Set oReseller.CurrentReseller = oSeller
    Dim bOCreateNew As Boolean
    bOCreateNew = bCreateNew
    Call oReseller.CommitCurrentReseller(bCreateNew)
    If bOCreateNew Then
        
        Call oReseller.PopulateResellers(directConn, fs_LoadHeader, oSeller.RecID, SESSIONCHAR, Nothing, True)
    
    End If
        
    Call PopulateVISP
    
    If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Stop
    If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "Saved the Reseller information to main frame server. The weblink will update automatically, the URL for this Reseller Profile is click this message to see the page.", "http://www.projectalpha.com.au/vispprofile.php?nVirtualID=" & lRecID
    
    SaveInformation = lRecID
    
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

'Public Sub SaveAddresses(lRecID As Variant)
'
'
'    '*[ Error Checking Variables ]**********************************************************************************
'
'
'    Const RoutineName = "SaveAddresses"
'    Const ContainerName = "frmVISP"
'    '***************************************************************************************************************
'
'
''
''***********************************************************************************************
''**  Project Alpha ® 2003, 2004 +                                                             **
''***********************************************************************************************
''**  This code is not to be distributed, reverse engineered or simulated in any way without   **
''**  Premission from the author. The authors of this code is as follows: Simon Antony Roberts **                                                     **
''**Jarrett Cliff Costi, these two are the only people you can communicate with about this code**
''***********************************************************************************************
''**  Project Alpha is a product of Exitstencil Press Australia                                **
''***********************************************************************************************
''**                                                                                           **
''**  Routine:                                                                                 **
''**  Arguments:                                                                               **
''**  Description:    Subroutine, Function or Property of The Nexus                        **
''**  Author:         Simon Roberts                                                            **
''**  Date Last mod:  19-01-2004                                                               **
''**                                                                                           **
''********************************************** Copyright © 2004 Exitstencil Press Australia ***
''
''
''
'    If bDebug = -1 Then
'        On Error GoTo 0
'    ElseIf bDebug = 1 Then
'        On Error Resume Next
'    Else
'        On Error GoTo ErrorOccur
'    End If
'
'
'    Dim itmx As ListItem
'    Dim sa As Integer
'    Dim rsSave As ADODB.Recordset
'
'    If lvAddresses.ListItems.Count > 0 Then
'
'        For sa = 1 To lvAddresses.ListItems.Count
'            Set itmx = lvAddresses.ListItems(sa)
'            If itmx.Key = "" Then
'
'
'                On Error Resume Next
'                Do
'                    Err.Clear
'                    itmx.Key = "x" & MySQL.GetTMPRecID("visp_addresses", directConn)
'                    Call MySQL.Execute(directConn, "insert into visp_addresses (RecID) VALUES ('" & Mid(itmx.Key, 2) & "')")
'                    If Err.Number <> 0 Then cDebug Err.Description
'                Loop Until Err.Number = 0
'
'
'                Call MySQL.Execute(directConn, "update visp_addresses set visp_RecID = '" & lRecID & "', ContactName = '" & MySQL.ESC(itmx.Text) & "', " + _
'                                            "Street1 = '" & MySQL.ESC(itmx.SubItems(1)) & "', Street2 = '" & MySQL.ESC(itmx.SubItems(2)) & "', " + _
'                                            "Suburb = '" & MySQL.ESC(itmx.SubItems(3)) & "', State = '" & itmx.SubItems(4) & "', PostCode = '" & itmx.SubItems(5) & "', " + _
'                                            "Country = '" & MySQL.ESC(itmx.SubItems(6)) & "', Checked = '" & IIf(itmx.Checked = True, "-1", "0") & "' where RecID = '" & Mid(itmx.Key, 2) & "'")
'
'
'            ElseIf Left(itmx.Key, 1) = "e" Then
'
'                Call MySQL.Execute(directConn, "update visp_addresses set visp_RecID = '" & lRecID & "', ContactName = '" & MySQL.ESC(itmx.Text) & "', " + _
'                                            "Street1 = '" & MySQL.ESC(itmx.SubItems(1)) & "', Street2 = '" & MySQL.ESC(itmx.SubItems(2)) & "', " + _
'                                            "Suburb = '" & MySQL.ESC(itmx.SubItems(3)) & "', State = '" & itmx.SubItems(4) & "', PostCode = '" & itmx.SubItems(5) & "', " + _
'                                            "Country = '" & MySQL.ESC(itmx.SubItems(6)) & "', Checked = '" & IIf(itmx.Checked = True, "-1", "0") & "' where RecID = '" & Mid(itmx.Key, 2) & "'")
'
'            End If
'        Next
'
'    End If
'Exit Sub
'
'
'
'ErrorOccur:
'Select Case oErr.chkError(directConn,Val(Err.Number), Err.Description, RoutineName, ContainerName)
'Case vbResume
'    Resume
'Case vbExit
'
'Case vbResumeNext
'    Resume Next
'End Select
'
'End Sub
'
'Public Sub SavePhoneNumbers(lRecID As Variant)
'
'
'    '*[ Error Checking Variables ]**********************************************************************************
'
'
'    Const RoutineName = "SavePhoneNumbers"
'    Const ContainerName = "frmVISP"
'    '***************************************************************************************************************
'
'
''
''***********************************************************************************************
''**  Project Alpha ® 2003, 2004 +                                                             **
''***********************************************************************************************
''**  This code is not to be distributed, reverse engineered or simulated in any way without   **
''**  Premission from the author. The authors of this code is as follows: Simon Antony Roberts **                                                     **
''**Jarrett Cliff Costi, these two are the only people you can communicate with about this code**
''***********************************************************************************************
''**  Project Alpha is a product of Exitstencil Press Australia                                **
''***********************************************************************************************
''**                                                                                           **
''**  Routine:                                                                                 **
''**  Arguments:                                                                               **
''**  Description:    Subroutine, Function or Property of The Nexus                        **
''**  Author:         Simon Roberts                                                            **
''**  Date Last mod:  19-01-2004                                                               **
''**                                                                                           **
''********************************************** Copyright © 2004 Exitstencil Press Australia ***
''
''
''
'    If bDebug = -1 Then
'        On Error GoTo 0
'    ElseIf bDebug = 1 Then
'        On Error Resume Next
'    Else
'        On Error GoTo ErrorOccur
'    End If
'
'
'    Dim itmx As ListItem
'    Dim sa As Integer
'    Dim rsSave As ADODB.Recordset
'
'    If lvPhone.ListItems.Count > 0 Then
'
'        For sa = 1 To lvPhone.ListItems.Count
'            Set itmx = lvPhone.ListItems(sa)
'            If itmx.Key = "" Then
'
'
'                On Error Resume Next
'                Do
'                    Err.Clear
'                    itmx.Key = "x" & MySQL.GetTMPRecID("visp_phonenumbers", directConn)
'                    Call MySQL.Execute(directConn, "insert into visp_phonenumbers(RecID) VALUES ('" & Mid(itmx.Key, 2) & "')")
'                    If Err.Number <> 0 Then cDebug Err.Description
'                Loop Until Err.Number = 0
'
'
'                Call MySQL.Execute(directConn, "update visp_phonenumbers set visp_RecID = '" & lRecID & "', ContactName = '" & MySQL.ESC(itmx.Text) & "', " + _
'                                            "PhoneNumber = '" & MySQL.ESC(itmx.SubItems(1)) & "', Extension = '" & MySQL.ESC(itmx.SubItems(2)) & "', " + _
'                                            "ShortNote = '" & MySQL.ESC(itmx.SubItems(3)) & "', Checked = '" & IIf(itmx.Checked = True, "-1", "0") & "' where RecID = '" & Mid(itmx.Key, 2) & "'")
'
'
'            ElseIf Left(itmx.Key, 1) = "e" Then
'
'                Call MySQL.Execute(directConn, "update visp_phonenumbers set visp_RecID = '" & lRecID & "', ContactName = '" & MySQL.ESC(itmx.Text) & "', " + _
'                                            "PhoneNumber = '" & MySQL.ESC(itmx.SubItems(1)) & "', Extension = '" & MySQL.ESC(itmx.SubItems(2)) & "', " + _
'                                            "ShortNote = '" & MySQL.ESC(itmx.SubItems(3)) & "', Checked = '" & IIf(itmx.Checked = True, "-1", "0") & "' where RecID = '" & Mid(itmx.Key, 2) & "'")
'
'            End If
'        Next
'
'    End If
'
'Exit Sub
'
'
'
'ErrorOccur:
'Select Case oErr.chkError(directConn,Val(Err.Number), Err.Description, RoutineName, ContainerName)
'Case vbResume
'    Resume
'Case vbExit
'
'Case vbResumeNext
'    Resume Next
'End Select
'
'End Sub
'
'Public Sub SaveEmail(lRecID As Variant)
'
'
'    '*[ Error Checking Variables ]**********************************************************************************
'
'
'    Const RoutineName = "SaveEmail"
'    Const ContainerName = "frmVISP"
'    '***************************************************************************************************************
'
'
''
''***********************************************************************************************
''**  Project Alpha ® 2003, 2004 +                                                             **
''***********************************************************************************************
''**  This code is not to be distributed, reverse engineered or simulated in any way without   **
''**  Premission from the author. The authors of this code is as follows: Simon Antony Roberts **                                                     **
''**Jarrett Cliff Costi, these two are the only people you can communicate with about this code**
''***********************************************************************************************
''**  Project Alpha is a product of Exitstencil Press Australia                                **
''***********************************************************************************************
''**                                                                                           **
''**  Routine:                                                                                 **
''**  Arguments:                                                                               **
''**  Description:    Subroutine, Function or Property of The Nexus                        **
''**  Author:         Simon Roberts                                                            **
''**  Date Last mod:  19-01-2004                                                               **
''**                                                                                           **
''********************************************** Copyright © 2004 Exitstencil Press Australia ***
''
''
''
'    If bDebug = -1 Then
'        On Error GoTo 0
'    ElseIf bDebug = 1 Then
'        On Error Resume Next
'    Else
'        On Error GoTo ErrorOccur
'    End If
'
'
'    Dim itmx As ListItem
'    Dim sa As Integer
'    Dim rsSave As ADODB.Recordset
'
'    If lvEmail.ListItems.Count > 0 Then
'
'        For sa = 1 To lvEmail.ListItems.Count
'            Set itmx = lvEmail.ListItems(sa)
'            If itmx.Key = "" Then
'
'
'             On Error Resume Next
'                Do
'                    Err.Clear
'                    itmx.Key = "x" & MySQL.GetTMPRecID("visp_emailaddresses", directConn)
'                    Call MySQL.Execute(directConn, "insert into visp_emailaddresses (RecID) VALUES ('" & Mid(itmx.Key, 2) & "')")
'                    If Err.Number <> 0 Then cDebug Err.Description
'                Loop Until Err.Number = 0
'
'
'                Call MySQL.Execute(directConn, "update visp_emailaddresses set visp_RecID = '" & lRecID & "', ContactName = '" & MySQL.ESC(itmx.Text) & "', " + _
'                                            "Emailaddress = '" & MySQL.ESC(itmx.SubItems(1)) & "', Checked = '" & IIf(itmx.Checked = True, "-1", "0") & "' where RecID = '" & Mid(itmx.Key, 2) & "'")
'
'            ElseIf Left(itmx.Key, 1) = "e" Then
'
'                Call MySQL.Execute(directConn, "update visp_emailaddresses set visp_RecID = '" & lRecID & "', ContactName = '" & MySQL.ESC(itmx.Text) & "', " + _
'                                            "Emailaddress = '" & MySQL.ESC(itmx.SubItems(1)) & "', Checked = '" & IIf(itmx.Checked = True, "-1", "0") & "' where RecID = '" & Mid(itmx.Key, 2) & "'")
'
'            End If
'        Next
'
'    End If
'
'Exit Sub
'
'
'
'ErrorOccur:
'Select Case oErr.chkError(directConn,Val(Err.Number), Err.Description, RoutineName, ContainerName)
'Case vbResume
'    Resume
'Case vbExit
'
'Case vbResumeNext
'    Resume Next
'End Select
'
'End Sub
'
'


Private Sub cmdAddAddress_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdAddAddress_Click"
    Const ContainerName = "frmVISP"
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


    Dim ffrmSnailMail As frmSnailMail
    Set ffrmSnailMail = New frmSnailMail
    ffrmSnailMail.Show 1
        
    If ffrmSnailMail.iCloseState = frmCloseSave Then
    
        Dim itmX As ListItem
        Set itmX = lvAddresses.ListItems.Add(, "NEW" & oReseller.CurrentReseller.colResellers_SnailMail.Count + 1, "")
        
        With oReseller.CurrentReseller.colResellers_SnailMail.Add("NEW" & oReseller.CurrentReseller.colResellers_SnailMail.Count + 1, 0, oReseller.CurrentVirtualID, _
                                                                  ffrmSnailMail.FlagID, sysnow, ffrmSnailMail.sContactName, ffrmSnailMail.sStreetLine1, ffrmSnailMail.sStreetLine2, _
                                                                  ffrmSnailMail.sCountry, ffrmSnailMail.sState, ffrmSnailMail.sSuburb, ffrmSnailMail.sPostcode, False, True, "", _
                                                                  fs_NewLine_Insert, SESSIONCHAR, "NEW" & oReseller.CurrentReseller.colResellers_SnailMail.Count + 1)
           itmX.Text = .ContactName
           itmX.SubItems(1) = .Street1
           itmX.SubItems(2) = .Street2
           itmX.SubItems(3) = .Suburb
           itmX.SubItems(4) = .State
           itmX.SubItems(5) = .PostCode
           itmX.SubItems(6) = .Country
           itmX.Icon = oFlags.colFlags_IconCache.FINID(.FlagID)
           itmX.SmallIcon = itmX.Icon
           
        End With
        cmdSave.Enabled = True
        
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

Private Sub cmdAddEmail_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdAddEmail_Click"
    Const ContainerName = "frmVISP"
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


    Dim ffrmEmail As frmEmail
    Set ffrmEmail = New frmEmail
    If lvEmail.ListItems.Count = 0 Then
        ffrmEmail.sContactName = txtField(0).Text
        ffrmEmail.sEmailAddress = "@" & txtField(1).Text
    End If
    ffrmEmail.Show 1
    
    If ffrmEmail.iCloseState = frmCloseSave Then
        Dim itmX As ListItem
        
        With oReseller.CurrentReseller.colResellers_EmailAddy.Add("NEW" & oReseller.CurrentReseller.colResellers_EmailAddy.Count + 1, 0, oReseller.CurrentVirtualID, _
                                                            ffrmEmail.FlagID, sysnow, ffrmEmail.sEmailAddress, ffrmEmail.sContactName, False, True, "", fs_NewLine_Insert, _
                                                            SESSIONCHAR, "NEW" & oReseller.CurrentReseller.colResellers_EmailAddy.Count + 1)
            
            Set itmX = lvEmail.ListItems.Add(, .Key, .ContactName)
            itmX.Icon = oFlags.colFlags_IconCache.FINID(.FlagID)
            itmX.SmallIcon = itmX.Icon
            itmX.SubItems(1) = .EmailAddress
            
        End With
        cmdSave.Enabled = True
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



Private Sub cmdAddPhone_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdAddPhone_Click"
    Const ContainerName = "frmVISP"
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


    Dim ffrmPhoneNo As frmPhoneNumber
    Set ffrmPhoneNo = New frmPhoneNumber
    If lvPhone.ListItems.Count = 0 Then
        ffrmPhoneNo.sContactName = txtField(0).Text
    End If
    
    ffrmPhoneNo.Show 1
    
    If ffrmPhoneNo.iCloseState = frmCloseSave Then
        Dim itmX As ListItem
        With lvPhone.ListItems.Add(, , ffrmPhoneNo.sContactName)
            .SubItems(1) = ffrmPhoneNo.sPhonenumber
            .SubItems(2) = ffrmPhoneNo.sExtension
            .SubItems(3) = ffrmPhoneNo.sNote
            .Icon = oFlags.colFlags_IconCache.FINID(ffrmPhoneNo.sFlagID)
            .SmallIcon = .Icon
            .Key = "NEW" & oReseller.CurrentReseller.colResellers_FoneNum.Count + 1
            Call oReseller.CurrentReseller.colResellers_FoneNum.Add("NEW" & oReseller.CurrentReseller.colResellers_FoneNum.Count + 1, 0, oReseller.CurrentVirtualID, ffrmPhoneNo.sFlagID, sysnow, _
                .SubItems(1), .SubItems(2), .Text, False, True, "", fs_NewLine_Insert, SESSIONCHAR, "NEW" & oReseller.CurrentReseller.colResellers_FoneNum.Count + 1)
        End With
        cmdSave.Enabled = True
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

Public Function DisplayReseller(objVISP As clsReseller)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "DisplayReseller"
    Const ContainerName = "frmVISP"
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
    Dim rsload As adodb.Recordset
    Dim bResult As Boolean
    Dim itmX As ListItem

    If IsObject(objVISP) Then
    
        objVISP.FetchStatus = fs_LoadingData
    
        txtField(0).Text = objVISP.Description
        txtField(1).Text = objVISP.Realm
        txtField(2).Text = objVISP.ABN
        txtField(3).Text = objVISP.ACN
        txtField(7).Text = objVISP.Subscribed
        txtField(8).Text = objVISP.JoiningFee
        txtBriefDesc.Text = objVISP.BriefDesc
        txtLogoURL.Text = objVISP.LogoURL
        
        txtVISPFTP(0).Text = IIf(Len(objVISP.ftpHostName) = 0, "202.172.123.25", objVISP.ftpHostName)
        txtVISPFTP(1).Text = IIf(Len(objVISP.ftpBasePath) = 0, "$FileDB$", objVISP.ftpBasePath)
        txtFileDB_URL.Text = IIf(Len(objVISP.ftpURLPath) = 0, "http://202.172.123.25/$FileDB$/", objVISP.ftpURLPath)
        txtFileDB_Username.Text = IIf(Len(objVISP.ftpUsername) = 0, "daemon", objVISP.ftpUsername)
        txtFileDB_Password.Tag = IIf(Len(objVISP.ftpPassword) = 0, "daemon", objVISP.ftpPassword)
        txtFileDB_Port.Text = IIf(Len(objVISP.ftpPort) = 0, "21", objVISP.ftpPort)
        txtFileDB_CostPerMB.Text = IIf(objVISP.ftpCostPerMB = 0, "0.337", objVISP.ftpCostPerMB)
        optFileDBMode(objVISP.ftpFileDBMode).Value = True
        chkIEProxy.Value = objVISP.ftpIEProxy
        chkPingAlive.Value = objVISP.ftpPingAlive
        lblFTPStats.Caption = "Your resource has " & objVISP.ftpNumberofFolders & " folders and a total of " & objVISP.ftpNumberofFiles & " Files"
        Select Case objVISP.Cycle_IntervalType
        Case "m"
            lblFileDB_CycleIntervals.Caption = "Monthly Bases. At intervals of " & objVISP.Cycle_IntervalLength & " month" & IIf(objVISP.Cycle_IntervalLength > 1, "s. ", ". ") & "Next Invoice " & Format(objVISP.NextCycle, "mm/yyyy")
        Case "d"
            lblFileDB_CycleIntervals.Caption = "Daily Bases. At intervals of " & objVISP.Cycle_IntervalLength & " day" & IIf(objVISP.Cycle_IntervalLength > 1, "s. ", ". ") & "Next Invoice " & Format(objVISP.NextCycle, "mm/yyyy")
        Case "w"
            lblFileDB_CycleIntervals.Caption = "Weekday Bases. At intervals of " & objVISP.Cycle_IntervalLength & " weekday" & IIf(objVISP.Cycle_IntervalLength > 1, "s. ", ". ") & "Next Invoice " & Format(objVISP.NextCycle, "mm/yyyy")
        Case "ww"
            lblFileDB_CycleIntervals.Caption = "Weekly Bases. At intervals of " & objVISP.Cycle_IntervalLength & " week" & IIf(objVISP.Cycle_IntervalLength > 1, "s. ", ". ") & "Next Invoice " & Format(objVISP.NextCycle, "mm/yyyy")
        Case "h"
            lblFileDB_CycleIntervals.Caption = "Hourly Bases. At intervals of " & objVISP.Cycle_IntervalLength & " hour" & IIf(objVISP.Cycle_IntervalLength > 1, "s. ", ". ") & "Next Invoice " & Format(objVISP.NextCycle, "mm/yyyy")
        Case "n"
            lblFileDB_CycleIntervals.Caption = "Bases on Minutes. At intervals of " & objVISP.Cycle_IntervalLength & " minute" & IIf(objVISP.Cycle_IntervalLength > 1, "s. ", ". ") & "Next Invoice " & Format(objVISP.NextCycle, "mm/yyyy")
        Case "s"
            lblFileDB_CycleIntervals.Caption = "Bases on Seconds. At intervals of " & objVISP.Cycle_IntervalLength & " second" & IIf(objVISP.Cycle_IntervalLength > 1, "s. ", ". ") & "Next Invoice " & Format(objVISP.NextCycle, "mm/yyyy")
        Case "q"
            lblFileDB_CycleIntervals.Caption = "Quarterly Bases. At intervals of " & objVISP.Cycle_IntervalLength & " quarter" & IIf(objVISP.Cycle_IntervalLength > 1, "s. ", ". ") & "Next Invoice " & Format(objVISP.NextCycle, "mm/yyyy")
        Case "yyyy"
            lblFileDB_CycleIntervals.Caption = "Yearly Bases. At intervals of " & objVISP.Cycle_IntervalLength & " year" & IIf(objVISP.Cycle_IntervalLength > 1, "s. ", ". ") & "Next Invoice " & Format(objVISP.NextCycle, "mm/yyyy")
        End Select
        
        
        txtField(4).Tag = objVISP.SysopID
        
        txtTax(0).Text = objVISP.cTaxCode
        txtTax(1).Text = objVISP.cTaxCountry
        txtTax(2).Text = objVISP.cTaxExemptCode
        
        optTax(IIf(Val(objVISP.bTaxMode) <> 0, 1, 0)).Value = True
        
        bResult = MySQL.OpenTable(directConn, rsload, , "select RecID, Username, decode(Password,'" + odb.colSalts.ReturnSalt(PWSalt) + "') as Password, Description from sysops where RecID = " & txtField(4).Tag & " Limit 1")
        If rsload.RecordCount > 0 Then
            txtField(4).Text = rsload!Description
            txtField(5).Text = rsload!Username
            txtField(6).Text = rsload!Password
        Else
            txtField(4).Text = "Not Set"
            txtField(5).Text = ""
            txtField(6).Text = ""
        End If
    
            
            
        If objVISP.colResellers_SnailMail.Count > 0 Then
            Dim sMail As clsResellers_Addresses
            lvAddresses.ListItems.Clear
            For Each sMail In objVISP.colResellers_SnailMail
                With lvAddresses.ListItems.Add(, sMail.Key, sMail.ContactName, oFlags.colFlags_IconCache.FINID(sMail.FlagID), oFlags.colFlags_IconCache.FINID(sMail.FlagID))
                    .SubItems(1) = sMail.Street1
                    .SubItems(2) = sMail.Street2
                    .SubItems(3) = sMail.Suburb
                    .SubItems(4) = sMail.State
                    .SubItems(5) = sMail.PostCode
                    .SubItems(6) = sMail.Country
                    .Checked = IIf(sMail.Checked <> 0, True, False)
                End With
            Next
        End If
        
        If objVISP.colResellers_EmailAddy.Count > 0 Then
            Dim sPOP3 As clsResellers_EmailAddy
            lvEmail.ListItems.Clear
            For Each sPOP3 In objVISP.colResellers_EmailAddy
                With lvEmail.ListItems.Add(, sPOP3.Key, sPOP3.ContactName, oFlags.colFlags_IconCache.FINID(sPOP3.FlagID), oFlags.colFlags_IconCache.FINID(sPOP3.FlagID))
                    .SubItems(1) = sPOP3.EmailAddress
                    .Checked = IIf(sPOP3.Checked <> 0, True, False)
                End With
            Next
        End If
    
        If objVISP.colResellers_FoneNum.Count > 0 Then
            Dim sPhone As clsResellers_FoneNum
            lvPhone.ListItems.Clear
            For Each sPhone In objVISP.colResellers_FoneNum
                With lvPhone.ListItems.Add(, sPhone.Key, sPhone.ContactName, oFlags.colFlags_IconCache.FINID(sPhone.FlagID), oFlags.colFlags_IconCache.FINID(sPhone.FlagID))
                    .SubItems(1) = sPhone.PhoneNumber
                    .SubItems(2) = sPhone.Extension
                    .SubItems(3) = sPhone.ShortNote
                    .Checked = IIf(sPhone.Checked <> 0, True, False)
                End With
            Next
        End If
        
        Dim lx As Byte
        For lx = txtField.LBound To txtField.UBound
            txtField(lx).Locked = False
        Next
        
        objVISP.FetchStatus = fs_Idle
    
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

Private Sub cmdCreate_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdCreate_Click"
    Const ContainerName = "frmVISP"
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


    Select Case bCreateNew
    Case True
        bCreateNew = False
        cmdCreate.Caption = "&Create Reseller"
        cmdSave.Enabled = False
        bCreateNew = False
        
    Case False
        Set oReseller.CurrentReseller = New clsReseller
        cmdCreate.Caption = "&Cancel Create"
        bCreateNew = True
        cmdSave.Enabled = True
        
        Dim jk As Byte
        For jk = txtField.LBound To txtField.UBound
            txtField(jk).Text = ""
            txtField(jk).Locked = False
        Next
        
        lvPhone.ListItems.Clear
        lvEmail.ListItems.Clear
        lvAddresses.ListItems.Clear
        
        txtTax(0).Text = "GST"
        txtTax(1).Text = "AUS0001"
        txtTax(2).Text = ""
        txtLogoURL.Text = "http://"
        txtBriefDesc.Text = ""
        
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

Private Sub cmdSave_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdSave_Click"
    Const ContainerName = "frmVISP"
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


    If Trim(txtField(0)) = "" Then
        MsgBox "You must put the company name in!", vbCritical, "Missing Ingtomation"
        Exit Sub
    End If
    
    If Trim(txtField(1)) = "" Then
        MsgBox "You must put the company's domain name !", vbCritical, "Missing Ingtomation"
        Exit Sub
    End If
    
    If Trim(txtField(3)) = "" Then
        MsgBox "You must put the company's ACN or RBN in!", vbCritical, "Missing Ingtomation"
        Exit Sub
    End If
    
    If Trim(txtField(4)) = "" Then
        MsgBox "You must describe the companies primary sysop account!", vbCritical, "Missing Ingtomation"
        Exit Sub
    End If
    
    If Trim(txtField(5)) = "" Then
        MsgBox "You must enter the companies primary sysop username!", vbCritical, "Missing Ingtomation"
        Exit Sub
    End If
    
    If Trim(txtField(6)) = "" Then
        MsgBox "You must enter the companies primary sysop password!", vbCritical, "Missing Ingtomation"
        Exit Sub
    End If
    
    SaveInformation
    
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

Private Sub cmdSubscribe_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdSubscribe_Click"
    Const ContainerName = "frmVISP"
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


    Dim fSub As frmSubscribe
    Set fSub = New frmSubscribe
    fSub.lSubscribe = Val(txtField(7))
    fSub.Show 1
    txtField(7) = "" & fSub.lSubscribe
    
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
    Const ContainerName = "frmVISP"
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

    Set lvVISP.SmallIcons = fIcon.il16x16
    Set lvVISP.Icons = fIcon.il32x32
    Set lvAddresses.SmallIcons = fIcon.il16x16
    Set lvAddresses.Icons = fIcon.il32x32
    Set lvPhone.SmallIcons = fIcon.il16x16
    Set lvPhone.Icons = fIcon.il32x32
    Set lvEmail.SmallIcons = fIcon.il16x16
    Set lvEmail.Icons = fIcon.il32x32

    frmTax.Enabled = Login.bMaster
    frmFileDBCosts.Enabled = Login.bMaster
    
    SESSIONCHAR = GetSessionChar(SESSIONCHAR, Me.hwnd, 14)
    
    Call GUI.LoadColWidths(lvVISP, Me)
    Call GUI.LoadColWidths(lvPhone, Me)
    Call GUI.LoadColWidths(lvEmail, Me)
    Call GUI.LoadColWidths(lvAddresses, Me)
        
    picResize.Left = GetSetting("projectalpha", Me.Name, "Resize", picResize.Left)
    lRecID = 0
    
    
    If bBigFont = True Then
    
        lvVISP.Font.Size = 18
        lvEmail.Font.Size = 18
        lvPhone.Font.Size = 18
        lvAddresses.Font.Size = 18
        ts.Font.Size = 16
        
    End If
    
    lvVISP.Visible = False
    Me.Show
    
    gSleep
    
    oReseller.Clear
    Call oReseller.PopulateResellers(directConn, fs_LoadHeader, Login.lVirtualID, SESSIONCHAR, Nothing, True)
    
    PopulateVISP
    
    lvVISP.Visible = True
    
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_QueryUnload"
    Const ContainerName = "frmVISP"
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

    
    Call GUI.SaveColWidths(lvVISP, Me)
    Call GUI.SaveColWidths(lvPhone, Me)
    Call GUI.SaveColWidths(lvEmail, Me)
    Call GUI.SaveColWidths(lvAddresses, Me)
    
    SaveSetting "projectalpha", Me.Name, "Resize", picResize.Left
    If bEditMade = True Then
        Select Case MsgBox("Do you wish to save the information entered for this Reseller?", vbQuestion + vbYesNo, "Save Information?")
        Case vbYes
            cmdSave.Enabled = True
            Call SaveInformation
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

Private Sub Form_Resize()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Resize"
    Const ContainerName = "frmVISP"
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


    If Me.WindowState = vbMinimized Then Exit Sub
    picResize.Move picResize.Left, picResize.Top, picResize.Width, Me.ScaleHeight - picResize.Top - 60
    picListview.Height = Me.ScaleHeight - picListview.Top - 120
    picTab.Height = Me.ScaleHeight - picListview.Top - 120
    Line1.X1 = 0
    Line1.X2 = Me.ScaleWidth
    picListview.Width = picResize.Left
    picTab.Width = IIf(Me.Width - picResize.Left - picResize.Width - 120 < 0, 10, Me.Width - picResize.Left - picResize.Width - 120)
    picTab.Left = picResize.Left + picResize.Width
    
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


Private Sub lvEmail_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvEmail_ColumnClick"
    Const ContainerName = "frmVISP"
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


    Call GUI.ColumnSort(ColumnHeader, lvEmail)
    
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

Private Sub lvEmail_DblClick()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvEmail_DblClick"
    Const ContainerName = "frmVISP"
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


    If lvEmail.Tag <> "" Then
    
        Dim ffrmEmail As frmEmail
        Set ffrmEmail = New frmEmail
        Dim itmX As ListItem
        Set itmX = lvEmail.SelectedItem
        
        ffrmEmail.sContactName = oReseller.CurrentReseller.colResellers_EmailAddy(itmX.Key).ContactName
        ffrmEmail.sEmailAddress = oReseller.CurrentReseller.colResellers_EmailAddy(itmX.Key).EmailAddress
        ffrmEmail.FlagID = oReseller.CurrentReseller.colResellers_EmailAddy(itmX.Key).FlagID
        
        ffrmEmail.Show 1
        
        If ffrmEmail.iCloseState = frmCloseSave Then
                        
            If Not ffrmEmail.sContactName = oReseller.CurrentReseller.colResellers_EmailAddy(itmX.Key).ContactName Or Not _
                        ffrmEmail.sEmailAddress = oReseller.CurrentReseller.colResellers_EmailAddy(itmX.Key).EmailAddress Or Not _
                        ffrmEmail.FlagID = oReseller.CurrentReseller.colResellers_EmailAddy(itmX.Key).FlagID Then
            
                oReseller.CurrentReseller.colResellers_EmailAddy(itmX.Key).FetchStatus = fs_Edited
                oReseller.CurrentReseller.colResellers_EmailAddy(itmX.Key).ContactName = ffrmEmail.sContactName
                oReseller.CurrentReseller.colResellers_EmailAddy(itmX.Key).EmailAddress = ffrmEmail.sEmailAddress
                oReseller.CurrentReseller.colResellers_EmailAddy(itmX.Key).FlagID = ffrmEmail.FlagID
                oReseller.CurrentReseller.colResellers_EmailAddy(itmX.Key).FetchStatus = fs_Edited
                
                cmdSave.Enabled = True
                
                itmX.Text = ffrmEmail.sContactName
                itmX.SubItems(1) = ffrmEmail.sEmailAddress
                itmX.Icon = oFlags.colFlags_IconCache.FINID(ffrmEmail.FlagID)
                itmX.SmallIcon = itmX.Icon
            End If
            
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

Private Sub lvEmail_ItemCheck(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvEmail_ItemCheck"
    Const ContainerName = "frmVISP"
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

    oReseller.CurrentReseller.colResellers_EmailAddy(Item.Key).Checked = Item.Checked
    oReseller.CurrentReseller.colResellers_EmailAddy(Item.Key).FetchStatus = fs_Edited
    
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

Private Sub lvEmail_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvEmail_ItemClick"
    Const ContainerName = "frmVISP"
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


    lvEmail.Tag = True
    
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

Private Sub lvPhone_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvPhone_ColumnClick"
    Const ContainerName = "frmVISP"
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


    Call GUI.ColumnSort(ColumnHeader, lvPhone)
    
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

Private Sub lvPhone_DblClick()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvPhone_DblClick"
    Const ContainerName = "frmVISP"
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


    If lvPhone.Tag <> "" Then
        
        Dim ffrmPhoneNo As frmPhoneNumber
        Dim itmX As ListItem
        Set ffrmPhoneNo = New frmPhoneNumber
        Set itmX = lvPhone.SelectedItem
        
        ffrmPhoneNo.sContactName = oReseller.CurrentReseller.colResellers_FoneNum(itmX.Key).ContactName
        ffrmPhoneNo.sExtension = oReseller.CurrentReseller.colResellers_FoneNum(itmX.Key).Extension
        ffrmPhoneNo.sPhonenumber = oReseller.CurrentReseller.colResellers_FoneNum(itmX.Key).PhoneNumber
        ffrmPhoneNo.sNote = oReseller.CurrentReseller.colResellers_FoneNum(itmX.Key).ShortNote
        ffrmPhoneNo.sFlagID = oReseller.CurrentReseller.colResellers_FoneNum(itmX.Key).FlagID
                    
        ffrmPhoneNo.Show 1
        
        If ffrmPhoneNo.iCloseState = frmCloseSave Then
            If Not oReseller.CurrentReseller.colResellers_FoneNum(itmX.Key).ContactName = ffrmPhoneNo.sContactName Or _
                Not oReseller.CurrentReseller.colResellers_FoneNum(itmX.Key).Extension = ffrmPhoneNo.sExtension Or _
                Not oReseller.CurrentReseller.colResellers_FoneNum(itmX.Key).PhoneNumber = ffrmPhoneNo.sPhonenumber Or _
                Not oReseller.CurrentReseller.colResellers_FoneNum(itmX.Key).ShortNote = ffrmPhoneNo.sNote Or _
                Not oReseller.CurrentReseller.colResellers_FoneNum(itmX.Key).FlagID = ffrmPhoneNo.sFlagID Then
                
                oReseller.CurrentReseller.colResellers_FoneNum(itmX.Key).FetchStatus = fs_Edited
                
                itmX.Text = ffrmPhoneNo.sContactName
                itmX.SubItems(1) = ffrmPhoneNo.sPhonenumber
                itmX.SubItems(2) = ffrmPhoneNo.sExtension
                itmX.SubItems(3) = ffrmPhoneNo.sNote
                itmX.SmallIcon = oFlags.colFlags_IconCache.FINID(ffrmPhoneNo.sFlagID)
                itmX.Icon = oFlags.colFlags_IconCache.FINID(ffrmPhoneNo.sFlagID)
                cmdSave.Enabled = True
                oReseller.CurrentReseller.colResellers_FoneNum(itmX.Key).ContactName = ffrmPhoneNo.sContactName
                oReseller.CurrentReseller.colResellers_FoneNum(itmX.Key).Extension = ffrmPhoneNo.sExtension
                oReseller.CurrentReseller.colResellers_FoneNum(itmX.Key).PhoneNumber = ffrmPhoneNo.sPhonenumber
                oReseller.CurrentReseller.colResellers_FoneNum(itmX.Key).ShortNote = ffrmPhoneNo.sNote
                oReseller.CurrentReseller.colResellers_FoneNum(itmX.Key).FlagID = ffrmPhoneNo.sFlagID
                oReseller.CurrentReseller.colResellers_FoneNum(itmX.Key).FetchStatus = fs_Edited
            End If
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

Private Sub lvPhone_ItemCheck(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvPhone_ItemCheck"
    Const ContainerName = "frmVISP"
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

    oReseller.CurrentReseller.colResellers_FoneNum(Item.Key).Checked = Item.Checked
    oReseller.CurrentReseller.colResellers_FoneNum(Item.Key).FetchStatus = fs_Edited
    
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

Private Sub lvPhone_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvPhone_ItemClick"
    Const ContainerName = "frmVISP"
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


    lvPhone.Tag = True
    
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
Private Sub lvAddresses_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvAddresses_ColumnClick"
    Const ContainerName = "frmVISP"
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


    Call GUI.ColumnSort(ColumnHeader, lvAddresses)
    
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

Private Sub lvAddresses_DblClick()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvAddresses_DblClick"
    Const ContainerName = "frmVISP"
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


    If lvAddresses.Tag <> "" Then
    
        Dim ffrmSnailMail As frmSnailMail
        Dim itmX As ListItem
        Set ffrmSnailMail = New frmSnailMail
        Set itmX = lvAddresses.SelectedItem
        
        
    
    
        ffrmSnailMail.sContactName = oReseller.CurrentReseller.colResellers_SnailMail(itmX.Key).ContactName
        ffrmSnailMail.sStreetLine1 = oReseller.CurrentReseller.colResellers_SnailMail(itmX.Key).Street1
        ffrmSnailMail.sStreetLine2 = oReseller.CurrentReseller.colResellers_SnailMail(itmX.Key).Street2
        ffrmSnailMail.sSuburb = oReseller.CurrentReseller.colResellers_SnailMail(itmX.Key).Suburb
        ffrmSnailMail.sState = oReseller.CurrentReseller.colResellers_SnailMail(itmX.Key).State
        ffrmSnailMail.sPostcode = oReseller.CurrentReseller.colResellers_SnailMail(itmX.Key).PostCode
        ffrmSnailMail.sCountry = oReseller.CurrentReseller.colResellers_SnailMail(itmX.Key).Country
        ffrmSnailMail.FlagID = oReseller.CurrentReseller.colResellers_SnailMail(itmX.Key).FlagID
        
    
        ffrmSnailMail.Show 1
        
        If ffrmSnailMail.iCloseState = frmCloseSave Then
            
            
            If Not ffrmSnailMail.sContactName = oReseller.CurrentReseller.colResellers_SnailMail(itmX.Key).ContactName Or Not _
                    ffrmSnailMail.sStreetLine1 = oReseller.CurrentReseller.colResellers_SnailMail(itmX.Key).Street1 Or Not _
                    ffrmSnailMail.sStreetLine2 = oReseller.CurrentReseller.colResellers_SnailMail(itmX.Key).Street2 Or Not _
                    ffrmSnailMail.sSuburb = oReseller.CurrentReseller.colResellers_SnailMail(itmX.Key).Suburb Or Not _
                    ffrmSnailMail.sState = oReseller.CurrentReseller.colResellers_SnailMail(itmX.Key).State Or Not _
                    ffrmSnailMail.sPostcode = oReseller.CurrentReseller.colResellers_SnailMail(itmX.Key).PostCode Or Not _
                    ffrmSnailMail.sCountry = oReseller.CurrentReseller.colResellers_SnailMail(itmX.Key).Country Or Not _
                    ffrmSnailMail.FlagID = oReseller.CurrentReseller.colResellers_SnailMail(itmX.Key).FlagID Then
                    
                    cmdSave.Enabled = True
                    
                    oReseller.CurrentReseller.colResellers_SnailMail(itmX.Key).ContactName = ffrmSnailMail.sContactName
                    oReseller.CurrentReseller.colResellers_SnailMail(itmX.Key).Street1 = ffrmSnailMail.sStreetLine1
                    oReseller.CurrentReseller.colResellers_SnailMail(itmX.Key).Street2 = ffrmSnailMail.sStreetLine2
                    oReseller.CurrentReseller.colResellers_SnailMail(itmX.Key).Suburb = ffrmSnailMail.sSuburb
                    oReseller.CurrentReseller.colResellers_SnailMail(itmX.Key).State = ffrmSnailMail.sState
                    oReseller.CurrentReseller.colResellers_SnailMail(itmX.Key).PostCode = ffrmSnailMail.sPostcode
                    oReseller.CurrentReseller.colResellers_SnailMail(itmX.Key).Country = ffrmSnailMail.sCountry
                    oReseller.CurrentReseller.colResellers_SnailMail(itmX.Key).FlagID = ffrmSnailMail.FlagID
                    oReseller.CurrentReseller.colResellers_SnailMail(itmX.Key).FetchStatus = fs_Edited
                                        
                    With itmX
                        .Text = ffrmSnailMail.sContactName
                        .SubItems(1) = ffrmSnailMail.sStreetLine1
                        .SubItems(2) = ffrmSnailMail.sStreetLine2
                        .SubItems(3) = ffrmSnailMail.sSuburb
                        .SubItems(4) = ffrmSnailMail.sState
                        .SubItems(5) = ffrmSnailMail.sPostcode
                        .SubItems(6) = ffrmSnailMail.sCountry
                        .Icon = oFlags.colFlags_IconCache.FINID(ffrmSnailMail.FlagID)
                        .SmallIcon = .Icon
                    End With
                
                End If
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

Private Sub lvAddresses_ItemCheck(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvAddresses_ItemCheck"
    Const ContainerName = "frmVISP"
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


    oReseller.CurrentReseller.colResellers_SnailMail(Item.Key).Checked = Item.Checked
    oReseller.CurrentReseller.colResellers_SnailMail(Item.Key).FetchStatus = fs_Edited
    
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

Private Sub lvAddresses_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvAddresses_ItemClick"
    Const ContainerName = "frmVISP"
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


    lvAddresses.Tag = True
    
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


Private Sub lvVISP_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvVISP_ColumnClick"
    Const ContainerName = "frmVISP"
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


    Call GUI.ColumnSort(ColumnHeader, lvVISP)
    
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

Private Sub lvVISP_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvVISP_ItemClick"
    Const ContainerName = "frmVISP"
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
    
    'If Not oReseller.CurrentReseller Is Nothing Then SaveInformation

    Set oReseller.CurrentReseller = oReseller(Item.Key)
            
    lblArticles(0).Tag = oReseller.PopulateResellers(directConn, fs_LoadAllContactDetails, oReseller.CurrentVirtualID, SESSIONCHAR, pb, True)
    If Val(lblArticles(0).Tag) > 0 Then
        lblArticles(0).Caption = "There where " & lblArticles(0).Tag & " articles and records retrieved."
    Else
        lblArticles(0).Caption = "There where no new articles and records to retrieved."
    End If
    lblArticles(1).Tag = "" & (Val(lblArticles(1).Tag) + Val(lblArticles(0).Tag))
    lblArticles(1).Caption = "You have loaded " & lblArticles(1).Tag & " articles or records."
    
    Set oReseller.CurrentReseller = oReseller(Item.Key)
    
    Call DisplayReseller(oReseller.CurrentReseller)
    
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

Private Sub mTimer_Timer()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mTimer_Timer"
    Const ContainerName = "frmVISP"
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

    Dim Rec As RECT, Point As POINTAPI
   
    GetCursorPos Point

    Point.X = Point.X * Screen.TwipsPerPixelX
    Point.Y = Point.Y * Screen.TwipsPerPixelY
    
    If Point.X - Me.Left < 10 Then
        picResize.Left = 0
    ElseIf Point.X - Me.Left > Me.Width - picResize.Width * 2 Then
        picResize.Left = Me.Width - picResize.Width * 2
    Else
        picResize.Left = Point.X - Me.Left - picResize.Width / 2
    End If
    
    picListview.Width = picResize.Left
    picTab.Width = IIf(Me.Width - picResize.Left - picResize.Width - 120 < 0, 10, Me.Width - picResize.Left - picResize.Width - 120)
    picTab.Left = picResize.Left + picResize.Width
    
    LastMovement.X = (Point.X - iLastPoint.X)
    LastMovement.Y = (Point.Y - iLastPoint.Y)
        
    gSleep
    
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

Private Sub optTax_Click(Index As Integer)

    Select Case Index
    Case 0
        frmTaxCode(0).Enabled = True
        frmTaxCode(1).Enabled = False
    Case 1
        frmTaxCode(1).Enabled = True
        frmTaxCode(0).Enabled = False
    End Select
    
End Sub

Private Sub picListview_Resize()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "picListview_Resize"
    Const ContainerName = "frmVISP"
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


    If picListview.ScaleWidth > 300 And picListview.ScaleHeight > 300 Then
        lvVISP.Move 60, 600, picListview.ScaleWidth - 120, picListview.ScaleHeight - 200 - pb.Height - 600
        pb.Move 60, lvVISP.Height + lvVISP.Top + 60, picListview.ScaleWidth - 120, pb.Height
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

Private Sub picResize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "picResize_MouseDown"
    Const ContainerName = "frmVISP"
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


    Dim Rec As RECT, Point As POINTAPI
    
    If Button = 1 Then
        If mButtonDown = False Then
            'picDrag.Drag vbBeginDrag
            GetWindowRect Me.hwnd, Rec
            GetCursorPos Point
            LastMovement.X = 0
            LastMovement.Y = 0
            iLastPoint.X = Point.X * Screen.TwipsPerPixelX
            iLastPoint.Y = Point.Y * Screen.TwipsPerPixelY
            mButtonDown = True
            mTimer.Enabled = True
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



Private Sub picResize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "picResize_MouseUp"
    Const ContainerName = "frmVISP"
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


   'picDrag.Drag 0

    mButtonDown = False
    mTimer.Enabled = False
    
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

Private Sub picTab_Resize()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "picTab_Resize"
    Const ContainerName = "frmVISP"
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


    If picTab.ScaleHeight > 800 And picTab.ScaleWidth > 800 Then
        ts.Move 60, 60, picTab.ScaleWidth - 120, picTab.ScaleHeight - 120
        If tsContainer.UBound >= ts.SelectedItem.Index - 1 Then
            tsContainer(ts.SelectedItem.Index - 1).Move ts.ClientLeft, ts.ClientTop, ts.ClientWidth, ts.ClientHeight
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

Private Sub tmrIcons_Timer()

    Static bSet As Boolean
    
    If Login.IconsSet = True Then
        If bSet = False Then
            
            Set lvVISP.SmallIcons = fIcon.il16x16
            Set lvVISP.Icons = fIcon.il32x32
            Set lvAddresses.SmallIcons = fIcon.il16x16
            Set lvAddresses.Icons = fIcon.il32x32
            Set lvPhone.SmallIcons = fIcon.il16x16
            Set lvPhone.Icons = fIcon.il32x32
            Set lvEmail.SmallIcons = fIcon.il16x16
            Set lvEmail.Icons = fIcon.il32x32
            
            bSet = True
        End If
    End If
    
End Sub

Private Sub ts_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "ts_Click"
    Const ContainerName = "frmVISP"
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
    
    For X = tsContainer.LBound To tsContainer.UBound
        If ts.SelectedItem.Index - 1 <> X Then tsContainer(X).Visible = False
    Next
    
    If ts.SelectedItem.Index - 1 <= tsContainer.UBound Then
        tsContainer(ts.SelectedItem.Index - 1).Move ts.ClientLeft, ts.ClientTop, ts.ClientWidth, ts.ClientHeight
        tsContainer(ts.SelectedItem.Index - 1).Visible = True
        tsContainer(ts.SelectedItem.Index - 1).ZOrder 0
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

Private Sub tsContainer_Resize(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "tsContainer_Resize"
    Const ContainerName = "frmVISP"
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

    
    Dim ix As Integer
    
    Select Case Index
    Case 0
        
        For ix = frame.LBound To frame.UBound
            frame(ix).Move frame(ix).Left, frame(ix).Top, tsContainer(Index).Width - frame(ix).Left * 2, frame(ix).Height
        Next
            
        For ix = txtField.LBound To txtField.UBound
            If frame(0).Width > 500 Then txtField(ix).Move txtField(ix).Left, txtField(ix).Top, frame(0).Width - txtField(ix).Left * 2 ', txtfield(iX).Height
        Next
        
        For ix = Label2.LBound To Label2.UBound
            If frame(0).Width > 500 Then Label2(ix).Move Label2(ix).Left, Label2(ix).Top, frame(0).Width - Label2(ix).Left * 2, Label2(ix).Height
        Next
        
        cmdCreate.Move cmdCreate.Left, cmdCreate.Top
        cmdSave.Move tsContainer(Index).ScaleWidth - cmdCreate.Width - 60, cmdSave.Top
        
        cmdSubscribe.Move tsContainer(Index).ScaleWidth - cmdSubscribe.Width - 60, tsContainer(Index).ScaleHeight - cmdSubscribe.Height - 60
    Case 1
        lvPhone.Move 60, 60, tsContainer(Index).ScaleWidth - 120, tsContainer(Index).ScaleHeight - 180 - cmdAddPhone.Height
        lvPhone.Refresh
        cmdAddPhone.Move 60, lvPhone.Top + lvPhone.Height + 60
    Case 2
        lvEmail.Move 60, 60, tsContainer(Index).ScaleWidth - 120, tsContainer(Index).ScaleHeight - 180 - cmdAddPhone.Height
        lvEmail.Refresh
        cmdAddEmail.Move 60, lvEmail.Top + lvEmail.Height + 60
    Case 3
        lvAddresses.Move 60, 60, tsContainer(Index).ScaleWidth - 120, tsContainer(Index).ScaleHeight - 180 - cmdAddPhone.Height
        lvAddresses.Refresh
        cmdAddAddress.Move 60, lvAddresses.Top + lvAddresses.Height + 60
    Case 4
        Frame2.Move 120, tsContainer(Index).ScaleHeight - Frame2.Height - 240, tsContainer(Index).ScaleWidth - 240, Frame2.Height
        txtLogoURL.Move 120, 240, Frame2.Width - 240, Frame2.Height - 460
        
        Frame1.Move 240, 240, tsContainer(Index).ScaleWidth - 480, tsContainer(Index).ScaleHeight - 600 - Frame2.Height
        txtBriefDesc.Move 120, 240, Frame1.Width - 240, Frame1.Height - 360
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

Private Sub txtfield_Change(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtfield_Change"
    Const ContainerName = "frmVISP"
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


    bEditMade = True
    
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

Public Sub PopulateVISP()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "PopulateVISP"
    Const ContainerName = "frmVISP"
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
    
    On Error GoTo 0
    Dim itmX As ListItem
    Dim bFound As Boolean
    
    If oReseller.Count > 0 Then
        Dim lk As Long
        pb.Value = 0
        pb.Max = oReseller.Count
        For lk = 1 To oReseller.Count
            If Not oReseller(lk).RecID = 0 And Not oReseller(lk).RecID = 1000 Then
                bFound = False
                For Each itmX In lvVISP.ListItems
                    If itmX.Key = oReseller(lk).Key Then
                        bFound = True
                        With itmX
                            If Not .Text = oReseller(lk).Description Then .Text = oReseller(lk).Description
                            If Not .SmallIcon = IIf(fIcon.il32x32.ListImages.Count = 1, 1, oReseller(lk).Icon) Then .SmallIcon = IIf(fIcon.il32x32.ListImages.Count = 1, 1, oReseller(lk).Icon)
                            If Not .Icon = IIf(fIcon.il32x32.ListImages.Count = 1, 1, oReseller(lk).Icon) Then .Icon = IIf(fIcon.il32x32.ListImages.Count = 1, 1, oReseller(lk).Icon)
                            If Not .SubItems(1) = oReseller(lk).ABN Then .SubItems(1) = oReseller(lk).ABN
                            If Not .SubItems(2) = oReseller(lk).ACN Then .SubItems(2) = oReseller(lk).ACN
                            If Not .SubItems(3) = oReseller(lk).NoSub Then .SubItems(3) = oReseller(lk).NoSub
                            If Not .SubItems(4) = oReseller(lk).Subscribed Then .SubItems(4) = oReseller(lk).Subscribed
                            If Not .SubItems(5) = oReseller(lk).LogoURL Then .SubItems(5) = oReseller(lk).LogoURL
                            If Not .SubItems(6) = oReseller(lk).NextCycle Then .SubItems(6) = oReseller(lk).NextCycle
                            If Not .SubItems(7) = oReseller(lk).ftpHostName Then .SubItems(7) = oReseller(lk).ftpHostName
                        End With
                        Exit For
                    End If
                Next
                
                If bFound = False Then
                    With lvVISP.ListItems.Add(, oReseller(lk).Key, oReseller(lk).Description, IIf(fIcon.il32x32.ListImages.Count = 1, 1, oReseller(lk).Icon), IIf(fIcon.il32x32.ListImages.Count = 1, 1, oReseller(lk).Icon))
                        .SubItems(1) = oReseller(lk).ABN
                        .SubItems(2) = oReseller(lk).ACN
                        .SubItems(3) = oReseller(lk).NoSub
                        .SubItems(4) = oReseller(lk).Subscribed
                        .SubItems(5) = oReseller(lk).LogoURL
                        .SubItems(6) = oReseller(lk).NextCycle
                        .SubItems(7) = oReseller(lk).ftpHostName
                    End With
                End If
                
            End If
            pb.Value = lk
            pb.Refresh
            gSleep
        Next
        lvVISP.Visible = True
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

Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtfield_KeyPress"
    Const ContainerName = "frmVISP"
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
    Case 8
        Select Case KeyAscii
        Case 48 To 57, 8
        Case Asc(".")
            If InStr(txtField(8).Text, ".") > 0 Then KeyAscii = 0
        Case 13
            SendKeys "{TAB}"
        Case Else
            KeyAscii = 0
        End Select
    Case Else
        Select Case KeyAscii
        Case 13
            SendKeys "{TAB}"
        End Select
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

Private Sub txtFileDB_CostPerMB_Change()

    txtFileDB_CostPerMB.ToolTipText = "GST Inclusive Price " & Format(Val(txtFileDB_CostPerMB) * oTax(Login.TaxCode, Login.TaxCountry) + Val(txtFileDB_CostPerMB), "Currency")
    lblFileDB_TotalMBs.Caption = Format(oReseller.CurrentReseller.ftpTotalMB, "###,###,###,###,###,###.##0 MBs")
    lblFileDB_TotalCost.Caption = Format((oReseller.CurrentReseller.ftpTotalMB * txtFileDB_CostPerMB) + ((oReseller.CurrentReseller.ftpTotalMB * txtFileDB_CostPerMB) * oTax(Login.TaxCode, Login.TaxCountry)), "currency")
    lblFileDB_TotalCost.ToolTipText = "GST Exclusive Price " & Format((oReseller.CurrentReseller.ftpTotalMB * txtFileDB_CostPerMB), "Currency")
    
End Sub

Private Sub txtFileDB_CostPerMB_DblClick()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtFileDB_CostPerMB_DblClick"
    Const ContainerName = "frmVISP"
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


    frmGSTCalc.Show 1
    txtFileDB_CostPerMB = "" & frmGSTCalc.cAmount
    
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

Private Sub txtFileDB_CostPerMB_GotFocus()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtFileDB_CostPerMB_GotFocus"
    Const ContainerName = "frmVISP"
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


    txtFileDB_CostPerMB.SelStart = 0
    txtFileDB_CostPerMB.SelLength = Len(txtFileDB_CostPerMB.Text)
    
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

Private Sub txtFileDB_CostPerMB_KeyPress(KeyAscii As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtFileDB_CostPerMB_KeyPress"
    Const ContainerName = "frmVISP"
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
        If InStr(txtFileDB_CostPerMB, ".") > 0 Then KeyAscii = 0
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

Private Sub txtFileDB_Password_Change()

    txtFileDB_Password.Tag = txtFileDB_Password.Text
    
End Sub

