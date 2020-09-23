VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAccHoldings 
   Caption         =   "Account Holdings"
   ClientHeight    =   3705
   ClientLeft      =   1935
   ClientTop       =   4065
   ClientWidth     =   9345
   Icon            =   "frmAccHoldings2_new.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3705
   ScaleWidth      =   9345
   Begin VB.CommandButton Command1 
      BackColor       =   &H004FD2F9&
      Caption         =   "Export as HTML"
      Height          =   315
      Left            =   1170
      MaskColor       =   &H00EC7A71&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   540
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   1000
      Left            =   2520
      Top             =   210
   End
   Begin VB.TextBox txtRefreshMin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   8040
      TabIndex        =   7
      Text            =   "10"
      Top             =   210
      Width           =   765
   End
   Begin MSComctlLib.ImageList ilTreeview 
      Left            =   3240
      Top             =   180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   87
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":15A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":19F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":1E48
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":229A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":26EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":2B3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":2F90
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":33E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":3834
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":3C86
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":40D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":452A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":497C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":4DCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":5220
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":5672
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":5F4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":6826
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":7100
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":79DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":82B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":8B8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":9468
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":9D42
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":A61C
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":AEF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":B7D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":C0AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":C984
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":D25E
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":DB38
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":E412
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":ECEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":F5C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":FEA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":1077A
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":11054
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":1192E
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":12208
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":12AE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":133BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":13C96
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":14570
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":14E4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":15724
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":15FFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":168D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":171B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":17A8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":18366
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":18C40
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":1951A
            Key             =   "book"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":19834
            Key             =   "news"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":19B4E
            Key             =   "filter"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":19E68
            Key             =   "world"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":1A182
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":1A5D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":1AA26
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":1AE78
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":1B2CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":1B71C
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":1BA36
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":1BE88
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":1C2DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":1C72C
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":1CB7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":1CFD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":1D422
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":1D874
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":1DCC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":1E118
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":1E56A
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":1E9BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":1EE0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":1F260
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":1F6B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":1FB04
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":1FF56
            Key             =   ""
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":206A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":20AFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":20F4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":2139E
            Key             =   ""
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":217F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":21C42
            Key             =   ""
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":22094
            Key             =   ""
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccHoldings2_new.frx":224E6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer mTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   9000
      Top             =   30
   End
   Begin VB.PictureBox picListview 
      BorderStyle     =   0  'None
      Height          =   2805
      Index           =   0
      Left            =   3360
      ScaleHeight     =   2805
      ScaleWidth      =   5925
      TabIndex        =   4
      Top             =   900
      Width           =   5925
      Begin MSComctlLib.ProgressBar PB1 
         Height          =   135
         Left            =   60
         TabIndex        =   5
         Top             =   2940
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
      End
      Begin MSComctlLib.ListView lvAccountHoldings 
         Height          =   2385
         Left            =   60
         TabIndex        =   6
         Top             =   60
         Width           =   9345
         _ExtentX        =   16484
         _ExtentY        =   4207
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         PictureAlignment=   5
         _Version        =   393217
         ForeColor       =   16777215
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "AccountName"
            Text            =   "Account Name"
            Object.Width           =   5644
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "sfCycle_Upload"
            Text            =   "Upload"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "sfCycle_Download"
            Text            =   "Download"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "sfStartTime"
            Text            =   "Last Logged In"
            Object.Width           =   2540
         EndProperty
         Picture         =   "frmAccHoldings2_new.frx":22938
      End
   End
   Begin VB.PictureBox picTreeView 
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   -60
      ScaleHeight     =   2775
      ScaleWidth      =   3255
      TabIndex        =   2
      Top             =   900
      Width           =   3255
      Begin MSComctlLib.TreeView tvAccountType 
         Height          =   2685
         Left            =   60
         TabIndex        =   3
         Top             =   90
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   4736
         _Version        =   393217
         Indentation     =   441
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         ImageList       =   "ilTreeview"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox picResize 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   3240
      MousePointer    =   9  'Size W E
      ScaleHeight     =   2775
      ScaleWidth      =   75
      TabIndex        =   1
      Top             =   900
      Width           =   75
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   450
      Left            =   8806
      TabIndex        =   8
      Top             =   210
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   794
      _Version        =   393216
      Value           =   10
      BuddyControl    =   "txtRefreshMin"
      BuddyDispid     =   196611
      OrigLeft        =   8730
      OrigTop         =   120
      OrigRight       =   8970
      OrigBottom      =   495
      Max             =   20
      Min             =   1
      SyncBuddy       =   -1  'True
      Wrap            =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Number of Minutes to Refresh:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4110
      TabIndex        =   9
      Top             =   270
      Width           =   3825
   End
   Begin VB.Label Title 
      AutoSize        =   -1  'True
      Caption         =   "Accounts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   1320
   End
   Begin VB.Image Image1 
      Height          =   780
      Left            =   60
      Picture         =   "frmAccHoldings2_new.frx":3041C
      Stretch         =   -1  'True
      Top             =   90
      Width           =   840
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   -60
      X2              =   9720
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "frmAccHoldings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Unloading As Boolean

Dim bActive As Boolean

Dim mButtonDown As Boolean
Dim iLastPoint As POINTAPI
Dim LastMovement As POINTAPI

Public Function DoItemCount()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "DoItemCount"
    Const ContainerName = "frmAccHoldings"
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
    Exit Function

    Dim rsload As adodb.Recordset
    Dim bResult As Boolean
    
    If Me.tvAccountType.NodeS.Count > 0 Then
        Dim lx As Long
        Dim sMySQLCount As String
        
        For lx = 1 To Me.tvAccountType.NodeS.Count
            Select Case Left(Me.tvAccountType.NodeS(lx).Key, 1)
            Case "v", "V"
                If Unloading = True Then Exit Function
                sMySQLCount = MySQL.virtualisp("select distinct count(*) from accountinfo, acci_services Where accountinfo.RecID = acci_services.acci_RecID And accountinfo.VirtualID = " & Mid(Me.tvAccountType.NodeS(lx).Key, 2), "accountinfo")
                If Unloading = True Then Exit Function
                If MySQL.OpenTable(directConn, rsload, , sMySQLCount) = True Then
                    Me.tvAccountType.NodeS(lx).Text = Me.tvAccountType.NodeS(lx).Text + " [" & rsload("Count(*)") & "]"
                    gSleep
                End If
            Case "s", "S"
            Case Else
                If Unloading = True Then Exit Function
                bResult = MySQL.OpenTable(directConn, rsload, , "select CountStatement from accountviewer where RecID = " & Mid(Me.tvAccountType.NodeS(lx).Key, 2))
                If Unloading = True Then Exit Function
                gSleep
                sMySQLCount = rsload!CountStatement
                'sMySQLCount = MySQL.virtualisp(sMySQLCount)
                If MySQL.OpenTable(directConn, rsload, , sMySQLCount) = True Then
                    Me.tvAccountType.NodeS(lx).Text = Me.tvAccountType.NodeS(lx).Text + " [" & rsload("Count(*)") & "]"
                    gSleep
                End If
            End Select
            If Unloading = True Then Exit Function
            
        Next
    
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

Public Function PopulateList()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "PopulateList"
    Const ContainerName = "frmAccHoldings"
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
    
    Dim NodeX As Node
    Dim NodX As Node
    
    Dim allchecked As Long
    
    For allchecked = 0 To ViSPMAP.Count - 1
    
        If ViSPMAP("r" & allchecked).RecIDb = Login.lVirtualID Then
            If Not ViSPMAP("r" & allchecked) Is Nothing Then Set NodX = tvAccountType.NodeS.Add(, , "r" & ViSPMAP("r" & allchecked).RecIDb, ViSPMAP("r" & allchecked).Desc, 58)
        Else
            If Not ViSPMAP("r" & allchecked) Is Nothing Then Set NodX = tvAccountType.NodeS("r" & ViSPMAP("r" & allchecked).RecIDa)
            If Not ViSPMAP("r" & allchecked) Is Nothing Then Set NodeX = tvAccountType.NodeS.Add(NodX, tvwChild, "r" & ViSPMAP("r" & allchecked).RecIDb, ViSPMAP("r" & allchecked).Desc, 58)
        End If
    
    
    Next
    
    
    For allchecked = 1 To NDECHRT.Count
        Set NodX = tvAccountType.NodeS(NDECHRT(allchecked).Key)
        If NDECHRT(allchecked).Key = "r" & Login.lVirtualID Then
        
            Set NodeX = tvAccountType.NodeS.Add(, , NDECHRT(allchecked).Tag, NDECHRT(allchecked).Description, NDECHRT(allchecked).IconNo)
            
        ElseIf NDECHRT(allchecked).Key = "troot" Then
        
            Set NodX = tvAccountType.NodeS.Add(, , NDECHRT(allchecked).Tag, NDECHRT(allchecked).Description, NDECHRT(allchecked).IconNo)
                
        Else
            If Left(NDECHRT(allchecked).Tag, 6) = "TMPSVR" Then
                Set NodX = tvAccountType.NodeS(NDECHRT(NDECHRT(allchecked).SubofRecID).Tag)
            ElseIf LCase(Left(NDECHRT(allchecked).Tag, 1)) = "t" Then
                Set NodX = tvAccountType.NodeS(NDECHRT(NDECHRT(allchecked).SubofRecID).Tag)
            End If
            Set NodeX = tvAccountType.NodeS.Add(NodX, tvwChild, NDECHRT(allchecked).Tag, NDECHRT(allchecked).Description, NDECHRT(allchecked).IconNo)
        End If
    Next
    
    Exit Function
    '---------------------------------------------------------------------------------------------------------------------------------------------------
    'If bDebug = True Then On Error GoTo 0 Else On Error GoTo ErrorOccur

    Dim rsload As adodb.Recordset
    Dim bResult As Boolean
    
    
    bResult = MySQL.OpenTable(directConn, rsload, , "select * from accountviewer where accountviewer.RecID = 1 or accountviewer.RecID = 2 or accountviewer.RecID = 3")
    gSleep
    
    Dim sdKey As String
    
    tvAccountType.NodeS.Clear
    
    If Not tvAccountType.SelectedItem Is Nothing Then
        sdKey = tvAccountType.SelectedItem.Key
    Else
        sdKey = ""
    End If
    Dim rsServices As adodb.Recordset
    
    bResult = MySQL.OpenTable(directConn, rsServices, , "select servicetypes.RecID, servicetypes.ServiceKey, servicetypes.Description, servicetypes.SubofRecID, servicetypes.ListOnRadius, servicetypes.HasUID, servicetypes.HasSysUID, servicetypes.BillImmediately, servicetypes_matrix.SubServiceID  from servicetypes inner join servicetypes_matrix on servicetypes.RecID = servicetypes_matrix.ServiceID and servicetypes_matrix.VirtualID = '" & Login.lVirtualID & "'")
    
    If rsload.RecordCount > 0 Then
        While Not rsload.EOF And Err.Number = 0
            If Not IsNull(rsload!SubofRecID) Then
                If rsload!SubofRecID <> 0 Then
                    If rsload!SubofRecID = rsload!RecID Then
                    
                        Set NodX = tvAccountType.NodeS.Add(, , "r" & rsload!RecID, IIf(IsNull(rsload!Description), "", rsload!Description), Val(rsload!IconNum))
                        NodX.Tag = rsload!Action
                    
                    Else
                    
                        Set NodeX = tvAccountType.NodeS("r" & rsload!SubofRecID)
                        Set NodX = tvAccountType.NodeS.Add(NodeX.Key, tvwChild, "r" & rsload!RecID, IIf(IsNull(rsload!Description), "Description Not set", rsload!Description), Val(rsload!IconNum))
                        'NodeX.Expanded = True
                        rsServices.MoveFirst

                        NodX.Tag = rsload!Action
                    End If
                Else
                    Set NodX = tvAccountType.NodeS.Add(, , "r" & rsload!RecID, IIf(IsNull(rsload!Description), "", rsload!Description), Val(rsload!IconNum))
                    NodX.Tag = rsload!Action
                End If
            Else
                Set NodX = tvAccountType.NodeS.Add(, , "r" & rsload!RecID, IIf(IsNull(rsload!Description), "", rsload!Description), Val(rsload!IconNum))
                NodX.Tag = rsload!Action
            End If
            If Unloading = True Then Exit Function
            rsload.MoveNext
            
        Wend
        
        Set NodX = tvAccountType.NodeS("k1")
        
        While Not rsServices.EOF And Err.Number = 0
            If Unloading = True Then Exit Function
            tvAccountType.NodeS.Add NodX.Key, tvwChild, "s0-" & rsServices!RecID, IIf(IsNull(rsServices!Description), "", rsServices!Description), 23
            If Unloading = True Then Exit Function
            rsServices.MoveNext
        Wend
    End If
   
    
    
    Dim NodeY As Node
    Dim rsVirtual As adodb.Recordset
    Dim rsPlans As adodb.Recordset
    
    bResult = MySQL.OpenTable(directConn, rsPlans, , "select * from plantypes")
    'MySQL.virtualisp("select virtualisp.RecID as VirtualID ,accountviewer.selectStatement ,accountviewer.CountStatement ,accountviewer.Description ,accountviewer.SubofRecID,accountviewer.RecID ,accountviewer.IconNum from accountviewer where accountviewer.RecID <> '1' and accountviewer.RecID <> '2' and accountviewer.RecID <> '3'", "accountviewer") + " GROUP BY virtualisp.RecID, accountviewer.SubofRecID Order By accountviewer.SubofRecID"
    bResult = MySQL.OpenTable(directConn, rsload, , "select * from virtualisp, accountviewer where accountviewer.VirtualID = virtualisp.RecID and virtualisp.VirtualID = '" & Login.lVirtualID & "' or virtualisp.RecID = '" & Login.lVirtualID & "'")
    gSleep
    
    On Error Resume Next
    
    If rsload.State = adStateClosed Then Exit Function
    Dim iLnCount As Double
    Dim iStartCount As Double
    
    If rsload.RecordCount > 0 Then
    
        Do
            iStartCount = iStartCount + 1
            rsload.MoveNext
            If Not IsNull(rsload!SubofRecID) Then
                If rsload!SubofRecID = 1 Then Exit Do
            End If
        Loop Until rsload.EOF Or Err.Number <> 0
        
        While Not rsload.EOF And Err.Number = 0
            If Not IsNull(rsload!SubofRecID) Then
                iLnCount = iLnCount + 1
                If rsload!SubofRecID = 1 And rsload!VirtualID <> 1000 Then
                    If tvAccountType.NodeS("v" & rsload!VirtualID) Is Nothing Then
                    
                        bResult = MySQL.OpenTable(directConn, rsVirtual, , "select * from virtualisp where RecID = " & rsload!VirtualID)
                        Set NodeX = tvAccountType.NodeS.Add(, , "v" & rsload!VirtualID, IIf(IsNull(rsVirtual!Description), "", rsVirtual!Description), 20)
                        rsServices.MoveFirst
                        While Not rsServices.EOF And Err.Number = 0
                            tvAccountType.NodeS.Add NodeX.Key, tvwChild, "s" & rsVirtual!RecID & "-" & rsServices!RecID, IIf(IsNull(rsServices!Description), "", rsServices!Description), 23
                            If Unloading = True Then Exit Function
                            rsServices.MoveNext
                            If Unloading = True Then Exit Function
                        Wend
                        Set NodeX = tvAccountType.NodeS("v" & rsload!VirtualID)
                    Else
                        Set NodeX = tvAccountType.NodeS("v" & rsload!VirtualID)
                    End If
                Else
                    Set NodeX = tvAccountType.NodeS("r" & rsload!SubofRecID)
                End If
                If rsload!SubofRecID = 1 Then
                    If InStr(rsload!selectStatement, "ptRecID") > 0 Then
                        rsPlans.Filter = "RecID = " & Mid(rsload!selectStatement, InStr(rsload!selectStatement, "ptRecID = ") + 10, InStr(rsload!selectStatement, "ptRecID = ") + 10 - InStr(InStr(rsload!selectStatement, "ptRecID = ") + 10, rsload!selectStatement, " "))
                        Set NodeX = tvAccountType.NodeS("s" & rsload!VirtualID & "-" & rsPlans!ServiceID)
                    End If
                    Set NodX = tvAccountType.NodeS.Add(NodeX.Key, tvwChild, "r" & rsload!RecID, IIf(IsNull(rsload!Description), "Description Not set", rsload!Description), Val(rsload!IconNum))
                    NodX.Tag = rsload!Action
                ElseIf rsload!SubofRecID <> 1 Then
                    Set NodeX = tvAccountType.NodeS("r" & rsload!SubofRecID)
                    Set NodX = tvAccountType.NodeS.Add(NodeX.Key, tvwChild, "r" & rsload!RecID, IIf(IsNull(rsload!Description), "Description Not set", rsload!Description), Val(rsload!IconNum))
                    NodX.Tag = rsload!Action
                Else
                    Set NodX = tvAccountType.NodeS.Add(, , "r" & rsload!RecID, IIf(IsNull(rsload!Description), "", rsload!Description), Val(rsload!IconNum))
                    NodX.Tag = rsload!Action
                End If
            Else
                Set NodX = tvAccountType.NodeS.Add(, , "r" & rsload!RecID, IIf(IsNull(rsload!Description), "", rsload!Description), Val(rsload!IconNum))
                NodX.Tag = rsload!Action
            End If
            If Unloading = True Then Exit Function
            rsload.MoveNext
        Wend
        
        Call DoItemCount
    End If
    
    On Error Resume Next
    
    If sdKey <> "" Then
        Call tvAccountType_NodeClick(tvAccountType.NodeS(sdKey))
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




Private Sub Command1_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Command1_Click"
    Const ContainerName = "frmAccHoldings"
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


'    Call ExportHTML
    
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

Private Sub Form_Activate()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Activate"
    Const ContainerName = "frmAccHoldings"
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


    bActive = True
       
   
    
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

Private Sub Form_Deactivate()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Deactivate"
    Const ContainerName = "frmAccHoldings"
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


    bActive = False
    
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
    Const ContainerName = "frmAccHoldings"
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

   
    'Create the thread
    'ThreadControl1.CreateNewThread AddressOf DoItemCount, tpBelowNormal, True

    Me.Caption = "Account Holdings - [ " & Login.sUsername & " ]"
    
    Call MySQL.SetColumnHeaders("accinfo", lvAccountHoldings, "", directConn)
        
    Call GUI.LoadColWidths(lvAccountHoldings, Me)
    
    If bBigFont = True Then
        tvAccountType.Font.Size = 15
        lvAccountHoldings.Font.Size = 16
    End If
    
    
   'mnuPopup_ImportCSV.Enabled = Login.bMaster
    
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
    Const ContainerName = "frmAccHoldings"
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


    Unloading = True
    Call GUI.SaveColWidths(lvAccountHoldings, Me)
    
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
    Const ContainerName = "frmAccHoldings"
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

    If Me.WindowState <> vbMinimized Then
    
    If Me.ScaleHeight > 200 And Me.ScaleWidth > 3000 Then
        'tsAcc.Move 60, tsAcc.Top, Me.ScaleWidth - 180, Me.ScaleHeight - 180 - tsAcc.Top
        
        'picListview(0).Move
        picTreeView.Height = Me.ScaleHeight - picTreeView.Top - 60
        picListview(0).Width = IIf(Me.Width - picResize.Left - picResize.Width - 120 < 0, 10, Me.Width - picResize.Left - picResize.Width - 120)
        picListview(0).Height = Me.ScaleHeight - picListview(0).Top - 60
        picResize.Height = Me.ScaleHeight - picResize.Top
        Line1.X2 = Me.ScaleWidth
        txtRefreshMin.Move Me.ScaleWidth - txtRefreshMin.Width - UpDown1.Width - 60
        UpDown1.Move txtRefreshMin.Left + txtRefreshMin.Width
        Label1.Move txtRefreshMin.Left - Label1.Width - 60
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

Private Sub Form_Unload(Cancel As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Unload"
    Const ContainerName = "frmAccHoldings"
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


    Unloading = True
    
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

Private Sub lvAccountHoldings_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvAccountHoldings_ColumnClick"
    Const ContainerName = "frmAccHoldings"
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


    Call GUI.ColumnSort(ColumnHeader, lvAccountHoldings)
    
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

Private Sub lvAccountHoldings_DblClick()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvAccountHoldings_DblClick"
    Const ContainerName = "frmAccHoldings"
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

    If lvAccountHoldings.SelectedItem Is Nothing Then Exit Sub
    
    Dim ffrmCustomerRec As frmCustomerRec
    Set ffrmCustomerRec = New frmCustomerRec
    
    Screen.MousePointer = vbHourglass
    
    ffrmCustomerRec.osub.fRecID = CLng(Mid(lvAccountHoldings.SelectedItem.Key, 2))
    
    ffrmCustomerRec.Show
    
    Screen.MousePointer = vbDefault

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

Private Sub lvAccountHoldings_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvAccountHoldings_ItemClick"
    Const ContainerName = "frmAccHoldings"
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


    lvAccountHoldings.Tag = True
    
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

Private Sub mnuDeleteNode_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuDeleteNode_Click"
    Const ContainerName = "frmAccHoldings"
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


    Select Case Left(tvAccountType.SelectedItem.Key, 1)
    Case "k"
        
        Call MySQL.Execute(directConn, "delete from accountviewer where RecID = " & Mid(tvAccountType.SelectedItem.Key, 2))
        Call MySQL.Execute(directConn, "delete from accountviewer where SubofRecID = " & Mid(tvAccountType.SelectedItem.Key, 2))
    
        tvAccountType.NodeS.Remove tvAccountType.SelectedItem.Index
        
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

Private Sub mnuPopup_ImportCSV_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuPopup_ImportCSV_Click"
    Const ContainerName = "frmAccHoldings"
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


    Dim fImport As New frmIMport
    
    fImport.Show 1
    
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

Private Sub mnuPopup_SQL_Click(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuPopup_SQL_Click"
    Const ContainerName = "frmAccHoldings"
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

    Dim fRename As New frmRename
    Dim rsload As adodb.Recordset
    
    Select Case Left(tvAccountType.SelectedItem.Key, 1)
    Case "k", "K"
        If MySQL.OpenTable(directConn, rsload, , "select * from accountviewer where RecID = " & Mid(tvAccountType.SelectedItem.Key, 2)) = True Then
        
            Dim oSQL As New frmSQL
            
            Select Case Index
            Case 1
                oSQL.SQL = rsload!selectStatement
            Case 0
                oSQL.SQL = rsload!CountStatement
            End Select
            
        
            oSQL.Show 1
            
            If oSQL.SQL = "" Then Exit Sub
            
            Select Case Index
            Case 1
                MySQL.Execute directConn, "update accountviewer set selectStatement = '" & MySQL.ESC(oSQL.SQL) & "' where RecID = " & Mid(tvAccountType.SelectedItem.Key, 2)
            Case 0
                MySQL.Execute directConn, "update accountviewer set CountStatement = '" & MySQL.ESC(oSQL.SQL) & "' where RecID = " & Mid(tvAccountType.SelectedItem.Key, 2)
            End Select
            
        End If
    End Select
    

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

Private Sub mnuRename_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuRename_Click"
    Const ContainerName = "frmAccHoldings"
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

    Dim fRename As New frmRename
    Dim rsload As adodb.Recordset
    
    If MySQL.OpenTable(directConn, rsload, , "select * from accountviewer where RecID = " & Mid(tvAccountType.SelectedItem.Key, 2)) = True Then
    
        fRename.txt = rsload!Description
        
    
        fRename.Show 1
        
        If fRename.txt = "" Then Exit Sub
        
        If Not fRename.txt = tvAccountType.SelectedItem.Text Then
            Select Case Left(tvAccountType.SelectedItem.Key, 1)
            Case "k", "K"
                MySQL.Execute directConn, "update accountviewer set description = '" & MySQL.ESC(fRename.txt) & "' where RecID = " & Mid(tvAccountType.SelectedItem.Key, 2)
                tvAccountType.SelectedItem.Text = fRename.txt
            End Select
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

Private Sub lvAccountHoldings_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
    
    If Not lvAccountHoldings.SelectedItem Is Nothing Then
        Set frmMDIMain.itmX_AccHldings = lvAccountHoldings.SelectedItem
        frmMDIMain.mnuDrop_AccHld_lv_Name.Caption = frmMDIMain.itmX_AccHldings.SubItems(1)
        PopupMenu frmMDIMain.mnuDrop_AccHld_lv
    End If
    
    End If
End Sub

Private Sub mTimer_Timer()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mTimer_Timer"
    Const ContainerName = "frmAccHoldings"
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
    
    picTreeView.Width = picResize.Left
    picListview(0).Width = IIf(Me.Width - picResize.Left - picResize.Width - 120 < 0, 10, Me.Width - picResize.Left - picResize.Width - 120)
    picListview(0).Left = picResize.Left + picResize.Width
    
    LastMovement.X = (Point.X - iLastPoint.X)
    LastMovement.Y = (Point.Y - iLastPoint.Y)
        
    gSleep


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

Private Sub picListview_Resize(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "picListview_Resize"
    Const ContainerName = "frmAccHoldings"
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


    If picListview(0).Height > 300 And picListview(0).Width > 300 Then
        lvAccountHoldings.Move 60, 60, picListview(0).Width - 120, picListview(0).Height - pb1.Height - 180
        pb1.Move 60, lvAccountHoldings.Top + lvAccountHoldings.Height + 60, picListview(0).Width - 120
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
    Const ContainerName = "frmAccHoldings"
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



Public Function SaveColumnWidths()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "SaveColumnWidths"
    Const ContainerName = "frmAccHoldings"
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


    Call GUI.SaveColWidths(lvAccountHoldings, Me)

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

Public Function LoadColumnWidths()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "LoadColumnWidths"
    Const ContainerName = "frmAccHoldings"
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


    Call GUI.LoadColWidths(lvAccountHoldings, Me)

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


Private Sub picResize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "picResize_MouseUp"
    Const ContainerName = "frmAccHoldings"
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

Private Sub picTreeView_Resize()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "picTreeView_Resize"
    Const ContainerName = "frmAccHoldings"
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


    If picTreeView.Width > 180 And picTreeView.Height > 180 Then
        tvAccountType.Move 60, 60, picTreeView.Width - 120, picTreeView.Height - 120
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

Private Sub tmrRefresh_Timer()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "tmrRefresh_Timer"
    Const ContainerName = "frmAccHoldings"
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


    Static lCount As Long
    
    lCount = lCount + 1
    
    Dim itmX As ListItem
    
    If lCount / 60 >= Val(txtRefreshMin.Text) Then
        lCount = 0
    Else
        Me.Caption = "Account Holdings - [ " & Login.sUsername & " ] (" & Val(txtRefreshMin.Text) * 60 - lCount & " seconds till refresh)"
    End If
    
    If lCount = 1 Then
        frmMDIMain.Caption = "(Refreshing)"
        
            tvAccountType.NodeS.Clear
            PopulateList
        
        Screen.MousePointer = vbDefault
        Me.Caption = "Account Holdings - [ " & Login.sUsername & " ]"
        frmMDIMain.Caption = "The Nexus"
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

Private Sub tvAccountType_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "tvAccountType_MouseDown"
    Const ContainerName = "frmAccHoldings"
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

    
'    If Button = 2 Then
'
'        If tvAccountType.SelectedItem Is Nothing Or (Login.bMaster = False Or Login.bVISPPrimary = True) Then
'            mnuRename.Enabled = False
'        Else
'            mnuRename.Enabled = True
'        End If
'
'        mnuDeleteNode.Enabled = Login.bCreateSysop
'        mnuPopup_SQL(0).Enabled = False
'        mnuPopup_SQL(1).Enabled = False
'        mnuPopup_Icon.Enabled = False
'
'
'        Select Case Left(tvAccountType.SelectedItem.Key, 1)
'        Case "k"
'
'            mnuPopup_SQL(0).Enabled = Login.bMaster
'            mnuPopup_SQL(1).Enabled = Login.bMaster
'            mnuPopup_Icon.Enabled = True
'
'        End Select
'
'        PopupMenu mnuPopup
'    End If
'
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

Private Sub tvAccountType_NodeClick(ByVal Node As MSComctlLib.Node)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "tvAccountType_NodeClick"
    Const ContainerName = "frmAccHoldings"
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

    On Error Resume Next
    Dim itmX As ListItem
    Dim X As Variant
    Dim lLimitCount As Variant
    
    Dim rsload As adodb.Recordset
    Dim sMySQLCount As String
    Dim sMysql As String
    Dim bResult As Boolean
    Dim lTotalRecords  As Variant
    Dim iRecPos As Long
    
    Dim clHdr As ColumnHeader
    Dim iNodeIndx As Long
    
    
    
    iNodeIndx = NDECHRT.FindKey(Node.Key)
    
    Screen.MousePointer = vbHourglass
    If iNodeIndx > 0 Or Left(Node.Key, 1) = "r" Then
    
            If Left(Node.Key, 1) = "r" Then
                sMysql = "select * from accountinfo where VirtualID = " & Mid(Node.Key, 2)
            Else
                sMySQLCount = NDECHRT(iNodeIndx).CountStatement
                sMysql = NDECHRT(iNodeIndx).selectStatement
            End If
            
            lvAccountHoldings.ListItems.Clear
            
                       
            Err.Clear
            bResult = MySQL.OpenTable(directConn, rsload, , sMysql)
            
            Call MySQL.fillLV(directConn, rsload, lvAccountHoldings, False)
    Else
    
    
        
    End If
    Screen.MousePointer = vbDefault
    Err.Clear
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


Public Function ExportHTML(sDir As String) As Boolean


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "ExportHTML"
    Const ContainerName = "frmAccHoldings"
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


    Dim HTML As Variant
    
    HTML = HTML = ""
    
    HTML = HTML = "" & vbCrLf & "<HTML>"
    HTML = HTML = "" & vbCrLf & "<HEAD> <TITLE>"
    HTML = HTML = "" & vbCrLf & "Generated by The Nexus"
    HTML = HTML = "" & vbCrLf & "</TITLE>"
    HTML = HTML = "" & vbCrLf & "<style>"
    HTML = HTML = "" & vbCrLf & " .header {font-family:Helvetica,Arial; font-size:8pt; color:#000000;cursor:hand;}"
    HTML = HTML = "" & vbCrLf & " A.header:link {COLOR: #000000; text-decoration: none}"
    HTML = HTML = "" & vbCrLf & " A.header:active {COLOR: #003399; text-decoration: none}"
    HTML = HTML = "" & vbCrLf & " A.header:visited {COLOR: #000000; text-decoration: none}"
    HTML = HTML = "" & vbCrLf & " A.header:hover {COLOR: #003399; text-decoration: underline}"
    HTML = HTML = "" & vbCrLf & " .content {font-family:Helvetica,Arial; font-size:8pt; color:#000000;padding-left:2;}"
    HTML = HTML = "" & vbCrLf & " A.content:link {COLOR: #000000; text-decoration: none}"
    HTML = HTML = "" & vbCrLf & " A.content:active {COLOR: #003399; text-decoration: none}"
    HTML = HTML = "" & vbCrLf & " A.content:visited {COLOR: #000000; text-decoration: none}"
    HTML = HTML = "" & vbCrLf & " A.content:hover {COLOR: #003399; text-decoration: underline}"
    
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
