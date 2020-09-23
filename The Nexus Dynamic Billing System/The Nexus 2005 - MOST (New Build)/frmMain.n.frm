VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "cswsk32.ocx"
Begin VB.MDIForm frmMDIMain 
   BackColor       =   &H00B4E9F8&
   Caption         =   "The Nexus"
   ClientHeight    =   12330
   ClientLeft      =   2925
   ClientTop       =   1755
   ClientWidth     =   16215
   Icon            =   "frmMain.n.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin SocketWrenchCtrl.Socket sIRC 
      Left            =   180
      Top             =   6360
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
      Binary          =   0   'False
      Blocking        =   -1  'True
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   -1  'True
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.Timer tmrDBMapper 
      Enabled         =   0   'False
      Interval        =   66
      Left            =   330
      Top             =   6900
   End
   Begin MSComctlLib.ImageList ilToolbar 
      Left            =   1230
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   31
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.n.frx":030A
            Key             =   "blank"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.n.frx":0BE4
            Key             =   "receive"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.n.frx":14BE
            Key             =   "invoice"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.n.frx":17D8
            Key             =   "barcode"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.n.frx":1C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.n.frx":1F44
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.n.frx":225E
            Key             =   "graphs"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.n.frx":2578
            Key             =   "tools"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.n.frx":2E52
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.n.frx":372C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.n.frx":4006
            Key             =   "commission"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.n.frx":48E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.n.frx":55BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.n.frx":5E94
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.n.frx":61AE
            Key             =   "nightbackup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.n.frx":64C8
            Key             =   "know1"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.n.frx":67E2
            Key             =   "know2"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.n.frx":6AFC
            Key             =   "know3"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.n.frx":6E16
            Key             =   "know4"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.n.frx":6F70
            Key             =   "knowledgebase"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.n.frx":728A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.n.frx":75A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.n.frx":78BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.n.frx":7BD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.n.frx":7EF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.n.frx":820C
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.n.frx":8526
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.n.frx":8840
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.n.frx":8B5A
            Key             =   "accounts"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.n.frx":9834
            Key             =   "customer"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.n.frx":A10E
            Key             =   "auth"
         EndProperty
      EndProperty
   End
   Begin VB.Timer ircDataRecieved 
      Interval        =   2
      Left            =   210
      Top             =   7710
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   90
      Top             =   2220
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AccessType      =   1
      Protocol        =   4
      URL             =   "http://"
   End
   Begin VB.PictureBox picBg1 
      Align           =   1  'Align Top
      Height          =   525
      Left            =   0
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1077
      TabIndex        =   11
      Top             =   870
      Visible         =   0   'False
      Width           =   16215
      Begin VB.PictureBox picBg 
         AutoRedraw      =   -1  'True
         Height          =   345
         Left            =   3690
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   12
         Top             =   90
         Width           =   315
      End
   End
   Begin VB.Timer timTimeout 
      Interval        =   15000
      Left            =   330
      Top             =   5670
   End
   Begin VB.Timer tmrPopulateKeys 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   330
      Top             =   5160
   End
   Begin VB.Timer tmrStatusPanel 
      Interval        =   1000
      Left            =   360
      Top             =   4680
   End
   Begin VB.Timer tmrIRC 
      Interval        =   3500
      Left            =   360
      Top             =   3090
   End
   Begin VB.Timer sDebugTimeout 
      Interval        =   500
      Left            =   360
      Top             =   3690
   End
   Begin VB.Timer sysTime 
      Interval        =   100
      Left            =   390
      Top             =   2580
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   16215
      TabIndex        =   9
      Top             =   10830
      Width           =   16215
      Begin VB.TextBox txtDebug 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   1005
         Left            =   30
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   10
         Top             =   60
         Width           =   12495
      End
   End
   Begin VB.Timer tmSchedule 
      Interval        =   500
      Left            =   450
      Top             =   2130
   End
   Begin VB.Timer mTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   360
      Top             =   4170
   End
   Begin VB.PictureBox picColumn 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   9435
      Left            =   2250
      ScaleHeight     =   9435
      ScaleWidth      =   13965
      TabIndex        =   2
      Top             =   1395
      Width           =   13965
      Begin MSComctlLib.ImageList ilSchedule 
         Index           =   1
         Left            =   3510
         Tag             =   "File Types & Assocation Icons"
         Top             =   2340
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16777215
         UseMaskColor    =   0   'False
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList ilSchedule 
         Index           =   0
         Left            =   3510
         Tag             =   "File Types & Assocation Icons"
         Top             =   1740
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   22
         ImageHeight     =   22
         MaskColor       =   33023
         UseMaskColor    =   0   'False
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   18
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.n.frx":A9E8
               Key             =   "world"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.n.frx":E1C8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.n.frx":118FE
               Key             =   "black"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.n.frx":14D9B
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.n.frx":18511
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.n.frx":1BC6C
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.n.frx":1F3CF
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.n.frx":22B13
               Key             =   "progress0"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.n.frx":2614E
               Key             =   "progress1"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.n.frx":29818
               Key             =   "progress2"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.n.frx":2CEE0
               Key             =   "progress3"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.n.frx":305B1
               Key             =   "progress4"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.n.frx":33C33
               Key             =   "progress5"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.n.frx":37310
               Key             =   "progress6"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.n.frx":3A9D3
               Key             =   "progress7"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.n.frx":3E07D
               Key             =   "progress8"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.n.frx":416AD
               Key             =   "red"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.n.frx":44BB3
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox picOptions 
         Height          =   9315
         Index           =   0
         Left            =   9840
         ScaleHeight     =   9255
         ScaleWidth      =   4215
         TabIndex        =   13
         Tag             =   "<t>IRC<t/>"
         Top             =   480
         Visible         =   0   'False
         Width           =   4275
         Begin VB.Frame Frame1 
            Caption         =   "Global Chat Session Functions"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   90
            TabIndex        =   22
            Top             =   7980
            Width           =   13485
            Begin VB.Frame Frame2 
               Caption         =   "&Nickname"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   795
               Left            =   150
               TabIndex        =   23
               Top             =   270
               Width           =   4455
               Begin VB.CommandButton Command1 
                  Caption         =   "&Change"
                  Height          =   375
                  Left            =   3090
                  TabIndex        =   25
                  Top             =   300
                  Width           =   1215
               End
               Begin VB.TextBox txtIRCNick 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   120
                  TabIndex        =   24
                  Top             =   300
                  Width           =   2895
               End
            End
         End
         Begin VB.PictureBox picIRCChan 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6465
            Index           =   0
            Left            =   120
            ScaleHeight     =   6465
            ScaleWidth      =   13335
            TabIndex        =   14
            Top             =   120
            Width           =   13335
            Begin VB.TextBox txtIRC 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Terminal"
                  Size            =   9
                  Charset         =   255
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   5355
               Index           =   0
               Left            =   90
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   20
               Top             =   150
               Width           =   10215
            End
            Begin VB.Frame fraSendTXT 
               Caption         =   "Text to send to chat channel"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Index           =   0
               Left            =   90
               TabIndex        =   17
               Top             =   5520
               Width           =   10215
               Begin VB.TextBox txtIRCSend 
                  BackColor       =   &H80000010&
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
                  Height          =   465
                  Index           =   0
                  Left            =   120
                  TabIndex        =   19
                  Top             =   270
                  Width           =   8835
               End
               Begin VB.CommandButton cmdSendIRC 
                  Caption         =   "Send"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   465
                  Index           =   0
                  Left            =   9030
                  TabIndex        =   18
                  Top             =   270
                  Width           =   1065
               End
            End
            Begin VB.Frame fraUsers 
               Caption         =   "Users In Channel"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   6345
               Index           =   0
               Left            =   10440
               TabIndex        =   15
               Top             =   30
               Width           =   2805
               Begin MSComctlLib.ListView lvIRCUsers 
                  Height          =   5925
                  Index           =   0
                  Left            =   90
                  TabIndex        =   16
                  Top             =   300
                  Width           =   2595
                  _ExtentX        =   4577
                  _ExtentY        =   10451
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   -1  'True
                  HideSelection   =   -1  'True
                  FullRowSelect   =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483633
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "Users"
                     Object.Width           =   4577
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   2
                     SubItemIndex    =   1
                     Text            =   "Key"
                     Object.Width           =   1587
                  EndProperty
               End
            End
         End
         Begin MSComctlLib.TabStrip tsIRCChans 
            Height          =   7875
            Left            =   150
            TabIndex        =   21
            Top             =   150
            Width           =   13455
            _ExtentX        =   23733
            _ExtentY        =   13891
            MultiRow        =   -1  'True
            Placement       =   1
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   1
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "#ProjectAlpha"
                  Object.ToolTipText     =   "This is the main channel, You will find everyone in here."
                  ImageVarType    =   2
               EndProperty
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.PictureBox picOptions 
         BorderStyle     =   0  'None
         Height          =   3165
         Index           =   1
         Left            =   600
         ScaleHeight     =   3165
         ScaleWidth      =   1725
         TabIndex        =   4
         Tag             =   "<t>schedule<t/>"
         Top             =   540
         Width           =   1725
         Begin VB.Timer tmrProgressCylinders 
            Interval        =   35
            Left            =   990
            Top             =   60
         End
         Begin VB.CommandButton cmdRunEvent 
            Height          =   375
            Index           =   1
            Left            =   510
            Picture         =   "frmMain.n.frx":48297
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   60
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.CommandButton cmdRunEvent 
            Height          =   375
            Index           =   0
            Left            =   120
            Picture         =   "frmMain.n.frx":48821
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   60
            Width           =   405
         End
         Begin MSComctlLib.ImageList ilColumnIcons 
            Left            =   2370
            Top             =   1410
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   8
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.n.frx":48DAB
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.n.frx":49345
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.n.frx":498DF
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.n.frx":49E79
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.n.frx":4A413
                  Key             =   "status"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.n.frx":4A72D
                  Key             =   "closed"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.n.frx":4AB7F
                  Key             =   "open"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.n.frx":4AFD1
                  Key             =   "time"
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList ilListitemsIcons 
            Left            =   2370
            Top             =   840
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   12
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.n.frx":4B423
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.n.frx":4B875
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.n.frx":4BE0F
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.n.frx":4C261
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.n.frx":4C6B3
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.n.frx":4CB05
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.n.frx":4CF57
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.n.frx":4D3A9
                  Key             =   "visp"
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.n.frx":4D7FB
                  Key             =   "mapprod"
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.n.frx":4DC4D
                  Key             =   "closed"
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.n.frx":4E09F
                  Key             =   "old"
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.n.frx":4E4F1
                  Key             =   "open"
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView lvSchedule 
            Height          =   11325
            Left            =   120
            TabIndex        =   5
            Top             =   510
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   19976
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            SmallIcons      =   "ilBuf(7)"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   917
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   1
               Text            =   "Time"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Description"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Progress "
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Status Report"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.PictureBox picDrag 
         BackColor       =   &H00E0968F&
         BorderStyle     =   0  'None
         Height          =   16695
         Left            =   -30
         MousePointer    =   9  'Size W E
         ScaleHeight     =   16695
         ScaleWidth      =   240
         TabIndex        =   6
         Top             =   60
         Width           =   240
      End
      Begin MSComctlLib.TabStrip tsSchedule 
         Height          =   9885
         Left            =   270
         TabIndex        =   3
         Top             =   120
         Width           =   13935
         _ExtentX        =   24580
         _ExtentY        =   17436
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Scheduled Tasks"
               Object.Tag             =   "<t>schedule<t/>"
               Object.ToolTipText     =   "Here is the task scheduled to occur on your system at the moment."
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Chat/Help Network"
               Object.Tag             =   "<t>IRC<t/>"
               Object.ToolTipText     =   "This is where you can converse with other Project Alpha Sysops"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   16215
      _ExtentX        =   28601
      _ExtentY        =   1535
      ButtonWidth     =   1640
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ilToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New Client"
            Key             =   "AddCust"
            Description     =   "Add Customer"
            Object.ToolTipText     =   "This will add a customer to the database"
            ImageKey        =   "customer"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Scan Sale"
            Key             =   "ScanSale"
            Description     =   "Scan Sale"
            Object.ToolTipText     =   "Click here to scan in a new sale or another scan action."
            ImageKey        =   "barcode"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Resources"
            Key             =   "Resources"
            Description     =   "This is where you can access the online knowledge resource from support log issues and help references."
            Object.ToolTipText     =   "This is where you can access the online knowledge resource from support log issues and help references."
            ImageKey        =   "knowledgebase"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Graphs"
            Key             =   "Graphs"
            Description     =   "This is where you can view reports and graphs for today and previous days."
            Object.ToolTipText     =   "This is where you can view reports and graphs for today and previous days."
            ImageKey        =   "graphs"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Subscribers"
            Key             =   "AccHold"
            Description     =   "Displays your customer/subscriber database. From here you can access your clients records."
            Object.ToolTipText     =   "Displays your customer/subscriber database. From here you can access your clients records."
            ImageKey        =   "accounts"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "The Vault"
            Key             =   "Recieve"
            Description     =   "This is the receivables and pre-payment/vault screen."
            Object.ToolTipText     =   "This is the receivables and pre-payment/vault screen."
            ImageKey        =   "receive"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Invoices"
            Key             =   "State"
            Description     =   "This is the invoice payment section. Here is where you can access the current invoices on your database."
            Object.ToolTipText     =   "This is the invoice payment section. Here is where you can access the current invoices on your database."
            ImageKey        =   "invoice"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Comm Calc"
            Key             =   "comm"
            Description     =   "Commission Calculator"
            Object.ToolTipText     =   "This is where you can access your commission calculator for this session."
            ImageKey        =   "commission"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Relogin"
            Key             =   "Relogin"
            Description     =   "This is where you change the ownership of this session. You need to do this to default sales."
            Object.ToolTipText     =   "Relogin Sysop"
            ImageKey        =   "auth"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Tools"
            Key             =   "tools"
            Description     =   "This is where you can access the tools for this session."
            Object.ToolTipText     =   "This is where you can access the tools for this session."
            ImageKey        =   "tools"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbFooter 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   12045
      Width           =   16215
      _ExtentX        =   28601
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   6165
            MinWidth        =   6174
            Text            =   "Server Time: "
            TextSave        =   "Server Time: "
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Local Time:"
            TextSave        =   "Local Time:"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11271
         EndProperty
      EndProperty
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
   Begin MSWinsockLib.Winsock sockIRC 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuAction 
      Caption         =   "Actions && Systems"
      Begin VB.Menu mnuNewCust 
         Caption         =   "Add New Customer"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuAction_whois 
         Caption         =   "&Whois"
      End
      Begin VB.Menu mnuUpdates 
         Caption         =   "Check for Updates"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuaction_Sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenDoc 
         Caption         =   "Open Document Browser"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuAction_Comm 
         Caption         =   "Communication"
         Begin VB.Menu mnuAction_Comm_feature 
            Caption         =   "Send Feature Request"
         End
         Begin VB.Menu mnuAction_Comm_Bug 
            Caption         =   "Send Bug Report"
         End
      End
      Begin VB.Menu mnuAction_sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAction_Money 
         Caption         =   "Money Systems"
         Begin VB.Menu mnuAction_Accounts 
            Caption         =   "Display Recievables"
            Shortcut        =   {F6}
         End
         Begin VB.Menu mnuAction_InvoiceSystem 
            Caption         =   "Display Invoice System"
            Shortcut        =   {F7}
         End
         Begin VB.Menu mnuAction_Money_Exp 
            Caption         =   "Display Expenditures System"
            Shortcut        =   {F11}
         End
      End
      Begin VB.Menu mnuAction_AccountList 
         Caption         =   "Display Accounts List"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuAction_VISPS 
         Caption         =   "Display VISP's Configuration"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuAction_Commission 
         Caption         =   "Display Commission Calculator"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuAction_Sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAction_Refund 
         Caption         =   "Process Refund"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuAction_Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAction_Night 
         Caption         =   "Night Maintenance"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOptions_Maintenance 
         Caption         =   "Maintenance"
      End
      Begin VB.Menu mnuAction_Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAction_Relogin 
         Caption         =   "Relogin Sysop"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Settings"
      Begin VB.Menu mnuInt 
         Caption         =   "Internal Configuration"
         Begin VB.Menu mnuInt_Columns 
            Caption         =   "Listview Column Mapping"
         End
         Begin VB.Menu mnuSettings_TAX 
            Caption         =   "Tax"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuSettings_Stationary 
            Caption         =   "Stationary"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuOptions_settings 
         Caption         =   "Options and Sysops"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSettings_Vendors 
         Caption         =   "&Vendors && Affliates"
      End
      Begin VB.Menu mnuOptions_Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSettings_SMTP 
         Caption         =   "SMTP Server Setting"
      End
      Begin VB.Menu mnuOptions_setting_AccountType 
         Caption         =   "Sales Channel - Sales Items Admin"
      End
      Begin VB.Menu mnuSettings_Templates 
         Caption         =   "Sales Channel - Templates Admin"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindows_About 
         Caption         =   "About"
      End
      Begin VB.Menu mnuWindow_Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindow_Back 
         Caption         =   "Redraw Background"
      End
      Begin VB.Menu mnuWindows_Sysops 
         Caption         =   "Sysops"
      End
      Begin VB.Menu mnuWindows_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindow_ServiceStatus 
         Caption         =   "Server Status Bar"
      End
      Begin VB.Menu mnuOptions_View_Debug 
         Caption         =   "Debug Window"
      End
      Begin VB.Menu mnuOptions_View_Tasks 
         Caption         =   "Task Scheduler"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuCust 
      Caption         =   "CutomerFRM"
      Visible         =   0   'False
      Begin VB.Menu mnuCust_lvEmail_Primary 
         Caption         =   "lvEmailAddress"
         Begin VB.Menu mnuCust_lvEmail_Reciept 
            Caption         =   "Send Generated Reciept"
         End
         Begin VB.Menu mnuCust_lvEmail_Statement 
            Caption         =   "Send Generated Statement"
         End
         Begin VB.Menu mnuCust_lvEmail_Invoice 
            Caption         =   "Send Generated Invoice"
         End
      End
      Begin VB.Menu mnuCust_lvPlans 
         Caption         =   "LvPlans"
         Begin VB.Menu mnuCust_lvPlans_ViewPO 
            Caption         =   "View Associated Purchase Order"
         End
         Begin VB.Menu mnuCust_lvPlans_sep3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCust_lvPlans_SetActivation 
            Caption         =   "Set Activation Date"
         End
         Begin VB.Menu mnuCust_lvPlans_Properties 
            Caption         =   "Properties"
         End
         Begin VB.Menu mnuCust_lvPlans_sep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCust_lvPlans_Delete 
            Caption         =   "Delete"
         End
         Begin VB.Menu mnuCust_lvPlans_Password 
            Caption         =   "Change Password"
         End
         Begin VB.Menu mnuCust_lvPlans_Sep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCust_lvPlans_billingDate 
            Caption         =   "Set Billing Date to now"
         End
      End
   End
   Begin VB.Menu mnuSupplier 
      Caption         =   "Supplier"
      Visible         =   0   'False
      Begin VB.Menu mnuSupplier_Manage 
         Caption         =   "Manage Supplier"
      End
      Begin VB.Menu mnuSupplier_Items 
         Caption         =   "Manage Items"
      End
   End
   Begin VB.Menu mnuDrop 
      Caption         =   "Popup Menus"
      Begin VB.Menu mnuDrop_AccHoldings 
         Caption         =   "AccountHoldings"
         Begin VB.Menu mnuDrop_AccHld_lv 
            Caption         =   "Listview"
            Begin VB.Menu mnuDrop_AccHld_lv_Name 
               Caption         =   "Name"
            End
            Begin VB.Menu mnuDrop_AccHld_lv_sep1 
               Caption         =   "-"
            End
            Begin VB.Menu mnuDrop_AccHld_lv_Activate 
               Caption         =   "Activate All New Radius Services for Billing Process"
            End
            Begin VB.Menu mnuDrop_AccHld_lv_Deactivate 
               Caption         =   "Deactivate all Radius Services for Billing Process"
            End
            Begin VB.Menu mnuDrop_AccHld_lv_Cancel 
               Caption         =   "Cancel Customer/Client/Subscriber"
            End
            Begin VB.Menu mnuDrop_AccHld_lv_sep2 
               Caption         =   "-"
            End
            Begin VB.Menu mnuDrop_AccHld_lv_email 
               Caption         =   "Send client an Email"
               Enabled         =   0   'False
            End
            Begin VB.Menu mnuDrop_AccHld_lv_scansale 
               Caption         =   "Intialise Scan Sale for client."
               Enabled         =   0   'False
            End
         End
      End
   End
End
Attribute VB_Name = "frmMDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public itmX_AccHldings As ListItem


Public frmCust As frmCustomerRec
Public bRefresh As Boolean
Public msCount As Byte

Dim IRC As New clsIRC

Dim bMode As Byte

Dim oldnick$    'the backup if the nick change failed
Dim nick$       'our global nickname var
Dim channel$    'our channel var
Dim data$       'the var that will hold the data of a single command
Dim connected As Boolean  'this var will be used to check if we timed out, and will be set to true if get connected


Dim Loading As Boolean

Dim mButtonDown As Boolean
Dim iLastPoint As POINTAPI
Dim LastMovement As POINTAPI

Const GetAll = 1
Const RequestKeys = 2
Const SendKeys = 3
Const ResendKeys = 4
    

Private Sub cmdRunEvent_Click(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdRunEvent_Click"
    Const ContainerName = "frmMDIMain"
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
    
    For ix = lvSchedule.ListItems.Count To 1 Step -1
        If lvSchedule.ListItems(ix).Checked = True Then
            RunBot ix
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

Private Sub lvBuf_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Call GUI.ColumnSort(ColumnHeader, lvBuf)
    
End Sub

Private Sub ircDataRecieved_Timer()

    If sIRC.RecvNext > 36 Then
        Dim tmpBuff As String
        
        sIRC.Read tmpBuff, sIRC.RecvNext
        
        If InStr(tmpBuff, "PING") > 0 Then
            Dim params$    ' parameters that will be filtered from the pong message
            params$ = Right$(data$, Len(data$) - (InStr(data$, "PING") + 4))
            send "PONG " + params$   ' send the pong message to the server, together with the parameters
        Else
            txtIRC(0).Text = txtIRC(0).Text + tmpBuff + vbCrLf
            txtIRC(0).SelStart = Len(txtIRC(0).Text) - Len(tmpBuff + vbCrLf)
        End If
    
    End If
    
End Sub

Private Sub lvSchedule_ItemCheck(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvSchedule_ItemCheck"
    Const ContainerName = "frmMDIMain"
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
    
    For ix = lvSchedule.ListItems.Count To 1 Step -1
        If lvSchedule.ListItems(ix).Index = Item.Index Then
            Item.Checked = True
        Else
            lvSchedule.ListItems(ix).Checked = False
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

Private Sub mnuDrop_AccHld_lv_Activate_Click()

    Dim SQL As String
    
    SQL = "Update acci_services, accountinfo, acci_dslconnections set acci_dslconnections.AccountActive = '1' , acci_services.NextCycle = '" & Format(sysnow, "yyyy-mm-dd ttttt") & "', acci_services.ActivationSet = '1', acci_services.Activation = '" & Format(sysnow, "yyyy-mm-dd ttttt") & "', accountinfo.ActivationDate = '" & Format(sysnow, "yyyy-mm-dd ttttt") & "' where (accountinfo.ActivationDate = '1899-12-30 12:00:00' or accountinfo.ActivationDate is NULL or accountinfo.ActivationDate = '0000-00-00 00:00:00') and (accountinfo.RecID = acci_services.acci_RecID and acci_dslconnections.acci_RecID = acci_services.acci_RecID) and acci_services.acci_RecID = '" & Mid(itmX_AccHldings.Key, 2) & "' and acci_dslconnections.AccountActive <> '1'"
    
    MySQL.Execute directConn, SQL
    
    SQL = "update accountinfo set BillingDate = '" & Format(sysnow, "yyyy-mm-dd ttttt") & "' where BillingDate is NULL or BillingDate = '1899-12-30 12:00:00' or BillingDate = '0000-00-00 00:00:00' and RecID = '" & Mid(itmX_AccHldings.Key, 2) & "'"
    
    MySQL.Execute directConn, SQL
    
End Sub

Private Sub mnuDrop_AccHld_lv_Cancel_Click()

    Dim SQL As String
    SQL = "Update accountinfo set Cancelled = '-1' where RecID = '" & Mid(itmX_AccHldings.Key, 2) & "'"
    
End Sub

Private Sub mnuDrop_AccHld_lv_Deactivate_Click()

    Dim SQL As String
    
    SQL = "Update acci_services, accountinfo, acci_dslconnections set acci_dslconnections.AccountActive = '2' , acci_services.NextCycle = '" & Format(DateAdd("d", 1, sysnow), "yyyy-mm-dd ttttt") & "', acci_services.ActivationSet = '2', acci_services.Activation = NULL, accountinfo.ExpiryDate = '" & Format(sysnow, "yyyy-mm-dd ttttt") & "' where (accountinfo.ActivationDate <> '1899-12-30 12:00:00' or accountinfo.ActivationDate is not NULL or accountinfo.ActivationDate <> '0000-00-00 00:00:00') and (accountinfo.RecID = acci_services.acci_RecID and acci_dslconnections.acci_RecID = acci_services.acci_RecID) and acci_services.acci_RecID = '" & Mid(itmX_AccHldings.Key, 2) & "' and acci_dslconnections.AccountActive = '1'"
    MySQL.Execute directConn, SQL
    
    SQL = "update accountinfo set BillingDate = '" & Format(DateAdd("h", 4, sysnow), "yyyy-mm-dd ttttt") & "' where RecID = '" & Mid(itmX_AccHldings.Key, 2) & "'"
    MySQL.Execute directConn, SQL
    
End Sub

Private Sub mnuDrop_AccHld_lv_Name_Click()

    Dim ffrmCustomerRec As frmCustomerRec
    Set ffrmCustomerRec = New frmCustomerRec
    Screen.MousePointer = vbHourglass
    ffrmCustomerRec.osub.fRecID = CLng(Mid(itmX_AccHldings.Key, 2))
    ffrmCustomerRec.Show
    
End Sub

'
'Private Sub cmdChannel_Click()
'    send "PART " + channel$  'leave the current channel
'    lstUsers.Clear  'clear the user list
'    send "JOIN " + txtChannel.Text  'join the new channel
'    channel$ = txtChannel.Text  'store the current channel
'End Sub
'
'Private Sub cmdChat_Click()
'    If Trim(txtChatMsg.Text) = "" Then Exit Sub  'if there's no message exit the sub
'    If Left(txtChatMsg.Text, 1) = "/" Then
'        send Trim(txtChatMsg.Text)    'send the message to the channel
'    Else
'        send "PRIVMSG " + channel$ + " :" + Trim(txtChatMsg.Text) 'send the message to the channel
'    End If
'    displaychat ">> " + Trim(txtChatMsg.Text)   'display the message
'    txtChatMsg.Text = ""    'clear the field
'    txtChatMsg.SetFocus     'give the focus back to the message field
'End Sub
'
'Private Sub cmdDisconnect_Click()
'    Unload Me   'unload the form
'End Sub
'
'Private Sub cmdNick_Click()
'    oldnick$ = nick$    'make a backup in case the nick change failes
'    display "<!> Changing nickname to " + txtNick.Text  'display a notice
'    nick$ = txtNick.Text    'get the content of the text field
'    send "NICK :" + txtNick.Text    'change the nickname
'End Sub
'
'Private Sub cmdPriv_Click()
'    If Trim(txtPrivMsg.Text) = "" Or Trim(txtTarget.Text) = "" Then Exit Sub    'if there is no target or message, exit
'    displaychat ">> MSG -> " + Trim(txtTarget.Text) + ": " + Trim(txtPrivMsg.Text)
'
'
'
'    txtPrivMsg.Text = ""    'clear the field
'    txtPrivMsg.SetFocus     'give the focus back to the message field
'End Sub
'
'
'Private Sub Label3_Click()
'    Clipboard.Clear     'remove the current content of the clipboard
'    Clipboard.SetText txtStatus.Text    'place the text of the status field on the clipboard
'End Sub
'
'Private Sub Label4_Click()
'    Clipboard.Clear     'remove the current content of the clipboard
'    Clipboard.SetText txtChat.Text    'place the text of the chat field on the clipboard
'End Sub
'
'Private Sub lstUsers_Click()
'    txtTarget.Text = lstUsers.Text  'set the target text to the nick of the one you clicked
'End Sub

Private Sub mnuInt_Columns_Click()

    frmclmLayout.Show
    
End Sub

Private Sub mnuNewCust_Click()

    If Login.bAddCust = False Then
        MsgBox "You do not have permission to create new customers"
        Exit Sub
    End If
    
        Dim RSsUBS As adodb.Recordset
        Dim fCustomerRec As New frmCustomerRec
        
        bResult = MySQL.OpenTable(directConn, RSsUBS, , "select RecID, NoSub, Subscribed  from virtualisp where RecID = " & Login.lVirtualID)
        
        If RSsUBS.State = adStateOpen Then
        
            If RSsUBS.RecordCount > 0 Then
                
                If RSsUBS!NoSub < RSsUBS!Subscribed Then
                    fCustomerRec.Show
                Else
                    RSsUBS!Subscribed = RSsUBS!Subscribed + 100
                    RSsUBS.Update
                    MsgBox "You have reach the maximum amount of customers supported by your licence subscription, we have added another block of 100 users to your Virtual ISP settings.", vbCritical, "User Licences has reached a maximum quota"
                    fCustomerRec.Show
                End If
            Else
                fCustomerRec.Show
            End If
        Else
            fCustomerRec.Show
        End If
        
        
        If fCustomerRec.FormState = Waiting And Screen.MousePointer <> vbDefault Then
            Screen.MousePointer = vbDefault
        End If
        
        DoEvents
        
        Do
            DoEvents
        Loop While fCustomerRec.Visible = False
        

        

End Sub

Private Sub mnuOpenDoc_Click()

    Dim fdoc As New frmDOC
    
    fdoc.Show
    
End Sub

Private Sub mnuSettings_SMTP_Click()

    frmSMTP.Show 1
    
End Sub

Private Sub mnuWindow_Back_Click()

        picBg.Width = (Screen.Width / Screen.TwipsPerPixelX) * 1.4
        picBg.Height = Screen.Height / Screen.TwipsPerPixelY
        
        Dim posX As Long
        Dim SR As Integer
        Dim SG As Integer
        Dim SB As Integer
        Dim ER As Integer
        Dim EG As Integer
        Dim EB As Integer
        
         Randomize Now
         SR = Round(Rnd * 250)
         Randomize Now
         SG = Round(Rnd * 250)
         Randomize Now
         SB = Round(Rnd * 250)
         Randomize Now
         ER = Round(Rnd * 250)
         Randomize Now
         EG = Round(Rnd * 250)
         Randomize Now
         EB = Round(Rnd * 250)
        
        ' white heaven 97582810
        
        Dim colour As Long
        
        Dim ttlLenZ As Long
        ttlLenZ = ((Screen.Width / Screen.TwipsPerPixelX) * 1.4) * 8
        For posX = 1 To ttlLenZ
            colour = RGB(IIf(SR > ER, SR - ((SR - ER) * (posX / (ttlLenZ))), ER - ((ER - SR) * (posX / (ttlLenZ)))), _
                          IIf(SG > ER, SG - ((SG - EG) * (posX / (ttlLenZ))), EG - ((EG - SG) * (posX / (ttlLenZ)))), _
                          IIf(SB > EB, SB - ((SB - EB) * (posX / (ttlLenZ))), EB - ((EB - SB) * (posX / (ttlLenZ)))))
            Me.picBg.Line (posX, 0)-(0, (Screen.Height / Screen.TwipsPerPixelY) * (posX / (ttlLenZ))), colour
        Next
        
        Me.Picture = Me.picBg.Image
        Me.Hide
        Me.Show
        
        
End Sub

Private Sub mnuWindow_ServiceStatus_Click()

    Picture1.Visible = Not mnuWindow_ServiceStatus.Checked
    mnuWindow_ServiceStatus.Checked = Picture1.Visible
    
End Sub

Private Sub picIRCChan_Resize(Index As Integer)

    txtIRC(Index).Move txtIRC(Index).Left, txtIRC(Index).Top, picIRCChan(Index).Width * 0.766029246344207, picIRCChan(Index).Height * 0.82830626450116
    fraSendTXT(Index).Move fraSendTXT(Index).Left, picIRCChan(Index).Height - 60 - (0.132250580046404 * picIRCChan(Index).Height), picIRCChan(Index).Width * 0.766029246344207, 0.132250580046404 * picIRCChan(Index).Height
    txtIRCSend(Index).Move txtIRCSend(Index).Left, txtIRCSend(Index).Top, fraSendTXT(Index).Width - cmdSendIRC(Index).Width - 120, fraSendTXT(Index).Height - txtIRCSend(Index).Top - 80
    cmdSendIRC(Index).Move fraSendTXT(Index).Width - 60 - cmdSendIRC(Index).Width, txtIRCSend(Index).Top, cmdSendIRC(Index).Width, txtIRCSend(Index).Height
    fraUsers(Index).Move picIRCChan(Index).Width - 60 - (picIRCChan(Index).Width * 0.210348706411699), fraUsers(Index).Top, picIRCChan(Index).Width * 0.210348706411699, picIRCChan(Index).Height * 0.981438515081207
    lvIRCUsers(Index).Move lvIRCUsers(Index).Left, lvIRCUsers(Index).Top, fraUsers(Index).Width - lvIRCUsers(Index).Left * 2, fraUsers(Index).Height - lvIRCUsers(Index).Top - 60
    
End Sub

Private Sub sDebugTimeout_Timer()

    Static Msg As String
    Static dtevent As Date
    Static Process As Integer
    
    Static ScreenWidth As Long
    Static ScreenHeight As Long
    
    If Msg <> sbFooter.Panels(4).Text Then
        Msg = sbFooter.Panels(4).Text
        If Not sbFooter.Panels(4).Text = Login.ViSPDesc Then
            Process = Process + 1
            Select Case Process
            Case 1
                If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Play "Processing"
            Case 4
                If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Play "Write"
            Case 7
                If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Play "DoMagic1"
            Case 11
                If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Play "Thinking"
            Case 15
                Process = -2
            End Select
        End If
        dtevent = Now
    End If
    
    If Msg = sbFooter.Panels(4).Text And DateDiff("s", dtevent, sysnow) > 7 Then
        If Not sbFooter.Panels(4).Text = Login.ViSPDesc Then
            sbFooter.Panels(4).Text = Login.ViSPDesc
            
                If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Stop
            
        End If
    End If
    
    gSleep
    Static FirstTime As Boolean
    
    If ScreenWidth <> Screen.Width Or Screen.Height <> ScreenHeight Then
    
        
        
        ScreenWidth = Screen.Width
        ScreenHeight = Screen.Height
        
        picBg.Width = (Screen.Width / Screen.TwipsPerPixelX) * 1.4
        picBg.Height = Screen.Height / Screen.TwipsPerPixelY
        
        If FirstTime = False Then
            FirstTime = True
            Exit Sub
        End If
        
        Me.Visible = False
        
        Dim posX As Long
        Dim SR As Integer
        Dim SG As Integer
        Dim SB As Integer
        Dim ER As Integer
        Dim EG As Integer
        Dim EB As Integer
        
        Randomize Now
         
         SR = Round(Rnd * 250)
         SG = Round(Rnd * 250)
         SB = Round(Rnd * 250)
         ER = Round(Rnd * 250)
         EG = Round(Rnd * 250)
         EB = Round(Rnd * 250)
        
        ' white heaven 97582810
        
        Dim colour As Long
        picBg.Picture = LoadPicture()
        Dim ttlLenZ As Long
        ttlLenZ = ((Screen.Width / Screen.TwipsPerPixelX) * 1.4) * 8
        For posX = 1 To ttlLenZ
            colour = RGB(IIf(SR > ER, SR - ((SR - ER) * (posX / (ttlLenZ))), ER - ((ER - SR) * (posX / (ttlLenZ)))), _
                          IIf(SG > ER, SG - ((SG - EG) * (posX / (ttlLenZ))), EG - ((EG - SG) * (posX / (ttlLenZ)))), _
                          IIf(SB > EB, SB - ((SB - EB) * (posX / (ttlLenZ))), EB - ((EB - SB) * (posX / (ttlLenZ)))))
            picBg.Line (posX, 0)-(0, (Screen.Height / Screen.TwipsPerPixelY) * (posX / (ttlLenZ))), colour
        Next
       
        picBg.Refresh
        Me.Picture = picBg.Image
        
        Me.ZOrder 0
        Me.Visible = True
        
    End If
    
End Sub
'
'Private Sub sockIRC_Close()
'
'    connected = False
'    tmrIRC.Enabled = True
'    tmrIRC.Interval = 100
'
'End Sub
'




Private Sub sIRC_Disconnect()

    connected = False
    
End Sub

Private Sub sIRC_Read(DataLength As Integer, IsUrgent As Integer)

        Dim Temp$
        
        sIRC.Read Temp$, DataLength  'get 1 byte out of the data stream and store it in temp$
        
        'processCommand Temp$
        
        'MsgBox Temp$
            
End Sub

Private Sub timTimeout_Timer()
    If Not (connected) Then
        'cDebug "The connection to the server timed out!"
        
        tmrIRC.Interval = 100

    End If
    gSleep
    
End Sub

Private Sub tmrIRC_Timer()

    Static attempts As Byte
    Static LastServer As Long
    
    If InStr(UCase(Command), "/IRC:OFF") > 0 Then
        tmrIRC.Enabled = False
        tmrIRC.Interval = 0
        Exit Sub
    End If
    
    tmrIRC.Interval = 15000
    
    'On Error Resume Next
    
    If connected = False Then
    
        Dim rsIRC As adodb.Recordset
        Dim iRecPos As Long
        Dim rmtHost As String
        Dim rmtPort As String
        
        gSleep
        
        
         
             rmtHost = Trim("sydney.oz.org")
             rmtPort = "6667"
             nick$ = Login.sUsername
             channel$ = "#projectalpha"
             sIRC.Disconnect
                
            gSleep
                
             If MySQL.OpenTable(directConn, rsIRC, , "select count(*) as RecCount from ircservices") = True Then
                 If rsIRC.State = adStateOpen Then
                     If rsIRC!RecCount > 0 Then
                         Do
                         
                             Select Case Round(Rnd * 5)
                             Case 0 To 1
                                Randomize Now / 384.78 ^ 0.9 * 5.8
                             Case 2 To 3
                                Randomize Now / 38.478 ^ 0.9
                             Case 4 To 5
                                Randomize Now / 38478
                             End Select
                             
                             
                             iRecPos = rsIRC!RecCount - Round(Rnd * rsIRC!RecCount)
                             gSleep
                             If Not LastServer = iRecPos Then Exit Do
                         Loop
                         If iRecPos > 0 Then iRecPos = iRecPos - 1
                         LastServer = iRecPos
                         If MySQL.OpenTable(directConn, rsIRC, , "select * from ircservices " & iRecPos & " where VirtualID = '" & Login.lVirtualID & "' limit " & Round(Rnd * rsIRC.RecordCount) & "', 1'") = True Then
                             If rsIRC.State = adStateOpen Then
                                 If rsIRC.RecordCount > 0 Then
                                     rmtHost = IIf(IsNull(rsIRC!Server), Trim("sydney.oz.org"), "" & rsIRC!Server)
                                     rmtPort = IIf(IsNull(rsIRC!Port), Trim("6667"), "" & Val(rsIRC!Port))
                                     channel$ = IIf(IsNull(rsIRC!channel), Trim("#projectalpha"), "" & rsIRC!channel)
                                 End If
                             End If
                             rsIRC.Close
                         End If
                     End If
                 End If
             End If
             
    
             connected = ircConnect(rmtHost, rmtPort)
             gSleep
             
'            MsgBox connected
            
        
        
    Else
        Static DateRefreshed As Boolean
        Static RefreshOnce As Boolean
        If RefreshOnce = False And DateDiff("s", sysnow, DateRefreshed) > 20 Then
            Me.bRefresh = True
            Me.bRefresh = True
        Else
'            MsgBox DateRefreshed
        End If
    End If
    
End Sub

'Private Sub tmrPopulateKeys_Timer()
'
'
'
'    '*[ Error Checking Variables ]**********************************************************************************
'    Const RoutineName = "MDIForm_Load"
'    Const ContainerName = "processCommand"
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
'
'
'
'    Static lSteps As Long
'    Static bSteps As Byte
'    Static lCount As Long
'    Static lUser As Long
'    Static NumDistro As Long
'
'    Dim A1 As String
'    Dim A2 As String
'    Dim A3 As String
'    Dim A4 As String
'    Dim A5 As String
'    Dim A6 As String
'    Dim A7 As String
'    Dim A8 As String
'    Dim B1 As String
'    Dim B2 As String
'    Dim B3 As String
'    Dim B4 As String
'    Dim B5 As String
'    Dim B6 As String
'    Dim B7 As String
'    Dim B8 As String
'
'
'    Static lstcnt As Long
'    Static LastChange As Date
'
'
'    Select Case lSteps
'    Case -1
'
'    Case 0
'
'        If lstcnt <> lstKeyRqst.ListCount Then
'            LastChange = Now
'            lstcnt = lstKeyRqst.ListCount
'            bMode = RequestKeys
'            tmrPopulateKeys.Interval = 400
'        ElseIf lstKeySend.ListCount = 0 And lvResend.ListItems.Count > 0 Then
'
'            If lvResend.ListItems.Count > 0 Then
'
'                lSteps = 0
'                bMode = ResendKeys
'                tmrPopulateKeys.Interval = 1026
'
'            Else
'
'
'            End If
'
'        Else
'            If DateDiff("s", sysNOW, LastChange) > 15 Or DateDiff("s", sysNOW, LastChange) < 15 Then
'
'                If lstKeySend.ListCount > 0 Then
'                    bMode = SendKeys
'                    lSteps = lstKeySend.ListCount * 18
'                    lCount = 0
'                    NumDistro = 0
'                    lblIRCStat.Caption = "Status: Populating Encryption Matrix - " & lSteps & " steps to do"
'                    tmrPopulateKeys.Interval = 1050
'                End If
'
'            End If
'        End If
'
'        Select Case bMode
'        Case RequestKeys
'            If lstKeyRqst.ListCount > 0 Then
'                Dim dStr As String
'                Dim lx As Long
'                If lstKeyRqst.List(0) <> nick$ Then
'
'
'                    Select Case lstKeyRqst.ItemData(0)
'                    Case 1
'                        send "PRIVMSG " + Trim(lstKeyRqst.List(0)) + " :[DATECRC]<D1>" & Format(GetSetting(App.ProductName, "IRCKey", "Created", DateAdd("d", -1, sysNOW)), "yyyy-mm-dd ttttt") & "</D1><D1>" & Format(GetSetting(App.ProductName, "IRCKey", "Created", DateAdd("d", -1, sysNOW)), "yyyy-mm-dd ttttt") & "</D1><D1>" & Format(GetSetting(App.ProductName, "IRCKey", "Created", DateAdd("d", -1, sysNOW)), "yyyy-mm-dd ttttt") & "</D1><D2>" & Format(sysNOW, "yyyy-mm-dd ttttt") & "</D2>"
'                        dStr = lstKeyRqst.List(0)
'                        lstKeyRqst.RemoveItem 0
'                        lstKeyRqst.AddItem dStr
'                        lstKeyRqst.ItemData(lstKeyRqst.ListCount - 1) = 2
'                        gSleep
'                    Case 2
'                        dStr = lstKeyRqst.List(0)
'                        lstKeyRqst.RemoveItem 0
'                        lstKeyRqst.AddItem dStr
'                        lstKeyRqst.ItemData(lstKeyRqst.ListCount - 1) = 2
'                        gSleep
'                    Case Else
'                        dStr = lstKeyRqst.List(0)
'
'                        For lx = lstKeyRqst.ListCount - 1 To 0 Step -1
'                            If lstKeyRqst.List(lx) = dStr Then
'                                lstKeyRqst.RemoveItem lx
'                            End If
'                        Next
'
'                        lstKeySend.AddItem dStr
'                        'lstKeyRqst.RemoveItem 0
'
'                    End Select
'
'                Else
'                    dStr = lstKeyRqst.List(0)
'
'                    For lx = lstKeyRqst.ListCount - 1 To 0 Step -1
'                        If lstKeyRqst.List(lx) = dStr Then
'                            lstKeyRqst.RemoveItem lx
'                        End If
'                    Next
'
'                End If
'            End If
'        End Select
'
'    Case Else
'        lCount = lCount + 1
'        Select Case bMode
'        Case ResendKeys
'
'            If lvResend.ListItems(0).Text <> nick$ Or lvResend.ListItems(0).SubItems(1) <> nick$ Then
'                Select Case bSteps
'                Case 0
'                     bSteps = bSteps + 1
'                     A1 = IRC.colKeys(IRC.colKeys.FindKey(lvResend.ListItems(0).Text)).A1
'                     send "PRIVMSG " + Trim(lvResend.ListItems(0).SubItems(1)) + " :[IRCKey]<A1>" & MySQL.NumCrypt(A1) & "</A1>" & "<N1>" & lvResend.ListItems(0).Text & "</N1> [PASSON]"
'                Case 1
'                     bSteps = bSteps + 1
'                     A2 = IRC.colKeys(IRC.colKeys.FindKey(lvResend.ListItems(0).Text)).A2
'                     send "PRIVMSG " + Trim(lvResend.ListItems(0).SubItems(1)) + " :[IRCKey]<A2>" & MySQL.NumCrypt(A2) & "</A2>" & "<N1>" & lvResend.ListItems(0).Text & "</N1> [PASSON]"
'                Case 2
'                     bSteps = bSteps + 1
'                     A3 = IRC.colKeys(IRC.colKeys.FindKey(lvResend.ListItems(0).Text)).A3
'                     send "PRIVMSG " + Trim(lvResend.ListItems(0).SubItems(1)) + " :[IRCKey]<A3>" & MySQL.NumCrypt(A3) & "</A3>" & "<N1>" & lvResend.ListItems(0).Text & "</N1> [PASSON]"
'                Case 3
'                     bSteps = bSteps + 1
'                     A4 = IRC.colKeys(IRC.colKeys.FindKey(lvResend.ListItems(0).Text)).A4
'                     send "PRIVMSG " + Trim(lvResend.ListItems(0).SubItems(1)) + " :[IRCKey]<A4>" & MySQL.NumCrypt(A4) & "</A4>" & "<N1>" & lvResend.ListItems(0).Text & "</N1> [PASSON]"
'                Case 4
'                     bSteps = bSteps + 1
'                     A5 = IRC.colKeys(IRC.colKeys.FindKey(lvResend.ListItems(0).Text)).A5
'                     send "PRIVMSG " + Trim(lvResend.ListItems(0).SubItems(1)) + " :[IRCKey]<A5>" & MySQL.NumCrypt(A5) & "</A5>" & "<N1>" & lvResend.ListItems(0).Text & "</N1> [PASSON]"
'                Case 5
'                     bSteps = bSteps + 1
'                     A6 = IRC.colKeys(IRC.colKeys.FindKey(lvResend.ListItems(0).Text)).A6
'                     send "PRIVMSG " + Trim(lvResend.ListItems(0).SubItems(1)) + " :[IRCKey]<A6>" & MySQL.NumCrypt(A6) & "</A6>" & "<N1>" & lvResend.ListItems(0).Text & "</N1> [PASSON]"
'
'                Case 6
'                     bSteps = bSteps + 1
'                     A7 = IRC.colKeys(IRC.colKeys.FindKey(lvResend.ListItems(0).Text)).A7
'                     send "PRIVMSG " + Trim(lvResend.ListItems(0).SubItems(1)) + " :[IRCKey]<A7>" & MySQL.NumCrypt(A7) & "</A7>" & "<N1>" & lvResend.ListItems(0).Text & "</N1> [PASSON]"
'
'                 Case 7
'                     bSteps = bSteps + 1
'                     A8 = IRC.colKeys(IRC.colKeys.FindKey(lvResend.ListItems(0).Text)).A8
'                     send "PRIVMSG " + Trim(lvResend.ListItems(0).SubItems(1)) + " :[IRCKey]<A8>" & MySQL.NumCrypt(A8) & "</A8>" & "<N1>" & lvResend.ListItems(0).Text & "</N1> [PASSON]"
'                Case 8
'                     bSteps = bSteps + 1
'                     B1 = IRC.colKeys(IRC.colKeys.FindKey(lvResend.ListItems(0).Text)).B1
'                     send "PRIVMSG " + Trim(lvResend.ListItems(0).SubItems(1)) + " :[IRCKey]<B1>" & MySQL.NumCrypt(B1) & "</B1>" & "<N1>" & lvResend.ListItems(0).Text & "</N1> [PASSON]"
'                Case 9
'                     bSteps = bSteps + 1
'                     B2 = IRC.colKeys(IRC.colKeys.FindKey(lvResend.ListItems(0).Text)).B2
'                     send "PRIVMSG " + Trim(lvResend.ListItems(0).SubItems(1)) + " :[IRCKey]<B2>" & MySQL.NumCrypt(B2) & "</B2>" & "<N1>" & lvResend.ListItems(0).Text & "</N1> [PASSON]"
'                Case 10
'                     bSteps = bSteps + 1
'                     B3 = IRC.colKeys(IRC.colKeys.FindKey(lvResend.ListItems(0).Text)).B3
'                     send "PRIVMSG " + Trim(lvResend.ListItems(0).SubItems(1)) + " :[IRCKey]<B3>" & MySQL.NumCrypt(B3) & "</B3>" & "<N1>" & lvResend.ListItems(0).Text & "</N1> [PASSON]"
'                Case 11
'                     bSteps = bSteps + 1
'                     B4 = IRC.colKeys(IRC.colKeys.FindKey(lvResend.ListItems(0).Text)).B4
'                     send "PRIVMSG " + Trim(lvResend.ListItems(0).SubItems(1)) + " :[IRCKey]<B4>" & MySQL.NumCrypt(B4) & "</B4>" & "<N1>" & lvResend.ListItems(0).Text & "</N1> [PASSON]"
'                Case 12
'                     bSteps = bSteps + 1
'                     B5 = IRC.colKeys(IRC.colKeys.FindKey(lvResend.ListItems(0).Text)).B5
'                     send "PRIVMSG " + Trim(lvResend.ListItems(0).SubItems(1)) + " :[IRCKey]<B5>" & MySQL.NumCrypt(B5) & "</B5>" & "<N1>" & lvResend.ListItems(0).Text & "</N1> [PASSON]"
'                Case 13
'                     bSteps = bSteps + 1
'                     B6 = IRC.colKeys(IRC.colKeys.FindKey(lvResend.ListItems(0).Text)).B6
'                     send "PRIVMSG " + Trim(lvResend.ListItems(0).SubItems(1)) + " :[IRCKey]<B6>" & MySQL.NumCrypt(B6) & "</B6>" & "<N1>" & lvResend.ListItems(0).Text & "</N1> [PASSON]"
'                Case 14
'                     bSteps = bSteps + 1
'                     B7 = IRC.colKeys(IRC.colKeys.FindKey(lvResend.ListItems(0).Text)).B7
'                     send "PRIVMSG " + Trim(lvResend.ListItems(0).SubItems(1)) + " :[IRCKey]<B7>" & MySQL.NumCrypt(B7) & "</B7>" & "<N1>" & lvResend.ListItems(0).Text & "</N1> [PASSON]"
'                 Case 15
'                     bSteps = bSteps + 1
'                     B8 = IRC.colKeys(IRC.colKeys.FindKey(lvResend.ListItems(0).Text)).B8
'                     send "PRIVMSG " + Trim(lvResend.ListItems(0).SubItems(1)) + " :[IRCKey]<B8>" & MySQL.NumCrypt(B8) & "</B8>" & "<N1>" & lvResend.ListItems(0).Text & "</N1> [PASSON]"
'                  Case 16
'                    send "PRIVMSG " + Trim(lstKeySend.List(0)) + " :[IRCKey]<D3>" & lvKeys.ListItems(IRC.colKeys(IRC.colKeys.FindKey(lvResend.ListItems(0).Text)).Key).SubItems(1) & "</D3>"
'                    bSteps = bSteps + 1
'                  Case 17
'                    send "PRIVMSG " + Trim(lstKeySend.List(0)) + " :[IRCKey]<D4>" & Format(sysNOW, "yyyy-mm-dd ttttt") & "</D4>"
'                    cDebug "IRC:// Completed sending IRC Key to " & Trim(lstKeySend.List(0)) & " on system stamp " & Format(sysNOW, "yyyy-mm-dd ttttt")
'                    bSteps = bSteps + 1
'                  Case 18
'                    bSteps = 0
'                    lvResend.ListItems.Remove 1
'                End Select
'            Else
'                lvResend.ListItems.Remove 1
'                lCount = lCount + 18
'            End If
'
'            lblIRCStat.Caption = "Status: Retransmitting new keys - " & Round((lCount / lSteps) * 100) & "% done"
'
'            If Not pbar.Max = 100 Then pbar.Max = 100
'            If pbar.Max < Round((lCount / lSteps) * 100) Then pbar.Max = Round((lCount / lSteps) * 102)
'            pbar.Value = Round((lCount / lSteps) * 100)
'
'        Case SendKeys
'           If lstKeySend.List(0) <> nick$ Then
'                Select Case bSteps
'                Case 0
'                     bSteps = bSteps + 1
'                     A1 = GetSetting(App.ProductName, "IRCKey", "A-1", 187)
'                     A1 = A1 + "," + GetSetting(App.ProductName, "IRCKey", "A-2", 187)
'                     A1 = A1 + "," + GetSetting(App.ProductName, "IRCKey", "A-3", 187)
'                     A1 = A1 + "," + GetSetting(App.ProductName, "IRCKey", "A-4", 187)
'                     A1 = A1 + "," + GetSetting(App.ProductName, "IRCKey", "A-5", 187)
'                     A1 = A1 + "," + GetSetting(App.ProductName, "IRCKey", "A-6", 187)
'                     A1 = A1 + "," + GetSetting(App.ProductName, "IRCKey", "A-7", 187)
'                     A1 = A1 + "," + GetSetting(App.ProductName, "IRCKey", "A-8", 187)
'                     A1 = A1 + "," + GetSetting(App.ProductName, "IRCKey", "A-9", 187)
'                     A1 = A1 + "," + GetSetting(App.ProductName, "IRCKey", "A-10", 187)
'                     A1 = A1 + "," + GetSetting(App.ProductName, "IRCKey", "A-11", 187)
'                     send "PRIVMSG " + Trim(lstKeySend.List(0)) + " :[IRCKey]<A1>" & MySQL.NumCrypt(A1) & "</A1>"
'                Case 1
'                     bSteps = bSteps + 1
'                     A2 = GetSetting(App.ProductName, "IRCKey", "A-12", 187)
'                     A2 = A2 + "," + GetSetting(App.ProductName, "IRCKey", "A-13", 187)
'                     A2 = A2 + "," + GetSetting(App.ProductName, "IRCKey", "A-14", 187)
'                     A2 = A2 + "," + GetSetting(App.ProductName, "IRCKey", "A-15", 187)
'                     A2 = A2 + "," + GetSetting(App.ProductName, "IRCKey", "A-16", 187)
'                     A2 = A2 + "," + GetSetting(App.ProductName, "IRCKey", "A-17", 187)
'                     A2 = A2 + "," + GetSetting(App.ProductName, "IRCKey", "A-18", 187)
'                     A2 = A2 + "," + GetSetting(App.ProductName, "IRCKey", "A-19", 187)
'                     A2 = A2 + "," + GetSetting(App.ProductName, "IRCKey", "A-20", 187)
'                     A2 = A2 + "," + GetSetting(App.ProductName, "IRCKey", "A-21", 187)
'                     A2 = A2 + "," + GetSetting(App.ProductName, "IRCKey", "A-22", 187)
'                     send "PRIVMSG " + Trim(lstKeySend.List(0)) + " :[IRCKey]<A2>" & MySQL.NumCrypt(A2) & "</A2>"
'                Case 2
'                     bSteps = bSteps + 1
'                     A3 = A3 + "" + GetSetting(App.ProductName, "IRCKey", "A-23", 187)
'                     A3 = A3 + "," + GetSetting(App.ProductName, "IRCKey", "A-24", 187)
'                     A3 = A3 + "," + GetSetting(App.ProductName, "IRCKey", "A-25", 187)
'                     A3 = A3 + "," + GetSetting(App.ProductName, "IRCKey", "A-26", 187)
'                     A3 = A3 + "," + GetSetting(App.ProductName, "IRCKey", "A-27", 187)
'                     A3 = A3 + "," + GetSetting(App.ProductName, "IRCKey", "A-28", 187)
'                     A3 = A3 + "," + GetSetting(App.ProductName, "IRCKey", "A-29", 187)
'                     A3 = A3 + "," + GetSetting(App.ProductName, "IRCKey", "A-30", 187)
'                     A3 = A3 + "," + GetSetting(App.ProductName, "IRCKey", "A-31", 187)
'                     A3 = A3 + "," + GetSetting(App.ProductName, "IRCKey", "A-32", 187)
'                     A3 = A3 + "," + GetSetting(App.ProductName, "IRCKey", "A-33", 187)
'                     send "PRIVMSG " + Trim(lstKeySend.List(0)) + " :[IRCKey]<A3>" & MySQL.NumCrypt(A3) & "</A3>"
'                Case 3
'                     bSteps = bSteps + 1
'                     A4 = A4 + "" + GetSetting(App.ProductName, "IRCKey", "A-34", 187)
'                     A4 = A4 + "," + GetSetting(App.ProductName, "IRCKey", "A-35", 187)
'                     A4 = A4 + "," + GetSetting(App.ProductName, "IRCKey", "A-36", 187)
'                     A4 = A4 + "," + GetSetting(App.ProductName, "IRCKey", "A-37", 187)
'                     A4 = A4 + "," + GetSetting(App.ProductName, "IRCKey", "A-38", 187)
'                     A4 = A4 + "," + GetSetting(App.ProductName, "IRCKey", "A-39", 187)
'                     A4 = A4 + "," + GetSetting(App.ProductName, "IRCKey", "A-40", 187)
'                     A4 = A4 + "," + GetSetting(App.ProductName, "IRCKey", "A-41", 187)
'                     A4 = A4 + "," + GetSetting(App.ProductName, "IRCKey", "A-42", 187)
'                     A4 = A4 + "," + GetSetting(App.ProductName, "IRCKey", "A-43", 187)
'                     A4 = A4 + "," + GetSetting(App.ProductName, "IRCKey", "A-44", 187)
'                     send "PRIVMSG " + Trim(lstKeySend.List(0)) + " :[IRCKey]<A4>" & MySQL.NumCrypt(A4) & "</A4>"
'                Case 4
'                     bSteps = bSteps + 1
'                     A5 = A5 + "" + GetSetting(App.ProductName, "IRCKey", "A-45", 187)
'                     A5 = A5 + "," + GetSetting(App.ProductName, "IRCKey", "A-46", 187)
'                     A5 = A5 + "," + GetSetting(App.ProductName, "IRCKey", "A-47", 187)
'                     A5 = A5 + "," + GetSetting(App.ProductName, "IRCKey", "A-48", 187)
'                     A5 = A5 + "," + GetSetting(App.ProductName, "IRCKey", "A-49", 187)
'                     A5 = A5 + "," + GetSetting(App.ProductName, "IRCKey", "A-50", 187)
'                     A5 = A5 + "," + GetSetting(App.ProductName, "IRCKey", "A-51", 187)
'                     A5 = A5 + "," + GetSetting(App.ProductName, "IRCKey", "A-52", 187)
'                     A5 = A5 + "," + GetSetting(App.ProductName, "IRCKey", "A-53", 187)
'                     A5 = A5 + "," + GetSetting(App.ProductName, "IRCKey", "A-54", 187)
'                     A5 = A5 + "," + GetSetting(App.ProductName, "IRCKey", "A-55", 187)
'                     send "PRIVMSG " + Trim(lstKeySend.List(0)) + " :[IRCKey]<A5>" & MySQL.NumCrypt(A5) & "</A5>"
'                Case 5
'                     bSteps = bSteps + 1
'                     A6 = A6 + "" + GetSetting(App.ProductName, "IRCKey", "A-56", 187)
'                     A6 = A6 + "," + GetSetting(App.ProductName, "IRCKey", "A-57", 187)
'                     A6 = A6 + "," + GetSetting(App.ProductName, "IRCKey", "A-58", 187)
'                     A6 = A6 + "," + GetSetting(App.ProductName, "IRCKey", "A-59", 187)
'                     A6 = A6 + "," + GetSetting(App.ProductName, "IRCKey", "A-60", 187)
'                     A6 = A6 + "," + GetSetting(App.ProductName, "IRCKey", "A-61", 187)
'                     A6 = A6 + "," + GetSetting(App.ProductName, "IRCKey", "A-62", 187)
'                     A6 = A6 + "," + GetSetting(App.ProductName, "IRCKey", "A-63", 187)
'                     A6 = A6 + "," + GetSetting(App.ProductName, "IRCKey", "A-64", 187)
'                     A6 = A6 + "," + GetSetting(App.ProductName, "IRCKey", "A-65", 187)
'                     A6 = A6 + "," + GetSetting(App.ProductName, "IRCKey", "A-66", 187)
'                     send "PRIVMSG " + Trim(lstKeySend.List(0)) + " :[IRCKey]<A6>" & MySQL.NumCrypt(A6) & "</A6>"
'                Case 6
'                     bSteps = bSteps + 1
'                     A7 = A7 + "" + GetSetting(App.ProductName, "IRCKey", "A-67", 187)
'                     A7 = A7 + "," + GetSetting(App.ProductName, "IRCKey", "A-68", 187)
'                     A7 = A7 + "," + GetSetting(App.ProductName, "IRCKey", "A-69", 187)
'                     A7 = A7 + "," + GetSetting(App.ProductName, "IRCKey", "A-70", 187)
'                     A7 = A7 + "," + GetSetting(App.ProductName, "IRCKey", "A-71", 187)
'                     A7 = A7 + "," + GetSetting(App.ProductName, "IRCKey", "A-72", 187)
'                     A7 = A7 + "," + GetSetting(App.ProductName, "IRCKey", "A-73", 187)
'                     A7 = A7 + "," + GetSetting(App.ProductName, "IRCKey", "A-74", 187)
'                     A7 = A7 + "," + GetSetting(App.ProductName, "IRCKey", "A-75", 187)
'                     A7 = A7 + "," + GetSetting(App.ProductName, "IRCKey", "A-76", 187)
'                     A7 = A7 + "," + GetSetting(App.ProductName, "IRCKey", "A-77", 187)
'                     send "PRIVMSG " + Trim(lstKeySend.List(0)) + " :[IRCKey]<A7>" & MySQL.NumCrypt(A7) & "</A7>"
'                 Case 7
'                     bSteps = bSteps + 1
'                     A8 = A8 + "" + GetSetting(App.ProductName, "IRCKey", "A-78", 187)
'                     A8 = A8 + "," + GetSetting(App.ProductName, "IRCKey", "A-79", 187)
'                     A8 = A8 + "," + GetSetting(App.ProductName, "IRCKey", "A-80", 187)
'                     A8 = A8 + "," + GetSetting(App.ProductName, "IRCKey", "A-81", 187)
'                     A8 = A8 + "," + GetSetting(App.ProductName, "IRCKey", "A-82", 187)
'                     A8 = A8 + "," + GetSetting(App.ProductName, "IRCKey", "A-83", 187)
'                     A8 = A8 + "," + GetSetting(App.ProductName, "IRCKey", "A-84", 187)
'                     A8 = A8 + "," + GetSetting(App.ProductName, "IRCKey", "A-85", 187)
'                     A8 = A8 + "," + GetSetting(App.ProductName, "IRCKey", "A-86", 187)
'                     A8 = A8 + "," + GetSetting(App.ProductName, "IRCKey", "A-87", 187)
'                     send "PRIVMSG " + Trim(lstKeySend.List(0)) + " :[IRCKey]<A8>" & MySQL.NumCrypt(A8) & "</A8>"
'                Case 8
'                     bSteps = bSteps + 1
'                     B1 = GetSetting(App.ProductName, "IRCKey", "B-1", 187)
'                     B1 = B1 + "," + GetSetting(App.ProductName, "IRCKey", "B-2", 187)
'                     B1 = B1 + "," + GetSetting(App.ProductName, "IRCKey", "B-3", 187)
'                     B1 = B1 + "," + GetSetting(App.ProductName, "IRCKey", "B-4", 187)
'                     B1 = B1 + "," + GetSetting(App.ProductName, "IRCKey", "B-5", 187)
'                     B1 = B1 + "," + GetSetting(App.ProductName, "IRCKey", "B-6", 187)
'                     B1 = B1 + "," + GetSetting(App.ProductName, "IRCKey", "B-7", 187)
'                     B1 = B1 + "," + GetSetting(App.ProductName, "IRCKey", "B-8", 187)
'                     B1 = B1 + "," + GetSetting(App.ProductName, "IRCKey", "B-9", 187)
'                     B1 = B1 + "," + GetSetting(App.ProductName, "IRCKey", "B-10", 187)
'                     B1 = B1 + "," + GetSetting(App.ProductName, "IRCKey", "B-11", 187)
'                     send "PRIVMSG " + Trim(lstKeySend.List(0)) + " :[IRCKey]<B1>" & MySQL.NumCrypt(B1) & "</B1>"
'                Case 9
'                     bSteps = bSteps + 1
'                     B2 = GetSetting(App.ProductName, "IRCKey", "B-12", 187)
'                     B2 = B2 + "," + GetSetting(App.ProductName, "IRCKey", "B-13", 187)
'                     B2 = B2 + "," + GetSetting(App.ProductName, "IRCKey", "B-14", 187)
'                     B2 = B2 + "," + GetSetting(App.ProductName, "IRCKey", "B-15", 187)
'                     B2 = B2 + "," + GetSetting(App.ProductName, "IRCKey", "B-16", 187)
'                     B2 = B2 + "," + GetSetting(App.ProductName, "IRCKey", "B-17", 187)
'                     B2 = B2 + "," + GetSetting(App.ProductName, "IRCKey", "B-18", 187)
'                     B2 = B2 + "," + GetSetting(App.ProductName, "IRCKey", "B-19", 187)
'                     B2 = B2 + "," + GetSetting(App.ProductName, "IRCKey", "B-20", 187)
'                     B2 = B2 + "," + GetSetting(App.ProductName, "IRCKey", "B-21", 187)
'                     B2 = B2 + "," + GetSetting(App.ProductName, "IRCKey", "B-22", 187)
'                     send "PRIVMSG " + Trim(lstKeySend.List(0)) + " :[IRCKey]<B2>" & MySQL.NumCrypt(B2) & "</B2>"
'                Case 10
'                     bSteps = bSteps + 1
'                     B3 = B3 + "" + GetSetting(App.ProductName, "IRCKey", "B-23", 187)
'                     B3 = B3 + "," + GetSetting(App.ProductName, "IRCKey", "B-24", 187)
'                     B3 = B3 + "," + GetSetting(App.ProductName, "IRCKey", "B-25", 187)
'                     B3 = B3 + "," + GetSetting(App.ProductName, "IRCKey", "B-26", 187)
'                     B3 = B3 + "," + GetSetting(App.ProductName, "IRCKey", "B-27", 187)
'                     B3 = B3 + "," + GetSetting(App.ProductName, "IRCKey", "B-28", 187)
'                     B3 = B3 + "," + GetSetting(App.ProductName, "IRCKey", "B-29", 187)
'                     B3 = B3 + "," + GetSetting(App.ProductName, "IRCKey", "B-30", 187)
'                     B3 = B3 + "," + GetSetting(App.ProductName, "IRCKey", "B-31", 187)
'                     B3 = B3 + "," + GetSetting(App.ProductName, "IRCKey", "B-32", 187)
'                     B3 = B3 + "," + GetSetting(App.ProductName, "IRCKey", "B-33", 187)
'                     send "PRIVMSG " + Trim(lstKeySend.List(0)) + " :[IRCKey]<B3>" & MySQL.NumCrypt(B3) & "</B3>"
'                Case 11
'                     bSteps = bSteps + 1
'                     B4 = B4 + "" + GetSetting(App.ProductName, "IRCKey", "B-34", 187)
'                     B4 = B4 + "," + GetSetting(App.ProductName, "IRCKey", "B-35", 187)
'                     B4 = B4 + "," + GetSetting(App.ProductName, "IRCKey", "B-36", 187)
'                     B4 = B4 + "," + GetSetting(App.ProductName, "IRCKey", "B-37", 187)
'                     B4 = B4 + "," + GetSetting(App.ProductName, "IRCKey", "B-38", 187)
'                     B4 = B4 + "," + GetSetting(App.ProductName, "IRCKey", "B-39", 187)
'                     B4 = B4 + "," + GetSetting(App.ProductName, "IRCKey", "B-40", 187)
'                     B4 = B4 + "," + GetSetting(App.ProductName, "IRCKey", "B-41", 187)
'                     B4 = B4 + "," + GetSetting(App.ProductName, "IRCKey", "B-42", 187)
'                     B4 = B4 + "," + GetSetting(App.ProductName, "IRCKey", "B-43", 187)
'                     B4 = B4 + "," + GetSetting(App.ProductName, "IRCKey", "B-44", 187)
'                     send "PRIVMSG " + Trim(lstKeySend.List(0)) + " :[IRCKey]<B4>" & MySQL.NumCrypt(B4) & "</B4>"
'                Case 12
'                     bSteps = bSteps + 1
'                     B5 = B5 + "" + GetSetting(App.ProductName, "IRCKey", "B-45", 187)
'                     B5 = B5 + "," + GetSetting(App.ProductName, "IRCKey", "B-46", 187)
'                     B5 = B5 + "," + GetSetting(App.ProductName, "IRCKey", "B-47", 187)
'                     B5 = B5 + "," + GetSetting(App.ProductName, "IRCKey", "B-48", 187)
'                     B5 = B5 + "," + GetSetting(App.ProductName, "IRCKey", "B-49", 187)
'                     B5 = B5 + "," + GetSetting(App.ProductName, "IRCKey", "B-50", 187)
'                     B5 = B5 + "," + GetSetting(App.ProductName, "IRCKey", "B-51", 187)
'                     B5 = B5 + "," + GetSetting(App.ProductName, "IRCKey", "B-52", 187)
'                     B5 = B5 + "," + GetSetting(App.ProductName, "IRCKey", "B-53", 187)
'                     B5 = B5 + "," + GetSetting(App.ProductName, "IRCKey", "B-54", 187)
'                     B5 = B5 + "," + GetSetting(App.ProductName, "IRCKey", "B-55", 187)
'                     send "PRIVMSG " + Trim(lstKeySend.List(0)) + " :[IRCKey]<B5>" & MySQL.NumCrypt(B5) & "</B5>"
'                Case 13
'                     bSteps = bSteps + 1
'                     B6 = B6 + "" + GetSetting(App.ProductName, "IRCKey", "B-56", 187)
'                     B6 = B6 + "," + GetSetting(App.ProductName, "IRCKey", "B-57", 187)
'                     B6 = B6 + "," + GetSetting(App.ProductName, "IRCKey", "B-58", 187)
'                     B6 = B6 + "," + GetSetting(App.ProductName, "IRCKey", "B-59", 187)
'                     B6 = B6 + "," + GetSetting(App.ProductName, "IRCKey", "B-60", 187)
'                     B6 = B6 + "," + GetSetting(App.ProductName, "IRCKey", "B-61", 187)
'                     B6 = B6 + "," + GetSetting(App.ProductName, "IRCKey", "B-62", 187)
'                     B6 = B6 + "," + GetSetting(App.ProductName, "IRCKey", "B-63", 187)
'                     B6 = B6 + "," + GetSetting(App.ProductName, "IRCKey", "B-64", 187)
'                     B6 = B6 + "," + GetSetting(App.ProductName, "IRCKey", "B-65", 187)
'                     B6 = B6 + "," + GetSetting(App.ProductName, "IRCKey", "B-66", 187)
'                     send "PRIVMSG " + Trim(lstKeySend.List(0)) + " :[IRCKey]<B6>" & MySQL.NumCrypt(B6) & "</B6>"
'                Case 14
'                     bSteps = bSteps + 1
'                     B7 = B7 + "" + GetSetting(App.ProductName, "IRCKey", "B-67", 187)
'                     B7 = B7 + "," + GetSetting(App.ProductName, "IRCKey", "B-68", 187)
'                     B7 = B7 + "," + GetSetting(App.ProductName, "IRCKey", "B-69", 187)
'                     B7 = B7 + "," + GetSetting(App.ProductName, "IRCKey", "B-70", 187)
'                     B7 = B7 + "," + GetSetting(App.ProductName, "IRCKey", "B-71", 187)
'                     B7 = B7 + "," + GetSetting(App.ProductName, "IRCKey", "B-72", 187)
'                     B7 = B7 + "," + GetSetting(App.ProductName, "IRCKey", "B-73", 187)
'                     B7 = B7 + "," + GetSetting(App.ProductName, "IRCKey", "B-74", 187)
'                     B7 = B7 + "," + GetSetting(App.ProductName, "IRCKey", "B-75", 187)
'                     B7 = B7 + "," + GetSetting(App.ProductName, "IRCKey", "B-76", 187)
'                     B7 = B7 + "," + GetSetting(App.ProductName, "IRCKey", "B-77", 187)
'                     send "PRIVMSG " + Trim(lstKeySend.List(0)) + " :[IRCKey]<B7>" & MySQL.NumCrypt(B7) & "</B7>"
'                 Case 15
'                     bSteps = bSteps + 1
'                     B8 = B8 + "" + GetSetting(App.ProductName, "IRCKey", "B-78", 187)
'                     B8 = B8 + "," + GetSetting(App.ProductName, "IRCKey", "B-79", 187)
'                     B8 = B8 + "," + GetSetting(App.ProductName, "IRCKey", "B-80", 187)
'                     B8 = B8 + "," + GetSetting(App.ProductName, "IRCKey", "B-81", 187)
'                     B8 = B8 + "," + GetSetting(App.ProductName, "IRCKey", "B-82", 187)
'                     B8 = B8 + "," + GetSetting(App.ProductName, "IRCKey", "B-83", 187)
'                     B8 = B8 + "," + GetSetting(App.ProductName, "IRCKey", "B-84", 187)
'                     B8 = B8 + "," + GetSetting(App.ProductName, "IRCKey", "B-85", 187)
'                     B8 = B8 + "," + GetSetting(App.ProductName, "IRCKey", "B-86", 187)
'                     B8 = B8 + "," + GetSetting(App.ProductName, "IRCKey", "B-87", 187)
'                     send "PRIVMSG " + Trim(lstKeySend.List(0)) + " :[IRCKey]<B8>" & MySQL.NumCrypt(B8) & "</B8>"
'                  Case 16
'                    send "PRIVMSG " + Trim(lstKeySend.List(0)) + " :[IRCKey]<D3>" & Format(GetSetting(App.ProductName, "IRCKey", "Created", DateAdd("d", -1, sysNOW)), "yyyy-mm-dd ttttt") & "</D3>"
'                    bSteps = bSteps + 1
'                  Case 17
'                    send "PRIVMSG " + Trim(lstKeySend.List(0)) + " :[IRCKey]<D4>" & Format(sysNOW, "yyyy-mm-dd ttttt") & "</D4>"
'                    cDebug "IRC:// Completed sending IRC Key to " & Trim(lstKeySend.List(0)) & " on system stamp " & Format(sysNOW, "yyyy-mm-dd ttttt")
'                    bSteps = bSteps + 1
'                  Case 18
'
'
'                    If lstKeySend.ListCount - NumDistro - 3 > 1 Then
'                        Dim j As Long
'                        Dim XMLVal As String
'                        Dim iCnt As Byte
'
'                        XMLVal = "<NICK>" & nick$ & "</NICK>"
'                        XMLVal = XMLVal + "<KEYRQST>"
'                        iCnt = 1
'                        For j = lstKeySend.ListCount - NumDistro To lstKeySend.ListCount - 3 - NumDistro
'                            If j > 0 Then
'                                iCnt = iCnt + 1
'                                XMLVal = XMLVal + "<N" & iCnt & ">" & lstKeySend.List(j) & "</N" & iCnt & ">"
'                                NumDistro = NumDistro + 1
'                            End If
'                        Next j
'                        XMLVal = XMLVal + "</KEYRQST>"
'                        send "PRIVMSG " + Trim(lstKeySend.List(0)) + " :[IRCKey]<C1>" & XMLVal & "</C1>"
'                    End If
'
'                    bSteps = 0
'                    lstKeySend.RemoveItem 0
'
'                    If lstKeySend.ListCount = 0 Then
'                        lSteps = 0
'                        bMode = RequestKeys
'                        tmrPopulateKeys.Interval = 400
'                    End If
'
'                End Select
'            Else
'                lstKeySend.RemoveItem 0
'                lCount = lCount + 19
'
'                If lstKeySend.ListCount = 0 Then
'                    lSteps = 0
'                    bMode = RequestKeys
'                    tmrPopulateKeys.Interval = 400
'                End If
'
'            End If
'
'        End Select
'        If lSteps > 0 Then
'            lblIRCStat.Caption = "Status: Populating Encryption Matrix - " & Round((lCount / lSteps) * 100) & "% done"
'            If Not pbar.Max = 100 Then pbar.Max = 100
'            If pbar.Max < Round((lCount / lSteps) * 100) Then pbar.Max = Round((lCount / lSteps) * 102)
'            pbar.Value = Round((lCount / lSteps) * 100)
'        Else
'            lblIRCStat.Caption = "Status: Populating Encryption Matrix - " & lCount & " steps done done"
'            If pbar.Max < lCount Then pbar.Max = lCount + 20
'            pbar.Value = lCount
'        End If
'
'
'
'    End Select
'
'
'End Sub

Private Sub ts2_Click()

    'Dim bx As Byte
    
    'For bx = picTS2.L To picTS2.UBound
        
    '    Select Case ts2.SelectedItem.Index - 1
    '    Case bx
    '        picTS2(bx).ZOrder 0
    '        picTS2(bx).Move ts2.ClientLeft, ts2.ClientTop, ts2.ClientWidth, ts2.ClientHeight
    '        picTS2(bx).Visible = True
    '        picTS2(bx).Refresh
    '    Case Else
    '        picTS2(bx).Visible = False
        
    '    End Select
    
    'Next bx
    
End Sub

Private Sub tmrProgressCylinders_Timer()

    Dim hIcon As String
    
    Dim itmX As ListItem
    
    For Each itmX In lvSchedule.ListItems
        
        If Left(IIf(Len(itmX.SmallIcon) >= 4, itmX.SmallIcon, "      "), 4) = "prog" And itmX.SubItems(3) = "100%" Then
            If IsDate(itmX.Tag) = False Then itmX.Tag = sysnow
            If DateDiff("s", itmX.Tag, sysnow) > 50 Then
                itmX.SmallIcon = "red"
                itmX.SubItems(3) = "0%"
                itmX.SubItems(4) = "Idle"
            Else
                If Not itmX.SmallIcon = "progress8" Then itmX.SmallIcon = "progress8"
            End If
        Else
            If Val(MySQL.ReplaceString(itmX.SubItems(3), "%", "")) > 0 Or Val(MySQL.ReplaceString(itmX.SubItems(3), "%", "")) < 100 Then
                Select Case Val(MySQL.ReplaceString(itmX.SubItems(3), "%", ""))
                Case 0
                    If Not itmX.SmallIcon = "black" Then hIcon = "black"
                Case 1 To 12
                    If Not itmX.SmallIcon = "progress0" Then hIcon = "progress0"
                Case 12.5 To 25
                    If Not itmX.SmallIcon = "progress1" Then hIcon = "progress1"
                Case 37.5 To 49
                    If Not itmX.SmallIcon = "progress2" Then hIcon = "progress2"
                Case 50 To 61
                    If Not itmX.SmallIcon = "progress3" Then hIcon = "progress3"
                Case 62.5 To 74
                    If Not itmX.SmallIcon = "progress4" Then hIcon = "progress4"
                Case 75 To 86
                    If Not itmX.SmallIcon = "progress5" Then hIcon = "progress5"
                Case 87.5 To 94
                    If Not itmX.SmallIcon = "progress6" Then hIcon = "progress6"
                Case 95 To 99
                    If Not itmX.SmallIcon = "progress7" Then hIcon = "progress7"
                End Select
            ElseIf Val(MySQL.ReplaceString(itmX.SubItems(3), "%", "")) >= 100 Then
                
                If Not itmX.SmallIcon = "progress8" Then
                    hIcon = "progress8"
                    itmX.Tag = sysnow
                End If
            
            ElseIf Val(MySQL.ReplaceString(itmX.SubItems(3), "%", "")) = 0 And (itmX.SmallIcon <> "black" And itmX.SmallIcon <> "red") Then
                
                If Not itmX.SmallIcon = "black" Then
                    itmX.SmallIcon = "black"
                End If
                
            End If
        
            If Len(hIcon) > 0 Then
                If Not itmX.SmallIcon = hIcon Then itmX.SmallIcon = hIcon
                If Not itmX.Icon = hIcon Then itmX.Icon = hIcon
                hIcon = "black"
            End If
        End If
        
    Next itmX
    
    DoEvents
        
End Sub

Private Sub tsIRCChans_Click()

    Dim ix As Long
    
    For ix = picIRCChan.LBound To picIRCChan.UBound
        If tsIRCChans.SelectedItem.Index - 1 = ix Then
            picIRCChan(ix).Move tsIRCChans.ClientLeft, tsIRCChans.ClientTop, tsIRCChans.ClientWidth, tsIRCChans.ClientHeight
            picIRCChan(ix).Visible = True
            picIRCChan(ix).ZOrder 0
        Else
            picIRCChan(ix).Visible = False
        End If
    Next ix
    
End Sub

Private Sub tsSchedule_Click()

    Dim bx As Byte
    
    For bx = picOptions.LBound To picOptions.UBound
        Select Case tsSchedule.SelectedItem.Tag
        Case picOptions(bx).Tag
            picOptions(bx).BorderStyle = 0
            picOptions(bx).ZOrder 0
            picOptions(bx).Move tsSchedule.ClientLeft, tsSchedule.ClientTop, tsSchedule.ClientWidth, tsSchedule.ClientHeight
            picOptions(bx).Visible = True
            picOptions(bx).Refresh
        Case Else
            picOptions(bx).Visible = False
        End Select
    Next bx
    
End Sub
'
'Private Sub txtChannel_GotFocus()
'    cmdChannel.Default = True   'set the channel "change" button as the default button
'End Sub
'
'Private Sub txtChatMsg_GotFocus()
'    cmdChat.Default = True  'set the chat button as the default button
'End Sub
'
'Private Sub txtNick_GotFocus()
'    cmdNick.Default = True  'set the nick "change" button as the default button
'End Sub
'
'Private Sub txtPrivMsg_GotFocus()
'    cmdPriv.Default = True  'set the private button as the default button
'End Sub
'
'Private Sub txtStatus_DblClick()
'    c$ = InputBox("Please enter a command (eg. PRIVMSG Bot :Hello bot)" + vbCrLf + vbCrLf + "Command:", "Custom command")
'        'let the user enter a command
'    c$ = Trim(c$)   'clear any leading whitespace characters
'    If c$ = "" Then Exit Sub    'if the user canceled exit...
'    If UCase(Left(c$, 4)) = "JOIN" Then 'if the user wants to join a channel:
'        send "PART " + channel$ 'leave the current channel
'        lstUsers.Clear  'clear the user list
'        send "JOIN " + processParam(processRest(c$))    'only join first channel supplied by the user
'        channel$ = processParam(processRest(c$))    'store the channel
'        txtChannel.Text = channel$  'change the channel text box
'    ElseIf UCase(Left(c$, 4)) = "PART" Then     'if the user wants to leave the channel
'        lstUsers.Clear  'clear the user list
'        send "PART " + channel$ 'leave the channel
'        channel$ = ""   'clear the channel holder
'        txtChannel = channel$   'clear the text box
'    ElseIf UCase(Left(c$, 4)) = "NICK" Then     'if the user want to change his nickname
'        txtNick.Text = processParam(processRest(c$))    'store the first parameter in the nick text field
'        cmdNick_Click   'make it click the change nick button
'    ElseIf UCase(Left(c$, 4)) = "QUIT" Then 'if the user wants to quit
'        display "<!> QUIT message canceled! Please click the X button in the bottom right corner of the window!"
'            'dont do it, just display this message
'    Else    'if its an innocent command :)
'        send c$     'send it
'    End If
'End Sub
'
'Private Sub txtStatus_KeyPress(KeyAscii As Integer)
'    KeyAscii = 0    'ignore the keypress
'End Sub

'Sub display(Msg$)   'display a message in the status field:
'    'txtStatus.Text = txtStatus.Text + Msg$ + vbCrLf   ' add the message to the status field
'    'txtStatus.SelStart = Len(txtStatus.Text)  'select the end of the message
'    'txtStatus.SelLength = 0                'make sure nothing is displayed as "selected"
'
'    cDebug "IRC://" & Msg$
'
'End Sub
'
'Sub displaychat(Msg$)   'display a message in the chat field:
'    txtChat.Text = txtChat.Text + Msg$ + vbCrLf   ' add the message to the chat field
'    txtChat.SelStart = Len(txtChat.Text)  'select the end of the message
'    txtChat.SelLength = 0                'make sure nothing is displayed as "selected"
'End Sub
'
Public Sub send(Msg$)  'send a message to the IRC server
On Error GoTo oops  'if an error occures, goto the oops label
    'display ">> " + msg$    ' display the text in the main field
    sIRC.Write Msg$, Len(Msg$) 'send the data, along with a cariage return and a line feed
    Exit Sub    'skip the error handling section
oops:
    sockIRC.Close
End Sub
'
'Sub processCommand()
'
'
'
'    '*[ Error Checking Variables ]**********************************************************************************
'    Const RoutineName = "MDIForm_Load"
'    Const ContainerName = "processCommand"
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
'
'    Dim lx As Long
'    Dim CX As Long
'    Dim bPass As Boolean
'    Dim itmX As ListItem
'    Dim k As Long
'
'    ' the next line will reply to the PING message of the server
'    ' preventing us from going idle and being kicked
'    If InStr(data$, "PING") > 0 Then
'        Dim params$    ' parameters that will be filtered from the pong message
'        params$ = Right$(data$, Len(data$) - (InStr(data$, "PING") + 4))
'            'take the paramaters from the right of the message starting from the first character after the PING message
'        send "PONG " + params$   ' send the pong message to the server, together with the parameters
'        display "PING? PONG!"
'    End If
'
'    'This section processes all other commands
'    If Left$(data$, 1) = ":" Then   'if the message starts with a colon (standard IRC message)
'        Dim pos%, pos2%    '2 position variables we need to extract the nickname of whoever that issued the command
'        Dim from$, rest$    'these will hold the sender of the command and the rest of the message
'        Dim command$        'this will hold the type of the command (eg.: PRIVMSG)
'        params$ = ""        'and the parameters
'        pos% = InStr(data$, " ")    'get the position of the first space character
'        If pos% > 0 Then    'if a space is found
'            pos2% = InStr(data$, "!")   'search for an exclamation mark
'            If pos% < pos2% Or pos2% <= 0 Then pos2% = pos%   'if a space is found AFTER the space, it should not be used
'            from$ = Mid$(data$, 2, pos2% - 2)   'parse the sender, starting from the second character (after the ":")
'            rest$ = Mid$(data$, pos% + 1, Len(data$) - pos2%)  'parse the rest of the message starting from the first character AFTER the first space
'            'IMPORTANT: pos% is now used to hold the first space in (!) rest$ (!), *NOT* in data$
'            pos% = InStr(rest$, " ")   'get the position of the first space in rest$
'            If pos% > 0 Then    'if we found a space
'                command$ = Left$(rest$, pos% - 1)   'the part before this space is the type of command
'                params$ = Right$(rest$, Len(rest$) - pos%)   'the rest are parameters
'                Select Case command$    'base your actions on the type of command
'                    Case "NOTICE"   'if it's a notice
'                        displaychat "( " + from$ + " notices: " + params$ + " )" 'display it
'                    Case "PRIVMSG"  'if it's a private message
'                            Dim itmy As ListItem
'
'                            If processParam(params$) = channel$ Then
'                                If IRC.colKeys.FindKey(from$) = -1 Then
'                                    displaychat "[ " + from$ + " ] *** Encryption Keys not in memory Requesting from client ***"
'                                    Set itmy = lvBuf.ListItems.Add(, "ID" & lvBuf.ListItems.Count + 1, "ciphered", , "closed")
'                                    itmy.SubItems(1) = from$
'                                    itmy.SubItems(2) = ""
'                                    itmy.SubItems(3) = Trim(processParam(processRest(params$)))
'                                    itmy.SubItems(4) = Format(sysnow, "yyyy-mm-dd ttttt")
'                                    itmy.SubItems(5) = Format(Now, "yyyy-mm-dd ttttt")
'                                    send "PRIVMSG " + from$ + " :[IRCKey] Request Key " & Format(Now, "yyyy-mm-dd ttttt")
'                                Else
'                                    Set itmy = lvBuf.ListItems.Add(, "ID" & lvBuf.ListItems.Count + 1, "decrypt", , "open")
'                                    itmy.SubItems(1) = from$
'                                    itmy.SubItems(2) = IRC.colKeys.DecryptIRC(Trim(processParam(processRest(params$))), from$)
'                                    itmy.SubItems(3) = Trim(processParam(processRest(params$)))
'                                    itmy.SubItems(4) = Format(sysnow, "yyyy-mm-dd ttttt")
'                                    itmy.SubItems(5) = Format(Now, "yyyy-mm-dd ttttt")
'                                    displaychat "[ " + from$ + " ]  " & IRC.colKeys.DecryptIRC(Trim(processParam(processRest(params$))), from$)
'                                End If
'                                    'if you want autoreplies, autoevents, ... , just add them here
'                            ElseIf processParam(params$) = nick$ Then
'                                If Left(Trim(processParam(processRest(params$))), 8) = "[DATEOK]" Or Left(Trim(processParam(processRest(params$))), 8) = "[KEYREV]" Then
'                                    cDebug "IRC:// [DATEOK] recieved from " & from$
'                                    For lx = lstKeySend.ListCount - 1 To 0 Step -1
'                                        If lstKeySend.List(lx) = from$ Then
'                                            lstKeySend.RemoveItem lx
'                                            Exit For
'                                        End If
'                                    Next
'                                    For lx = lstKeyRqst.ListCount - 1 To 0 Step -1
'                                        If lstKeyRqst.List(lx) = from$ Then
'                                            lstKeyRqst.RemoveItem lx
'                                            Exit For
'                                        End If
'                                    Next
'                                ElseIf Left(Trim(processParam(processRest(params$))), 10) = "[CRCCHECK]" Then
'                                    send "PRIVMSG " + from$ + " :[DATECRC]<D1>" & Format(GetSetting(App.ProductName, "IRCKey", "Created", DateAdd("d", -1, sysnow)), "yyyy-mm-dd ttttt") & "</D1><D1>" & Format(GetSetting(App.ProductName, "IRCKey", "Created", DateAdd("d", -1, sysnow)), "yyyy-mm-dd ttttt") & "</D1><D1>" & Format(GetSetting(App.ProductName, "IRCKey", "Created", DateAdd("d", -1, sysnow)), "yyyy-mm-dd ttttt") & "</D1><D2>" & Format(sysnow, "yyyy-mm-dd ttttt") & "</D2>"
'                                    cDebug "IRC:// [CRCCHECK] request for Date CRC sent from " & from$ & " [SENT]"
'                                ElseIf Left(Trim(processParam(processRest(params$))), 9) = "[DATECRC]" Then
'                                    bPass = False
'                                    CX = -1
'
'                                    For lx = 1 To lvKeys.ListItems.Count
'                                        If lvKeys.ListItems(lx).Text = from$ Then
'                                            If XMLTag(1, Trim(processParam(processRest(params$))), "D1") = lvKeys.ListItems(lx).SubItems(1) Then
'                                                lvKeys.ListItems(lx).SubItems(2) = XMLTag(1, Trim(processParam(processRest(params$))), "D2")
'                                                bPass = True
'                                                send "PRIVMSG " + from$ + " :[DATEOK]"
'                                                cDebug "IRC:// [DATECRC] checked out aok sent [DATEOK] to " & from$
'                                            End If
'                                            CX = lx
'                                            Exit For
'                                        End If
'                                    Next
'
'                                    If bPass = False Then
'                                        If IRC.colKeys.FindKey(from$) <> -1 Then
'                                            IRC.colKeys.Remove IRC.colKeys.FindKey(from$)
'                                            send "PRIVMSG " + from$ + " :[IRCKey] Request Key " & Format(Now, "yyyy-mm-dd ttttt")
'                                        End If
'                                        If CX <> -1 Then
'                                            IRC.colKeys.Remove IRC.colKeys.FindKey(lvKeys.ListItems(CX).Text)
'                                            lvKeys.ListItems.Remove CX
'                                            CX = -1
'                                        End If
'                                    End If
'
'                                ElseIf Left(Trim(processParam(processRest(params$))), 8) = "[IRCKey]" Then
'                                    cDebug "IRC:// [IRCKey] recieved from " & from$
'                                    Dim cKeys As clsKeys
'                                    Dim itx As ListItem
'                                    If InStr(processParam(processRest(params$)), "Request") <> 0 Then
'                                        lstKeyRqst.AddItem from$
'                                    ElseIf Right(processParam(processRest(params$)), 8) = "[PASSON]" Then
'
'                                        Dim orgFrom As String
'
'                                        orgFrom = XMLTag(1, Trim(processParam(processRest(params$))), "N1")
'                                        cDebug "IRC:// [IRCKey] [PASSON] from {" & from$ & "} recieved in reply to request for key from {" & orgFrom & "}"
'                                        If orgFrom <> "" Then
'
'                                            If IRC.colKeys.FindKey(orgFrom) = -1 Then
'                                                Set cKeys = IRC.colKeys.Add("ID" & IRC.colKeys.Count, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", orgFrom, "ID" & IRC.colKeys.Count)
'                                                Set itx = lvKeys.ListItems.Add(, cKeys.Key, orgFrom)
'
'                                            Else
'                                                Set cKeys = IRC.colKeys(IRC.colKeys.FindKey(orgFrom))
'                                                Set itx = lvKeys.ListItems(cKeys.Key)
'                                            End If
'
'                                            Select Case Mid(Trim(processParam(processRest(params$))), 10, 2)
'                                        Case "C1"
'                                            itx.Tag = XMLTag(1, Trim(processParam(processRest(params$))), "C1")
'
'                                            If XMLTag(1, itx.Tag, "N2") <> "" Then
'                                                Set itmX = lvResend.ListItems.Add(, cKeys.Key, from$)
'                                                itmX.SubItems(1) = XMLTag(1, itx.Tag, "N2")
'                                                itmX.SubItems(2) = itx.SubItems(1)
'                                                cDebug "IRC:// [PASSON] request from {" & from$ & "} recieved, queing for resend, a little latter."
'                                            End If
'
'                                            If XMLTag(1, itx.Tag, "N3") <> "" Then
'                                                Set itmX = lvResend.ListItems.Add(, cKeys.Key, from$)
'                                                itmX.SubItems(1) = XMLTag(1, itx.Tag, "N2")
'                                                itmX.SubItems(2) = itx.SubItems(1)
'                                                cDebug "IRC:// [PASSON] request from {" & from$ & "} recieved, queing for resend, a little latter."
'                                            End If
'
'                                            If XMLTag(1, itx.Tag, "N4") <> "" Then
'                                                Set itmX = lvResend.ListItems.Add(, cKeys.Key, from$)
'                                                itmX.SubItems(1) = XMLTag(1, itx.Tag, "N2")
'                                                itmX.SubItems(2) = itx.SubItems(1)
'                                                cDebug "IRC:// [PASSON] request from {" & from$ & "} recieved, queing for resending to " & itx.SubItems(1)
'                                            End If
'
'                                        Case "D3"
'                                            itx.SubItems(1) = XMLTag(1, Trim(processParam(processRest(params$))), "D3")
'                                            cDebug "IRC:// [DATELH] recieved from {" & from$ & "} date is '" & itx.SubItems(1) & "'"
'                                        Case "D4"
'                                            itx.SubItems(2) = XMLTag(1, Trim(processParam(processRest(params$))), "D4")
'                                            send "PRIVMSG " + from$ + " :[KEYREV]<D1>" & Format(GetSetting(App.ProductName, "IRCKey", "Created", DateAdd("d", -1, sysnow)), "yyyy-mm-dd ttttt") & "</D1>" + _
'                                                                      "<D2>" & Format(sysnow, "yyyy-mm-dd ttttt") & "</D2>"
'                                            cDebug "IRC:// [DATERH] recieved from {" & from$ & "} date is '" & itx.SubItems(1) & "'"
'                                        Case "A1"
'                                            cKeys.A1 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "A1"))
'                                            cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.A1) & "'"
'                                        Case "A2"
'                                            cKeys.A2 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "A2"))
'                                            cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.A2) & "'"
'                                        Case "A3"
'                                            cKeys.A3 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "A3"))
'                                            cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.A3) & "'"
'                                        Case "A4"
'                                            cKeys.A4 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "A4"))
'                                            cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.A4) & "'"
'                                        Case "A5"
'                                            cKeys.A5 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "A5"))
'                                            cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.A5) & "'"
'                                        Case "A6"
'                                            cKeys.A6 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "A6"))
'                                            cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.A6) & "'"
'                                        Case "A7"
'                                            cKeys.A7 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "A7"))
'                                            cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.A7) & "'"
'                                        Case "A8"
'                                            cKeys.A8 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "A8"))
'                                            cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.A8) & "'"
'                                        Case "B1"
'                                            cKeys.B1 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "B1"))
'                                            cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.B1) & "'"
'                                        Case "B2"
'                                            cKeys.B2 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "B2"))
'                                            cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.B2) & "'"
'                                        Case "B3"
'                                            cKeys.B3 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "B3"))
'                                            cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.B3) & "'"
'                                        Case "B4"
'                                            cKeys.B4 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "B4"))
'                                            cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.B4) & "'"
'                                        Case "B5"
'                                            cKeys.B5 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "B5"))
'                                            cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.B5) & "'"
'                                        Case "B6"
'                                            cKeys.B6 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "B6"))
'                                            cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.B6) & "'"
'                                        Case "B7"
'                                            cKeys.B7 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "B7"))
'                                            cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.B7) & "'"
'                                        Case "B8"
'                                            cKeys.B8 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "B8"))
'                                            cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.B8) & "'"
'                                        End Select
'
'                                            End If
'                                        End If
'
'
'                                    Else
'
'                                    End If
'                                    If IRC.colKeys.FindKey(from$) = -1 Then
'
'                                        Set cKeys = IRC.colKeys.Add("B" & IRC.colKeys.Count, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", from$, "ID" & IRC.colKeys.Count)
'                                        Set itx = lvKeys.ListItems.Add(, cKeys.Key, from$)
'                                        cDebug "IRC:// [UNKNOWN] Key for {" & from$ & "} not found in memory, creating memory allocations [" & "xH" & String(8 - Len(Hex(Val(cKeys.Key))), "0") & Hex(Val(cKeys.Key)) & "]"
'                                    Else
'                                        Set cKeys = IRC.colKeys(IRC.colKeys.FindKey(from$))
'                                        Set itx = lvKeys.ListItems(cKeys.Key)
'                                        cDebug "IRC:// [FOUND] Key for {" & from$ & "} found in memory, re-indexing [" & "xH" & String(8 - Len(Hex(Val(cKeys.Key))), "0") & Hex(Val(cKeys.Key)) & "]"
'                                    End If
'
'                                    Select Case Mid(Trim(processParam(processRest(params$))), 10, 2)
'                                    Case "C1"
'                                        itx.Tag = XMLTag(1, Trim(processParam(processRest(params$))), "C1")
'
'                                        If XMLTag(1, itx.Tag, "N2") <> "" Then
'                                            Set itmX = lvResend.ListItems.Add(, cKeys.Key, from$)
'                                            itmX.SubItems(1) = XMLTag(1, itx.Tag, "N2")
'                                            itmX.SubItems(2) = itx.SubItems(1)
'                                            cDebug "IRC:// [PASSON] request from {" & from$ & "} recieved, queing for resend, a little latter."
'                                        End If
'
'                                        If XMLTag(1, itx.Tag, "N3") <> "" Then
'                                            Set itmX = lvResend.ListItems.Add(, cKeys.Key, from$)
'                                            itmX.SubItems(1) = XMLTag(1, itx.Tag, "N2")
'                                            itmX.SubItems(2) = itx.SubItems(1)
'                                            cDebug "IRC:// [PASSON] request from {" & from$ & "} recieved, queing for resend, a little latter."
'                                        End If
'
'                                        If XMLTag(1, itx.Tag, "N4") <> "" Then
'                                            Set itmX = lvResend.ListItems.Add(, cKeys.Key, from$)
'                                            itmX.SubItems(1) = XMLTag(1, itx.Tag, "N2")
'                                            itmX.SubItems(2) = itx.SubItems(1)
'                                            cDebug "IRC:// [PASSON] request from {" & from$ & "} recieved, queing for resending to " & itx.SubItems(1)
'                                        End If
'
'                                    Case "D3"
'                                        itx.SubItems(1) = XMLTag(1, Trim(processParam(processRest(params$))), "D3")
'                                        cDebug "IRC:// [DATELH] recieved from {" & from$ & "} date is '" & itx.SubItems(1) & "'"
'                                    Case "D4"
'                                        itx.SubItems(2) = XMLTag(1, Trim(processParam(processRest(params$))), "D4")
'                                        send "PRIVMSG " + from$ + " :[KEYREV]<D1>" & Format(GetSetting(App.ProductName, "IRCKey", "Created", DateAdd("d", -1, sysnow)), "yyyy-mm-dd ttttt") & "</D1>" + _
'                                                                  "<D2>" & Format(sysnow, "yyyy-mm-dd ttttt") & "</D2>"
'                                        cDebug "IRC:// [DATERH] recieved from {" & from$ & "} date is '" & itx.SubItems(1) & "'"
'                                    Case "A1"
'                                        cKeys.A1 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "A1"))
'                                        cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.A1) & "'"
'                                    Case "A2"
'                                        cKeys.A2 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "A2"))
'                                        cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.A2) & "'"
'                                    Case "A3"
'                                        cKeys.A3 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "A3"))
'                                        cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.A3) & "'"
'                                    Case "A4"
'                                        cKeys.A4 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "A4"))
'                                        cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.A4) & "'"
'                                    Case "A5"
'                                        cKeys.A5 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "A5"))
'                                        cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.A5) & "'"
'                                    Case "A6"
'                                        cKeys.A6 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "A6"))
'                                        cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.A6) & "'"
'                                    Case "A7"
'                                        cKeys.A7 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "A7"))
'                                        cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.A7) & "'"
'                                    Case "A8"
'                                        cKeys.A8 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "A8"))
'                                        cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.A8) & "'"
'                                    Case "B1"
'                                        cKeys.B1 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "B1"))
'                                        cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.B1) & "'"
'                                    Case "B2"
'                                        cKeys.B2 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "B2"))
'                                        cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.B2) & "'"
'                                    Case "B3"
'                                        cKeys.B3 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "B3"))
'                                        cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.B3) & "'"
'                                    Case "B4"
'                                        cKeys.B4 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "B4"))
'                                        cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.B4) & "'"
'                                    Case "B5"
'                                        cKeys.B5 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "B5"))
'                                        cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.B5) & "'"
'                                    Case "B6"
'                                        cKeys.B6 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "B6"))
'                                        cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.B6) & "'"
'                                    Case "B7"
'                                        cKeys.B7 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "B7"))
'                                        cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.B7) & "'"
'                                    Case "B8"
'                                        cKeys.B8 = MySQL.NumDecrypt(XMLTag(1, Trim(processParam(processRest(params$))), "B8"))
'                                        cDebug "IRC:// [XERT" + Mid(Trim(processParam(processRest(params$))), 10, 2) + "] recieved from {" & from$ & "} sequence updated to '" & MySQL.NumCrypt(cKeys.B8) & "'"
'                                    End Select
'                                End If
'
'
'                    Case "JOIN" 'if someone joined
'                        displaychat "** " + from$ + " has joined " + processParam(params$) + " **"     'display it
'                        'check if the user is allready in the list
'                        X% = -1  'start checking from the first user (-1 + 1 = 0)
'                        Do
'                            X% = X% + 1     'increase x% with 1
'                            If X% = lstUsers.ListCount Then 'if the user is not found ...
'                                X% = -1     'set the user to be removed to -1 (ERROR :-) )
'                                Exit Do     'exit the loop
'                            End If
'                        Loop Until lstUsers.List(X%) = from$    'loop until we find the user
'                        'if x% = -1, the user was not found in the list, so we can add him
'                        If X% = -1 Then lstUsers.AddItem (from$)    'add this user to the user list
'                        lblCount.Caption = lstUsers.ListCount & " people in channel"    'update the user count
'
'                        For lx = 1 To lvKeys.ListItems.Count
'                            If lvKeys.ListItems(lx).Text = from$ Then
'                                send "PRIVMSG " + from$ + " :[CRCCHECK] Auto key re-enstatement gateway"
'                                Exit For
'                            End If
'                        Next
'
'                        'send "PRIVMSG " + from$ + " :[IRCKey] Request Key " & Format(Now, "yyyy-mm-dd ttttt")
'                        tmrPopulateKeys.Enabled = True
'
'                    Case "QUIT" 'if someone disconnected
'
'
'                        displaychat "** " + from$ + " has quited IRC (" + processParam(params$) + ") **"    'display it
'                        'check if the user is allready in the list
'                        X% = -1  'start checking from the first user (-1 + 1 = 0)
'                        Do
'                            X% = X% + 1     'increase x% with 1
'                            If X% = lstUsers.ListCount Then 'if the user is not found ...
'                                X% = -1     'set the user to be removed to -1 (ERROR :-) )
'                                Exit Do     'exit the loop
'                            End If
'                        Loop Until lstUsers.List(X%) = from$    'loop until we find the user
'                        If X% > -1 Then
'
'                            For k = lvKeys.ListItems.Count To 1 Step -1
'                                If lvKeys.ListItems(k).Text = lstUsers.List(X%) Then
'                                    lvKeys.ListItems(k).ForeColor = RGB(198, 0, 0)
'                                End If
'                            Next
'                            For k = lvResend.ListItems.Count To 1 Step -1
'                                If lvResend.ListItems(k).Text = lstUsers.List(X%) Then
'                                    lvResend.ListItems.Remove k
'
'                                End If
'                            Next
'                            For k = lstKeySend.ListCount - 1 To 1 Step -1
'                                If lstKeySend.List(k) = lstUsers.List(X%) Then
'
'                                    lstKeySend.RemoveItem k
'
'
'                                End If
'                            Next
'                            For k = lstKeyRqst.ListCount - 1 To 1 Step -1
'                                If lstKeyRqst.List(k) = lstUsers.List(X%) Then
'
'                                    lstKeySend.RemoveItem k
'
'                                End If
'                            Next
'                            lstUsers.RemoveItem (X%)    'if we found a matching user in the list, remove it
'                        End If
'                        lblCount.Caption = lstUsers.ListCount & " people in channel"    'update the user count
'                        'If IRC.colKeys.FindKey(from$) <> -1 Then IRC.colKeys.Remove IRC.colKeys.FindKey(from$)
'                    Case "NICK" 'if someone changed his nickname
'                        tmrPopulateKeys.Enabled = False
'                        displaychat "** " + from$ + " change his nickname to " + processParam(params$) + " **"    'display it
'                        'check if the user is allready in the list
'                        X% = -1  'start checking from the first user (-1 + 1 = 0)
'                        Do
'                            X% = X% + 1     'increase x% with 1
'                            If X% = lstUsers.ListCount Then 'if the user is not found ...
'                                X% = -1     'set the user to be removed to -1 (ERROR :-) )
'                                Exit Do     'exit the loop
'                            End If
'                        Loop Until lstUsers.List(X%) = from$    'loop until we find the user
'                        If X% > -1 Then
'                            'If IRC.colKeys.FindKey(from$) <> -1 Then IRC.colKeys.changenick from$, (processParam(params$))
'                            IRC.colKeys.changenick lstUsers.List(X%), (processParam(params$))
'                            Dim h As Long
'
'                            For h = lvKeys.ListItems.Count To 1 Step -1
'                                If lvKeys.ListItems(h).Text = lstUsers.List(X%) Then
'                                    lvKeys.ListItems(h).Text = (processParam(params$))
'                                    Exit For
'                                End If
'                            Next
'                            For h = lvResend.ListItems.Count To 1 Step -1
'                                If lvResend.ListItems(h).Text = lstUsers.List(X%) Then
'                                    lvResend.ListItems(h).Text = (processParam(params$))
'                                    Exit For
'                                End If
'                                If lvResend.ListItems(h).SubItems(1) = lstUsers.List(X%) Then
'                                    lvResend.ListItems(h).SubItems(1) = (processParam(params$))
'                                    Exit For
'                                End If
'                            Next
'                            For h = lstKeySend.ListCount - 1 To 1 Step -1
'                                If lstKeySend.List(h) = lstUsers.List(X%) Then
'
'                                    lstKeySend.RemoveItem h
'                                    lstKeySend.AddItem (processParam(params$))
'                                    Exit For
'                                End If
'                            Next
'                            For h = lstKeyRqst.ListCount - 1 To 1 Step -1
'                                If lstKeyRqst.List(h) = lstUsers.List(X%) Then
'
'                                    lstKeySend.RemoveItem h
'                                    lstKeySend.AddItem (processParam(params$))
'
'                                End If
'                            Next
'                            lstUsers.RemoveItem (X%)    'if we found a matching user in the list, remove it
'                            lstUsers.AddItem (processParam(params$))    'and add the new nick
'                        End If
'                        lblCount.Caption = lstUsers.ListCount & " people in channel"    'update the user count
'                        tmrPopulateKeys.Enabled = True
'                    Case "PART" ' if someone left the channel
'                        displaychat "** " + from$ + " has left " + params$ + " **"    'display it
'                        'check if the user is allready in the list
'                        X% = -1  'start checking from the first user (-1 + 1 = 0)
'                        Do
'                            X% = X% + 1
'                            If X% = lstUsers.ListCount Then 'if the user is not found ...
'                                X% = -1     'set the user to be removed to -1 (ERROR :-) )
'                                Exit Do     'exit the loop
'                            End If
'                        Loop Until lstUsers.List(X%) = from$    'loop until we find the user
'                        If X% > -1 Then
'
'                            For k = lvKeys.ListItems.Count To 1 Step -1
'                                If lvKeys.ListItems(k).Text = lstUsers.List(X%) Then
'                                    lvKeys.ListItems(k).ForeColor = RGB(255, 0, 0)
'                                End If
'                            Next
'                            For k = lvResend.ListItems.Count To 1 Step -1
'                                If lvResend.ListItems(k).Text = lstUsers.List(X%) Then
'                                    lvResend.ListItems.Remove k
'
'                                End If
'                            Next
'                            For k = lstKeySend.ListCount - 1 To 1 Step -1
'                                If lstKeySend.List(k) = lstUsers.List(X%) Then
'
'                                    lstKeySend.RemoveItem k
'
'
'                                End If
'                            Next
'                            For k = lstKeyRqst.ListCount - 1 To 1 Step -1
'                                If lstKeyRqst.List(k) = lstUsers.List(X%) Then
'
'                                    lstKeySend.RemoveItem k
'
'                                End If
'                            Next
'                            lstUsers.RemoveItem (X%)    'if we found a matching user in the list, remove it
'                        End If
'                        lblCount.Caption = lstUsers.ListCount & " people in channel"    'update the user count
'                    Case "MODE"     'if someone sets the mode on someone
'                        displaychat "** " + from$ + " sets mode " + processParam(processRest(params$)) + " on " + processParam(params$) + " **" 'display the mode change
'                    Case "TOPIC"    'if the topic message is received
'                        displaychat "TOPIC MESSAGE:"
'                        displaychat processParam(params$)             'Display the channel topic
'                    Case "331"  'if you recieve a message saying "no topic set"
'                        displaychat "No topic set in " + processParam(processRest(params$)) 'display it
'                            'by displaying the second parameter
'                    Case "353"  'if we received the channel user list
'                        display "<" + from$ + "> " + rest$ 'display the unprocessed message
'                        Dim nick2$, othernicks$    'take one nick at a time
'                        othernicks$ = processParam(processRest(processRest(processRest(params$))))   'cut of the channel parameter, the nick parameter and the "="
'                        Do
'                            nick2$ = processParam(othernicks$)   'take one nick
'                            othernicks$ = processRest(othernicks$)   'and take it out of the remaining nicks
'                            Do Until Left$(nick2$, 1) <> "@" And Left$(nick2$, 1) <> "+"  'cut of the @ and + flags at the beginning ...
'                                nick2$ = Right(nick2$, Len(nick2$) - 1) 'cut of the first character
'                            Loop
'                            X% = -1  'start checking from the first user (-1 + 1 = 0)
'                            Do
'                                X% = X% + 1     'increase x% with 1
'                                If X% = lstUsers.ListCount Then 'if the user is not found ...
'                                    X% = -1     'set the user to be removed to -1 (ERROR :-) )
'                                    Exit Do     'exit the loop
'                                End If
'                            Loop Until lstUsers.List(X%) = nick2$    'loop until we find the user
'                            'if x% = -1, the user was not found in the list, so we can add him
'                            If X% = -1 Then lstUsers.AddItem (nick2$)    'add this user to the user list
'                        Loop Until othernicks$ = ""     'loop through all the received nicknames
'                        lblCount.Caption = lstUsers.ListCount & " people in channel"    'update the user count
'                    Case "376"    'end of the motd
'                        display "<" + from$ + "> " + rest$ 'display the unprocessed message
'                        send "JOIN " + channel$ 'join the channel
'                    Case "431"  'if we failed to change the nickname
'                        nick$ = oldnick$    'change it back to the old one
'                        display "<!> Failed changing nickname (You have to supply a nickname)" 'let them know that it failed
'                        txtNick.Text = nick$    'change the content of the nick text field back
'                    Case "432"  'if we failed to change the nickname
'                        nick$ = oldnick$    'change it back to the old one
'                        display "<!> Failed changing nickname (The nickname you entered is not valid)" 'let them know that it failed
'                        txtNick.Text = nick$    'change the content of the nick text field back
'                    Case "433"  'if we failed to change the nickname
'                        nick$ = oldnick$    'change it back to the old one
'                        display "<!> Failed changing nickname (The nickname is allready in use)" 'let them know that it failed
'                        txtNick.Text = "[" & Login.lLevel & "]" & Login.sUsername & "-" & Format(Now, "ss")   'change the content of the nick text field back
'                        Call cmdNick_Click
'                    Case Else   'if it's another message
'                        display "<" + from$ + "> " + rest$ 'display the unprocessed message
'                End Select
'            Else   'if we failed
'                display "<" + from$ + "> " + rest$ 'display the unprocessed message
'            End If
'        Else    'if we failed
'            display "<" + from$ + "> " + rest$ 'display the unprocessed message
'        End If
'    End If
'
'    gSleep
'
'Exit Sub
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
'
Function processParam(Msg$) As String    'process a parameter (parse it from the other ones):

    If (Left$(Msg$, 1) = ":") Then  'if the parameter starts with a colon, the entire msg$ is a single parameter (containing spaces)
        processParam = Right$(Msg$, Len(Msg$) - 1)   'return the message, except for the colon
    Else    'if its not a multi word parameter
        If InStr(Msg$, " ") - 1 > 0 Then    'if there are any remaining parameters except the first one
            processParam = Mid$(Msg$, 1, InStr(Msg$, " ") - 1)    'return the part before the first space
        Else
            processParam = Msg$ 'if there is only one parameter in the string return it
        End If
    End If
End Function
Function processRest(Msg$) As String    'process the rest of the message:
    If (Left$(Msg$, 1) = ":") Then  'if the parameter starts with a colon, the entire msg$ is a single parameter (containing spaces)
        processRest = ""   'return nothing
    Else    'if its not a multi word parameter
        If InStr(Msg$, " ") > 0 Then
            processRest = Right$(Msg$, Len(Msg$) - InStr(Msg$, " "))   'return all parameters except the first one
        Else
            processRest = ""   'return nothing
        End If
    End If
End Function
Function processRest2(Msg$) As String    'process the rest of the message:
    If (Left$(Msg$, 1) = ":") Then  'if the parameter starts with a colon, the entire msg$ is a single parameter (containing spaces)
        processRest2 = ""   'return nothing
    Else    'if its not a multi word parameter
        If InStr(Msg$, " ") > 0 Then
            processRest2 = LTrim(Right$(Msg$, Len(Msg$) - InStr(Msg$, " ")))   'return all parameters except the first one
            Dim rdADO As adodb.Recordset
            Call MySQL.OpenTable(directConn, rdADO, , "select DECODE('" & Mid(processRest2, 6) & "','" & odb.colSalts.ReturnSalt("PublicKey") & "') as Result")
            If rdADO.State = adStateOpen Then
                If Mid(processRest2, 2, 3) = Left(rdADO!Result, 3) Then
                    If rdADO.RecordCount > 0 Then processRest2 = rdADO!Result
                Else
                    processRest2 = ""
                End If
            End If
        Else
            processRest2 = ""   'return nothing
        End If
    End If
End Function


Private Sub tmrDBMapper_Timer()

    DoEvents
    
    Static DoNow As Boolean
    
    If odb.colDBObjects.Count > 0 And DoNow = False Then
        If DateDiff(IIf(Right(lvSchedule.ListItems("mapdbstructure").SubItems(1), 1) = "m", "n", Right(lvSchedule.ListItems("mapdbstructure").SubItems(1), 1)), sysnow, CDate(lvSchedule.ListItems("mapdbstructure").Tag)) >= Val(Left(lvSchedule.ListItems("mapdbstructure").SubItems(1), Len(lvSchedule.ListItems("mapdbstructure").SubItems(1)) - 1)) Or DateDiff(IIf(Right(lvSchedule.ListItems("mapdbstructure").SubItems(1), 1) = "m", "n", Right(lvSchedule.ListItems("mapdbstructure").SubItems(1), 1)), CDate(lvSchedule.ListItems("mapdbstructure").Tag), sysnow) = 0 Then
            If DoNow = False Then odb.colDBObjects.Clear
        Else
            lvSchedule.ListItems("mapdbstructure").SubItems(3) = "480m"
            lvSchedule.ListItems("mapdbstructure").SubItems(4) = "Kept from previous cache"
            tmrDBMapper.Interval = 0
            Exit Sub
        End If
    Else
        DoNow = True
    End If
    
    Static rstSchema As adodb.Recordset
    Static iCount As Integer
    Static DBProgress As type_pbBar
    
    If rstSchema Is Nothing Then
    
        Set rstSchema = directConn.OpenSchema(adSchemaTables)
        Call odb.colDBObjects.Clear
        odb.colDBObjects.dbTables = rstSchema.RecordCount
        DBProgress.Max = DBProgress.Max + rstSchema.RecordCount
        If Not lvSchedule.ListItems("mapdbstructure").SubItems(3) = Round((DBProgress.Value / DBProgress.Max) * 100) & "%" Then lvSchedule.ListItems("mapdbstructure").SubItems(3) = Round((DBProgress.Value / DBProgress.Max) * 100) & "%"
        DoNow = False
    End If
    
    If rstSchema.EOF = True Then
        tmrDBMapper.Enabled = False
        tmrDBMapper.Interval = 0
        DBProgress.Value = 0
        DBProgress.Max = 1
        DBProgress.Min = 0
        rstSchema.Close
        
        lvSchedule.ListItems("mapdbstructure").SubItems(3) = "1200m"
        lvSchedule.ListItems("mapdbstructure").SubItems(4) = "" & odb.colDBObjects.dbTables & " tables + " & odb.colDBObjects.dbFields & " fields = " & (odb.colDBObjects.dbTables + odb.colDBObjects.dbFields) & " objects."
        
        Exit Sub
    End If
    
    Dim X As Long
    Static rsDesc As adodb.Recordset
    Static rsload As adodb.Recordset
    
    If rsDesc Is Nothing Then
        On Error GoTo ProfBuild
        bResult = MySQL.OpenTable(directConn, rsload, , "Select * from " & rstSchema!TABLE_NAME & " Limit 1,1", adOpenStatic, adLockReadOnly, True)
        bResult = MySQL.OpenTable(directConn, rsDesc, , "describe " & rstSchema!TABLE_NAME & "", adOpenStatic, adLockReadOnly, True)
        If Err.Number <> 0 Then GoTo ProfBuild
        iCount = 1
        tmrDBMapper.Interval = 10
    Else
        iCount = iCount + 1
        If Not tmrDBMapper.Interval = 7 Then tmrDBMapper.Interval = 2
        If iCount > rsload.Fields.Count Then
            rstSchema.MoveNext
            Set rsDesc = Nothing
            Set rsload = Nothing
            If DBProgress.Value + 1 < DBProgress.Max Then DBProgress.Value = DBProgress.Value + 1
            If Not lvSchedule.ListItems("mapdbstructure").SubItems(3) = Round((DBProgress.Value / DBProgress.Max) * 100) & "%" Then lvSchedule.ListItems("mapdbstructure").SubItems(3) = Round((DBProgress.Value / DBProgress.Max) * 100) & "%"
            Exit Sub
        End If
    End If
    
          
    Select Case LCase(Left(rstSchema!TABLE_NAME, 3))

    
    Case Else
     
        If rsload.State = adStateOpen Then
                
           lvSchedule.ListItems("mapdbstructure").SubItems(4) = "Building Database Profile [`" & "projectalpha" & "`.`" & Left(rstSchema!TABLE_NAME, 3) + Right(rstSchema!TABLE_NAME, 2) & "`.`" & rsload.Fields(iCount - 1).Name & "`]"
           
           
           
           If rsDesc.State = adStateOpen Then
               rsDesc.Filter = "Field = '" & rsload.Fields(iCount - 1).Name & "'"
               Call odb.colDBObjects.Add("f_" & X & "_" & rstSchema!TABLE_NAME, Login.lLevel, IIf(rsDesc!Null = "YES", True, False), IIf(rsDesc!Key = "PRI" And rsDesc!Extra = "auto_increment", True, False), IIf(IsNull(rsDesc!Extra), "", rsDesc!Extra), IIf(IsNull(rsDesc!Default), "", rsDesc!Default), IIf(IsNull(rsDesc!Key), "", rsDesc!Key), "projectalpha", rstSchema!TABLE_NAME, rsload.Fields(iCount - 1).Name, 0, rsload.Fields(iCount - 1).DefinedSize, rsload.Fields(iCount - 1).NumericScale, rsload.Fields(iCount - 1).Precision, rsload.Fields(iCount - 1).Status, rsload.Fields(iCount - 1).Type, MySQL.fldType(rsload.Fields(iCount - 1).Type), rsload.Fields(iCount - 1).Attributes)
           Else
               Call odb.colDBObjects.Add("f_" & X & "_" & rstSchema!TABLE_NAME, Login.lLevel, True, IIf(rstSchema!fIELD_NAME = "RecID", True, False), "", _
                                         "NULL", "", "projectalpha", rstSchema!TABLE_NAME, rsload.Fields(iCount - 1).Name, 0, rsload.Fields(iCount - 1).DefinedSize, rsload.Fields(iCount - 1).NumericScale, rsload.Fields(iCount - 1).Precision, rsload.Fields(iCount - 1).Status, rsload.Fields(iCount - 1).Type, MySQL.fldType(rsload.Fields(iCount - 1).Type), rsload.Fields(iCount - 1).Attributes)
           End If
           
           
           
        End If
             
    End Select
    
    
    DoEvents
    
    
    If Not lvSchedule.ListItems("mapdbstructure").SubItems(3) = Round((DBProgress.Value / DBProgress.Max) * 100) & "%" Then lvSchedule.ListItems("mapdbstructure").SubItems(3) = Round((DBProgress.Value / DBProgress.Max) * 100) & "%"
    
    Exit Sub
                 
ProfBuild:

If Err.Number <> 0 Then

     Select Case Err.Number
     Case -2147467259 ' Axxess is Denied
        Err.Clear
        
        Call odb.colDBObjects.Add("f_" & X & "_" & rstSchema!TABLE_NAME, 0, False, False, "", _
                                        "", "", "projectalpha", rstSchema!TABLE_NAME, "Access Denied", 0, 0, 0, 0, 0, 0, "Access Denied", 0)
        
     Case Else
        
        Call odb.colDBObjects.Add("f_" & X & "_" & rstSchema!TABLE_NAME, 0, False, False, "", _
                                        "", "", "projectalpha", rstSchema!TABLE_NAME, "Error " & Err.Number, 0, 0, 0, 0, 0, 0, Err.Description, 0)
        Err.Clear
        
        
     End Select

Else


    Call odb.colDBObjects.Add("f_" & X & "_" & rstSchema!TABLE_NAME, 0, False, False, "", _
         "", "", "projectalpha", rstSchema!TABLE_NAME, "General Error", 0, 0, 0, 0, 0, 0, "Could Not Select Table", 0)
                                
End If

    


    If Not lvSchedule.ListItems("mapdbstructure").SubItems(3) = Round((DBProgress.Value / DBProgress.Max) * 100) & "%" Then lvSchedule.ListItems("mapdbstructure").SubItems(3) = Round((DBProgress.Value / DBProgress.Max) * 100) & "%"
        
End Sub

Private Sub MDIForm_Load()


    '*[ Error Checking Variables ]**********************************************************************************
    Const RoutineName = "MDIForm_Load"
    Const ContainerName = "frmMDIMain"
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
    Call tsSchedule_Click
    Call ts2_Click
    
    mnuDrop.Visible = False
    
    
    frmMDIMain.Picture = frmSplash.picBg.Image
    frmSplash.bFinished = True
    
    
    
    Call mnuWindow_Back_Click
    
    Me.Show
    
    gSleep
    
    SetFormAccess
    
    
    
    If bDebug = -1 Then
        On Error GoTo 0
    ElseIf bDebug = 1 Then
        On Error Resume Next
    Else
        On Error GoTo ErrorOccur
    End If

    Call tsSchedule_Click

    nick$ = Trim(Login.sUsername)     'fetch the nickname from the dialog
    oldnick$ = Trim("[" & IIf(Len("" & Login.lLevel) < 3, String(3 - Len("" & Login.lLevel), "0") & "" & Login.lLevel, "" & Login.lLevel) & "]" & Login.sUsername) 'fetch the alternative nickname from the dialog
    channel$ = Trim("#projectalpha")  'fetch the channel form the dialog
    
    If bBigFont = True Then
        lvSchedule.Font.Size = 16
        sbFooter.Font.Size = 16
        sbFooter.Font.Bold = True
    End If
    
    Picture1.Visible = GetSetting("projectalpha", "Main", "ServerStatus", Login.bMaster)
    mnuWindow_ServiceStatus.Checked = Picture1.Visible
                 
'    frmAbout.Show 1
        
    picColumn.Width = GetSetting("projectalpha", "Main", "ScheduleWidth", picColumn.Width)
    
    Call tsSchedule_Click
    
    LoadColumnWidths
    
    
    If Login.bMaster = True Then
        'Set fLines = New frmLines
        'fLines.Show
    End If
    
    frmMDIMain_Loaded = True
    Me.bRefresh = True
    
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

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "MDIForm_MouseDown"
    Const ContainerName = "frmMDIMain"
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


    Static FstX As Single
    Static FstY As Single
    Static Clip As Boolean
    
    If InStr(Command, "/mapcord") > 0 Then
        If Button = 1 Then
            If FstX = 0 Then
                If Clip = False Then Clipboard.Clear
                Clip = True
                FstX = X
                FstY = Y
            Else
                st = "If x => " & FstX & " and x <= " & X & " and Y >= " & FstY & " and Y <= " & Y & " then "
                st = st + vbCrLf & vbTab & " shell " & Chr$(34) & "explorer.exe " & InputBox("Web To Visit") & Chr$(34)
                st = st + vbCrLf & "end if"
                On Error Resume Next
                'Clipboard.GetText gh
                'gh = gh + vbCrLf + st
                'Clipboard.SetText gh
                FstX = 0
                FstY = 0
            End If
        End If
    End If
    
    
    If X >= 315 And X <= 2085 And Y >= 1335 And Y <= 1815 Then
         Shell "explorer.exe http://www.ep.net.au/"
    End If

    If X >= 405 And X <= 1935 And Y >= 2130 And Y <= 3960 Then
         Shell "explorer.exe http://www.ep.net.au/"
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

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "MDIForm_MouseMove"
    Const ContainerName = "frmMDIMain"
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


    
    If X >= 315 And X <= 2085 And Y >= 1335 And Y <= 1815 Then
         Me.MousePointer = vbCrosshair
         Exit Sub
    End If

    If X >= 405 And X <= 1935 And Y >= 2130 And Y <= 3960 Then
         Me.MousePointer = vbCrosshair
         Exit Sub
    End If
    
    Me.MousePointer = vbDefault
    
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

Private Sub MDIForm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "MDIForm_MouseUp"
    Const ContainerName = "frmMDIMain"
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


    picDrag.Drag 0

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

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "MDIForm_QueryUnload"
    Const ContainerName = "frmMDIMain"
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

    Dim itmX As ListItem
    
    For Each itmX In lvSchedule.ListItems
        If Val(Left(itmX.SubItems(3), Len(itmX.SubItems(3)) - 1)) > 0 And Val(Left(itmX.SubItems(3), Len(itmX.SubItems(3)) - 1)) < 100 Then
            Select Case MsgBox("You should not quit at this time as a schedule task is in progress. Do you wish to quit anyway?", vbCritical & vbYesNo, "Process in action")
            Case vbNo
                Cancel = True
                Exit Sub
            End Select
            
        End If
    Next
    
    If frmAgent.Tag = "Processing" Then
        Cancel = True
        Exit Sub
    End If
    
    

    'send "QUIT Later Peeps - Project Alpha " & App.Major & "." & App.Minor & ".0." & App.Revision
    
    SaveSetting "projectalpha", "Main", "ServerStatus", Picture1.Visible
    SaveSetting "projectalpha", "Main", "ScheduleWidth", picColumn.Width
    
    SaveRegistry
    
    If oErr.colError.Count > 0 Then
        Dim ix As Long
        Dim fComm As New mnuComm
        On Error Resume Next
        For ix = 1 To oErr.colError.Count
            fComm.Part2 = fComm.Part2 + "C/A: " & oErr.colError(ix).CaseStatement & vbCrLf
            fComm.Part2 = fComm.Part2 + "lbl: " & oErr.colError(ix).LBL & vbCrLf
            fComm.Part2 = fComm.Part2 + "routine: " & oErr.colError(ix).RoutineName & vbCrLf
            fComm.Part2 = fComm.Part2 + "container: " & oErr.colError(ix).ContainerName & vbCrLf
            fComm.Part2 = fComm.Part2 + "Server: " & Format(sysnow, "dddd, ddddd hh:nn:ss." & Format(IIf(Len("" & Me.msCount) <= 1, "0" & Me.msCount, "" & Me.msCount))) & vbCrLf
            fComm.Part2 = fComm.Part2 + "err occ: " & Format(oErr.colError(ix).DateTime, "dddd, ddddd ttttt") & vbCrLf
            fComm.Part2 = fComm.Part2 + "---------------------------------------------------------------------------------------"
            gSleep
            fComm.Part2 = fComm.Part2 + "err no: " & oErr.colError(ix).ErrNumber & vbCrLf
            fComm.Part2 = fComm.Part2 + "err hex: " & Hex(oErr.colError(ix).ErrNumber) & vbCrLf
            fComm.Part2 = fComm.Part2 + "err oct: " & Oct(oErr.colError(ix).ErrNumber) & vbCrLf
            fComm.Part2 = fComm.Part2 + "---------------------------------------------------------------------------------------"
        Next
        fComm.Subject = "Error Report"
        fComm.Show 1
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

Private Sub MDIForm_Unload(Cancel As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "MDIForm_Unload"
    Const ContainerName = "frmMDIMain"
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


    SaveColumnWidths
    
    frmAgent.Agent(frmAgent.Agent.UBound).Characters.Unload "CharacterID"
    directConn.Close
    
    Unload fIcon
    
    ShellLauncher
    
    End
    
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

Private Sub mnuAction_AccountList_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuAction_AccountList_Click"
    Const ContainerName = "frmMDIMain"
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


    Dim fHold As New frmAccHoldings
    
    fHold.Show
    'Call frmAccHoldings.PopulateList
    
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

Private Sub mnuAction_Accounts_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuAction_Accounts_Click"
    Const ContainerName = "frmMDIMain"
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


    Load frmAccounts
    frmAccounts.loadRS
    
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

Private Sub mnuAction_Cert_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuAction_Cert_Click"
    Const ContainerName = "frmMDIMain"
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

Private Sub mnuAction_Comm_Bug_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuAction_Comm_Bug_Click"
    Const ContainerName = "frmMDIMain"
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


    Dim fmnuComm As New mnuComm
    fmnuComm.Subject = "Bug Report"
    fmnuComm.Show 1
    
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

Private Sub mnuAction_Comm_feature_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuAction_Comm_feature_Click"
    Const ContainerName = "frmMDIMain"
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


    Dim fmnuComm As New mnuComm
    fmnuComm.Subject = "Feature Request"
    fmnuComm.Show 1
    
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

Private Sub mnuAction_Commission_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuAction_Commission_Click"
    Const ContainerName = "frmMDIMain"
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


    frmCommissions.Show
    
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

Private Sub mnuAction_InvoiceSystem_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuAction_InvoiceSystem_Click"
    Const ContainerName = "frmMDIMain"
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


    frmStatements.Show
        
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

Private Sub mnuAction_Money_Exp_Click()

    Dim fExp As New frmEXP
    
    fExp.Show
    
End Sub

Private Sub mnuAction_Night_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuAction_Night_Click"
    Const ContainerName = "frmMDIMain"
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


    frmNight.Show 1
    
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

Private Sub mnuAction_Refund_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuAction_Refund_Click"
    Const ContainerName = "frmMDIMain"
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


    Dim ofrmRefunds As New frmRefunds
    ofrmRefunds.Show
    
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

Private Sub mnuAction_Relogin_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuAction_Relogin_Click"
    Const ContainerName = "frmMDIMain"
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


    frmLogin.Show 1
    SetFormAccess
    
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

Private Sub mnuAction_VISPS_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuAction_VISPS_Click"
    Const ContainerName = "frmMDIMain"
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


    frmVISP.Show
    
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

Private Sub mnuAction_whois_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuAction_whois_Click"
    Const ContainerName = "frmMDIMain"
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


    frmWhois.Show 1
    
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

Private Sub mnuAgency_assign_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuAgency_assign_Click"
    Const ContainerName = "frmMDIMain"
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


    Dim fSite As New frmAgSites
    
   fSite.Show 1
    
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

Private Sub mnuAgency_create_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuAgency_create_Click"
    Const ContainerName = "frmMDIMain"
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


    Dim fAgency As New frmAgency
    
    
    fAgency.Show 1
    
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

Private Sub mnuCust_lvPlans_billingDate_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuCust_lvPlans_billingDate_Click"
    Const ContainerName = "frmMDIMain"
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


 If Me.frmCust.lvPlans.SelectedItem Is Nothing Then
    
    Else
        Select Case MsgBox("Are you sure you want to set the cycle date to " + Format(sysnow, "dddd dd-mm-yyyy Hh:Nn:Ss") & "?", vbQuestion + vbYesNo, "Set Cycle Date")
        Case vbYes
      
            MySQL.Execute directConn, "Update acci_services Set NextCycle='" + Format(sysnow, "yyyy-mm-dd Hh:Nn:Ss") + "' where RecID = " & CLng(Mid(Me.frmCust.lvPlans.SelectedItem.Key, 2))
            
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

Private Sub mnuCust_lvPlans_Delete_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuCust_lvPlans_Delete_Click"
    Const ContainerName = "frmMDIMain"
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


    If Not Me.frmCust.lvPlans.SelectedItem Is Nothing Then
    
        Select Case MsgBox("Are you sure you wish to delete this service?", vbYesNo, "Delete Plan or Service")
        Case vbYes
            
            MySQL.Execute directConn, "Delete from acci_services where RecID = " & Mid(Me.frmCust.lvPlans.SelectedItem.Key, 2)
            Me.frmCust.lvPlans.ListItems.Remove Me.frmCust.lvPlans.SelectedItem.Index
            
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

Private Sub mnuCust_lvPlans_Password_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuCust_lvPlans_Password_Click"
    Const ContainerName = "frmMDIMain"
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


    If Me.frmCust.lvPlans.SelectedItem Is Nothing Then
    
    Else
        Dim fPass As New frmPassword
        fPass.Show 1
                
        If fPass.pass <> "" Then
        
            MySQL.Execute directConn, "Update acci_services Set Password=AES_ENCRYPT('" & MySQL.ESC(fPass.pass) & "','" + odb.colSalts.ReturnSalt("md5Password") + "') where RecID = " & CLng(Mid(Me.frmCust.lvPlans.SelectedItem.Key, 2))
        
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

Private Sub mnuCust_lvPlans_Properties_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuCust_lvPlans_Properties_Click"
    Const ContainerName = "frmMDIMain"
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


    If Me.frmCust.lvPlans.SelectedItem Is Nothing Then
    
    Else
        Dim fProp As New frmPlanProperties
        fProp.acciRecID = IIf(Me.frmCust.osub.fRecID = 0, 0, Me.frmCust.osub.fRecID)
        fProp.RecID = CLng(Mid(Me.frmCust.lvPlans.SelectedItem.Key, 2))
        fProp.Show 1
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

Private Sub mnuCust_lvPlans_SetActivation_Click()

    Dim fDate As New frmDateTime
    Dim rsAct As adodb.Recordset
    
    Call MySQL.OpenTable(directConn, rsAct, , "select Activation from acci_services where RecID = '" & Mid(Me.frmCust.lvPlans.SelectedItem.Key, 2) & "'")
    
    If rsAct.State = adStateOpen Then
        If rsAct.RecordCount > 0 Then
            If Not IsNull(rsAct!Activation) Then
                fDate.dDate = rsAct!Activation
            Else
                fDate.dDate = sysnow
            End If
        End If
    End If
    
    fDate.Show
    
    Select Case fDate.Cancel
    Case False
        
        Me.frmCust.lvPlans.SelectedItem.Checked = True
        Call MySQL.Execute(directConn, "update acci_services set Activation = '" & Format(fDate.dDate, "yyyy-mm-dd ttttt") & "', Checked = '" & IIf(Me.frmCust.lvPlans.SelectedItem.Checked = True, -1, 0) & "' where RecID = '" & Mid(Me.frmCust.lvPlans.SelectedItem.Key, 2) & "'")
        
    End Select
    
        
End Sub

Private Sub mnuOptions_Maintenance_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuOptions_Maintenance_Click"
    Const ContainerName = "frmMDIMain"
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


    frmBot.Show
    
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

Private Sub mnuOptions_setting_AccountType_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuOptions_setting_AccountType_Click"
    Const ContainerName = "frmMDIMain"
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


     frmAccountTypes.Show 1
    
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

Private Sub mnuOptions_settings_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuOptions_settings_Click"
    Const ContainerName = "frmMDIMain"
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


    Dim fOptions As New frmOptions
    Set fOptions.frmMDI = Me
    fOptions.Show 1
    
    Dim itmX  As ListItem
    
    If Login.bMaster = True Then
    
        Set itmX = lvSchedule.ListItems("radius")
        itmX.SubItems(1) = "" & reg.iRadiusHistory & "m"
        Set itmX = lvSchedule.ListItems("b0")
        itmX.SubItems(1) = "" & reg.iUpkeep & "m"
        Set itmX = lvSchedule.ListItems("kunpaid")
        itmX.SubItems(1) = "" & reg.iUnpaid & "m"
        Set itmX = lvSchedule.ListItems("q0")
        itmX.SubItems(1) = "" & reg.iSendPO & "m"
        Set itmX = lvSchedule.ListItems("uConstant")
        itmX.SubItems(1) = "" & reg.iUpdate & "m"
        
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

Private Sub TabStrip1_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "TabStrip1_Click"
    Const ContainerName = "frmMDIMain"
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

Private Sub mnuOptions_View_Debug_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuOptions_View_Debug_Click"
    Const ContainerName = "frmMDIMain"
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


    frmDebug.Show
    
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

Private Sub mnuOptions_View_Tasks_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuOptions_View_Tasks_Click"
    Const ContainerName = "frmMDIMain"
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


    If picColumn.Visible = True Then picColumn.Visible = False Else picColumn.Visible = True
    mnuOptions_View_Tasks.Checked = picColumn.Visible
    
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

Private Sub mnuReports_Bill_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuReports_Bill_Click"
    Const ContainerName = "frmMDIMain"
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
    
    If Login.bTestBench = True Then
        cr.Connect = "DSN=TESTBENCH;UID=ROOT;PWD=JOHN;DSQ=projectalpha"
    Else
        cr.Connect = "DSN=project alpha;UID=project alphareport;PWD=report;DSQ=projectalpha"
    End If
    
    cr.ReportFileName = IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\") + "Bill.rpt"
    cr.WindowTitle = "Account Reciable - Un/Partially Paid"
    cr.Action = 1
    
    If Err.Number <> 0 Then
        
        cDebug Err.Description
        Err.Clear
        
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

Private Sub mnuSettings_Radius_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuSettings_Radius_Click"
    Const ContainerName = "frmMDIMain"
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


    frmPools.Show 1
    
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

Private Sub mnuSettings_Stationary_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuSettings_Stationary_Click"
    Const ContainerName = "frmMDIMain"
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


    frmStationary.Show
    
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

Private Sub mnuSettings_Templates_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuSettings_Templates_Click"
    Const ContainerName = "frmMDIMain"
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


    frmTemplateConfig.Show
    
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

Private Sub mnuWindow_AutoSMS_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuWindow_AutoSMS_Click"
    Const ContainerName = "frmMDIMain"
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


    frmMessages.Show 1
    
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

Private Sub mnuSettings_vendors_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuSettings_vendors_Click"
    Const ContainerName = "frmMDIMain"
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


    frmVendors.Show
    
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

Private Sub mnuSupplier_Manage_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuSupplier_Manage_Click"
    Const ContainerName = "frmMDIMain"
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


    frmSupplier.Show
    
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

Private Sub mnuWindows_About_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuWindows_About_Click"
    Const ContainerName = "frmMDIMain"
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


    Dim fAbout As New frmAbout
    
    fAbout.Show 1
    
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

Private Sub mnuWindows_sysops_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuWindows_sysops_Click"
    Const ContainerName = "frmMDIMain"
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


    frmMSgDay.Show 1
    
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
    Const ContainerName = "frmMDIMain"
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
    
    picColumn.Width = ((Me.Width + Me.Left) - Point.X) - picDrag.Width / 2
    
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

Private Sub picColumn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "picColumn_MouseUp"
    Const ContainerName = "frmMDIMain"
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


    picDrag.Drag 0

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

Private Sub picColumn_Resize()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "picColumn_Resize"
    Const ContainerName = "frmMDIMain"
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


    If picColumn.Height > 200 And picColumn.Width > 250 Then
        On Error Resume Next
        tsSchedule.Move tsSchedule.Left, tsSchedule.Top, picColumn.Width - tsSchedule.Left * 2, picColumn.Height - tsSchedule.Top * 2
        Dim ix As Integer
        For ix = picOptions.LBound To picOptions.UBound
            picOptions(ix).Move tsSchedule.ClientLeft, tsSchedule.ClientTop, tsSchedule.ClientWidth, tsSchedule.ClientHeight
        Next
        
        picDrag.Move 0, 0, picDrag.Width, picColumn.ScaleHeight
        
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

Private Sub picDrag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "picDrag_MouseDown"
    Const ContainerName = "frmMDIMain"
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

Private Sub picDrag_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "picDrag_MouseUp"
    Const ContainerName = "frmMDIMain"
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

Private Sub picOptions_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "picOptions_MouseUp"
    Const ContainerName = "frmMDIMain"
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



    picDrag.Drag 0

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

Private Sub picOptions_Resize(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "picOptions_Resize"
    Const ContainerName = "frmMDIMain"
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


    If picOptions(Index).ScaleWidth > 200 And picOptions(Index).ScaleHeight > 200 Then
        Select Case Index
        Case 0
        
            'Scriptlet1.Height = picOptions(0).ScaleHeight - Scriptlet1.Top - 15 - 1
            picOptions(Index).Refresh
            gSleep
            'Scriptlet1.Height = Scriptlet1.Height + 1
        Case 1
            On Error Resume Next
            
            lvSchedule.Move lvSchedule.Left, lvSchedule.Top, picOptions(Index).ScaleWidth - lvSchedule.Left * 2, picOptions(Index).ScaleHeight - 90 * 2 - cmdRunEvent(0).Height
            
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



Private Sub Picture1_Resize()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Picture1_Resize"
    Const ContainerName = "frmMDIMain"
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


    If Picture1.ScaleWidth > 500 And Picture1.ScaleHeight > 1200 Then
        txtDebug.Move 10, 10, Picture1.ScaleWidth - 20, Picture1.ScaleHeight - 20
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

Function ircConnect(ByVal strHostName As String, ByVal nRemotePort As Integer) As Boolean
    Dim nResult As Integer
    Dim strResult As String

    FtpConnect = False
    
    strHostName = Trim(strHostName)
    If Len(strHostName) = 0 Then Exit Function
    If nRemotePort = 0 Then nRemotePort = IPPORT_FTP
    
    '
    ' Initialize the command socket; this is the connection
    ' that all commands will be sent over
    '
    sIRC.AddressFamily = AF_INET
    sIRC.AutoResolve = True
    sIRC.KeepAlive = True
    sIRC.Blocking = True
    sIRC.BufferSize = 9086
    sIRC.Protocol = IPPROTO_IP
    sIRC.SocketType = SOCK_STREAM
    sIRC.LocalPort = IPPORT_ANY
    
    sIRC.HostName = strHostName
    sIRC.RemotePort = nRemotePort
    sIRC.Timeout = 30000 ' 60 seconds

    '
    ' Attempt the connection
    '
    If sIRC.Connect() <> 0 Then
        Exit Function
    End If

    tmrIRC.Enabled = False
    tmrIRC.Interval = 10000

    send "PASS none"    ' according to the rfc it's better to send this before sending a nick / user
    send "NICK " + nick$    ' send the nick message
    send "USER " & nick$ & " " & sockIRC.LocalIP & " The Nexus: dynamic billing system"   ' the user message
        ' USER <username>            <hostname>       <servername>    <real name>

    ircConnect = True
    
End Function

Private Sub sysTime_Timer()

    Static StartUp As Boolean
    Dim rsNOW As adodb.Recordset
    
    Static diff As Long
    
    
    
    If bRefresh = True Then
        bRefresh = False
        If MySQL.OpenTable(directConn, rsNOW, , "select NOW() as sysNow") = True Then
            If rsNOW.State = adStateOpen Then
                If rsNOW.RecordCount > 0 Then
                    
                    sysnow = rsNOW!sysnow
                    
                    diff = DateDiff("s", sysnow, sysnow)
                    
                End If
            End If
        
        End If
    End If
    
    Select Case StartUp
    Case False
    
        If MySQL.OpenTable(directConn, rsNOW, , "select NOW() as sysNow") = True Then
            If rsNOW.State = adStateOpen Then
                If rsNOW.RecordCount > 0 Then
                    
                    sysnow = rsNOW!sysnow
                    
                    diff = DateDiff("s", sysnow, sysnow)
                    
                End If
            End If
        
        End If
        msCount = msCount + 1
        StartUp = True
        bRefresh = False
        
    Case True
        
        msCount = msCount + 1
        Select Case msCount
        Case 10
            
            sysnow = DateAdd("s", 1, sysnow)
            sbFooter.Panels(2).Text = "Server: " & Format(sysnow, "HH:NN:SS ddd, ddddd")
            sbFooter.Panels(3).Text = "Local: " & Format(Now, "HH:NN:SS ddd, ddddd")
            msCount = 0
            
        Case Else
        
            If Not sbFooter.Panels(2).Text = "Server: " & Format(sysnow, "HH:NN:SS ddd, ddddd") Then sbFooter.Panels(2).Text = "Server: " & Format(sysnow, "HH:NN:SS ddd, ddddd")
            If Not sbFooter.Panels(3).Text = "Local: " & Format(Now, "HH:NN:SS ddd, ddddd") Then sbFooter.Panels(3).Text = "Local: " & Format(Now, "HH:NN:SS ddd, ddddd")
            
            If diff > 0 Then
                If diff + 120 < DateDiff("s", sysnow, sysnow) Then bRefresh = True
            Else
                If diff - 120 > DateDiff("s", sysnow, sysnow) Then bRefresh = True
            End If
            
            
        End Select
        
    End Select
End Sub

Private Sub tmrStatusPanel_Timer()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "tmrStatusPanel_Timer"
    Const ContainerName = "frmMDIMain"
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


    Static lLoginSysopID As Variant
    
    If lLoginSysopID <> Login.lSysopID Then
        lLoginSysopID = Login.lSysopID
        sbFooter.Panels(1).Text = "[" & Login.lSysopID & "] - " & Login.sUsername
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


Private Sub tmSchedule_Timer()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "tmSchedule_Timer"
    Const ContainerName = "frmMDIMain"
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


    Static iTimer(32000) As Variant
    
    
    
    Dim ix As Integer
    If lvSchedule.ListItems.Count > 0 Then
        For ix = 1 To lvSchedule.ListItems.Count
            If Right(lvSchedule.ListItems(ix).SubItems(1), 1) = "m" Then
                iTimer(ix) = iTimer(ix) + 0.5
                If Val(Left(lvSchedule.ListItems(ix).SubItems(1), Len(lvSchedule.ListItems(ix).SubItems(1)) - 1)) <= iTimer(ix) / 60 Then
                    RunBot ix
                    iTimer(ix) = 0
                    Me.bRefresh = True
                End If
            ElseIf InStr(lvSchedule.ListItems(ix).SubItems(1), ":") > 0 Then
                If lvSchedule.ListItems(ix).SubItems(1) = Format(sysnow, "Hh:Nn:SS") Then
                    RunBot ix
                    Me.bRefresh = True
                End If
            End If
        Next ix
    End If
    
    gSleep 500 * Rnd
    
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

Private Sub toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "toolbar_ButtonClick"
    Const ContainerName = "frmMDIMain"
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
    Dim RSsUBS As adodb.Recordset
    Dim fCustomerRec As New frmCustomerRec
    
    Select Case LCase(Button.Key)
    Case "state"
        If Login.lLevel >= 75 Then
            Call mnuAction_InvoiceSystem_Click
        End If
    Case "comm"
        If Login.lLevel >= 10 Then
            Call mnuAction_Commission_Click
        End If
    Case "acchold"
        If Login.lLevel >= 10 Then
            Call mnuAction_AccountList_Click
        End If
    Case "relogin"
        Call mnuAction_Relogin_Click
    Case "recieve"
        If Login.lLevel >= 75 Then
            Call mnuAction_Accounts_Click
        End If
    Case "addcust"
        Call mnuNewCust_Click
    Case "tools"
        
    Case "graphs"
    
    Case "scansale"
    
    Case "resource"
        
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


Public Function LoadSchedule()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "LoadSchedule"
    Const ContainerName = "frmMDIMain"
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


    
    Dim itmX As ListItem
    Dim iCounter As Integer
    
    eSchedule = UpdateUserFile
    
    
    'If MainBots.Count > 0 Then
    
    '    Dim iX as Variant
    '    For iX = MainBots.Count To 1 Step -1
    '        MainBots.Remove (iX)
    '    Next
        
    'End If
    
    'MainBots.Add "b0", BillingCycle, "30m"
    
    lvSchedule.ListItems.Clear
        
    If lvSchedule.ColumnHeaders(1).Width < 600 Then lvSchedule.ColumnHeaders(1).Width = 600
    If lvSchedule.ColumnHeaders(2).Width < 900 Then lvSchedule.ColumnHeaders(2).Width = 900
    If lvSchedule.ColumnHeaders(3).Width < 1600 Then lvSchedule.ColumnHeaders(3).Width = 1600
    If lvSchedule.ColumnHeaders(4).Width < 1000 Then lvSchedule.ColumnHeaders(4).Width = 1000
    If lvSchedule.ColumnHeaders(5).Width < 7500 Then lvSchedule.ColumnHeaders(5).Width = 7500
    
'    If ilSchedule(0).ListImages.Count > ilSchedule(1).ListImages.Count Then
'        If ilSchedule(1).ListImages.Count = 0 Then
'            ilSchedule(1).ImageHeight = 18
'            ilSchedule(1).ImageWidth = 18
'        End If
'        Dim lstImage As ListImage
'        On Error Resume Next
'        For Each lstImage In ilSchedule(0).ListImages
'            With ilSchedule(1).ListImages.Add(, lstImage.Key, lstImage.ExtractIcon)
'
'            End With
'        Next
'    End If
    
    Set lvSchedule.Icons = ilSchedule(0)
    Set lvSchedule.SmallIcons = ilSchedule(0)
    
    Call GUI.SaveColWidths(lvSchedule, Me)
    Call GUI.LoadColWidths(lvSchedule, Me)
        
    With lvSchedule.ListItems.Add(, "uRecIdentity", "", , 1)
        .SubItems(1) = "360m"
        .SubItems(2) = "Search for new Identity Records"
        .SubItems(3) = "0%"
        .SubItems(4) = "Waiting for timeout"
    End With
        
    Dim indx1 As Integer, indx2 As Integer, indx3 As Integer
    
    With lvSchedule.ListItems.Add(, "mapvisp", "", , 1)
        .SubItems(1) = "250m"
        .SubItems(2) = "Mapping VISP Levels in Database"
        .SubItems(3) = "0%"
        .SubItems(4) = "Run First Time"
        indx1 = .Index
    End With
    
    With lvSchedule.ListItems.Add(, "mapproducts", "", , 1)
        .SubItems(1) = "150m"
        .SubItems(2) = "Mapping Product Tree in Database"
        .SubItems(3) = "0%"
        .SubItems(4) = "Run First Time"
        indx2 = .Index
    End With
        
    With lvSchedule.ListItems.Add(, "mapcat", "", , 1)
        .SubItems(1) = "200m"
        .SubItems(2) = "Mapping Category Tree in Database"
        .SubItems(3) = "0%"
        .SubItems(4) = "Run First Time"
        indx3 = .Index
    End With
    
    Dim indx4 As Integer
    
    With lvSchedule.ListItems.Add(, "mapdbstructure", "", , 1)
        .SubItems(1) = "480m"
        .SubItems(2) = "Mapping Database Structure"
        .SubItems(3) = "0%"
        .SubItems(4) = "Run First Time"
        .Tag = sysnow
        indx4 = .Index
    End With

    
    With lvSchedule.ListItems.Add(, "delhistory", "", , 1)
        .SubItems(1) = "00:00:01"
        .SubItems(2) = "Delete old history"
        .SubItems(3) = "0%"
        .SubItems(4) = "Waiting for timeout"
    End With
    
    Select Case Login.bRunMaintenance
    Case True
    
        With lvSchedule.ListItems.Add(, "kunpaid", "", , 1)
            .SubItems(1) = "" & reg.iUnpaid & "m"
            .SubItems(2) = "Search for Unpaid Customers"
            .SubItems(3) = "0%"
            .SubItems(4) = "Waiting for timeout"
        End With
            
        With lvSchedule.ListItems.Add(, "b0", "", , 1)
            .SubItems(1) = "" & reg.iUpkeep & "m"
            .SubItems(2) = "Check For Billing Cycle"
            .SubItems(3) = "0%"
            .SubItems(4) = "Waiting for timeout"
        End With
        
        'MainBots.Add "radius", RetrueveRadius, reg.iRadiusHistory & "m"
        With lvSchedule.ListItems.Add(, "radius", "", , 1)
            .SubItems(1) = "" & reg.iRadiusHistory & "m"
            .SubItems(2) = "Download Radius History"
            .SubItems(3) = "0%"
            .SubItems(4) = "Waiting for timeout"
        End With
        
        With lvSchedule.ListItems.Add(, "q0", "", , 1)
            .SubItems(1) = "" & reg.iSendPO & "m"
            .SubItems(2) = "Purchases orders and Quotas"
            .SubItems(3) = "0%"
            .SubItems(4) = "Waiting for timeout"
        End With
        
        With lvSchedule.ListItems.Add(, "uConstant", "", , 1)
            .SubItems(1) = "" & reg.iUpdate & "m"
            .SubItems(2) = "Update Users on Radius Server"
            .SubItems(3) = "0%"
            .SubItems(4) = "Waiting for timeout"
        End With
        
        If Login.bMaster = True Then
            'MainBots.Add "u0", SendUserList, "00:00:00"
            With lvSchedule.ListItems.Add(, "u" & iCounter, "", , 1)
                .SubItems(1) = "00:00:00"
                .SubItems(2) = "Update User File"
                .SubItems(3) = "0%"
                .SubItems(4) = "Waiting for timeout"
            End With
            
            Dim rsload As adodb.Recordset
            Dim bResult As Boolean
            
            bResult = MySQL.OpenTable(directConn, rsload, , "select distinct Activate from radiusaccounts Where AutoActivateFlag <> 0")
            
            If rsload.RecordCount > 0 Then
            
                While Not rsload.EOF And Err.Number = 0
                    iCounter = iCounter + 1
                    
                    'MainBots.Add "u" & iCounter, SendUserList, Format(rsLoad!Activate, "Hh:Nn:SS")
                    
                    Set itmX = lvSchedule.ListItems.Add(, "u" & iCounter, "", , 1)
                    itmX.SubItems(1) = Format(rsload!Activate, "Hh:Nn:SS")
                    itmX.SubItems(2) = "Update User File for activation"
                    itmX.SubItems(3) = "0%"
                    rsload.MoveNext
                Wend
        
            End If
            
            bResult = MySQL.OpenTable(directConn, rsload, , "select distinct Deactivate from radiusaccounts Where AutoActivateFlag <> 0")
            
            If rsload.RecordCount > 0 Then
            
                While Not rsload.EOF And Err.Number = 0
                    iCounter = iCounter + 1
                    Set itmX = lvSchedule.ListItems.Add(, "u" & iCounter, "", , 1)
                    itmX.SubItems(1) = Format(rsload!Deactivate, "Hh:Nn:SS")
                    
                    'MainBots.Add "u" & iCounter, SendUserList, Format(rsLoad!DeActivate, "Hh:Nn:SS")
                    
                    itmX.SubItems(2) = "Update User File for deactivation"
                    itmX.SubItems(3) = "0%"
                    rsload.MoveNext
                Wend
        
            End If
        End If
    
    End Select
    
    RunBot indx1
    RunBot indx2
    RunBot indx3
    RunBot indx4
    
    Exit Function


    
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

Private Sub tsSchedule_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "tsSchedule_MouseUp"
    Const ContainerName = "frmMDIMain"
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



    picDrag.Drag 0

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

Public Function SaveColumnWidths()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "SaveColumnWidths"
    Const ContainerName = "frmMDIMain"
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
    
    Call GUI.SaveColWidths(lvSchedule, Me)
        
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
    Const ContainerName = "frmMDIMain"
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


    Call GUI.LoadColWidths(lvSchedule, Me)
    

    

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

Public Sub RunBot(ix As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "RunBot"
    Const ContainerName = "frmMDIMain"
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

    Dim allchecked As Long
    
    Select Case lvSchedule.ListItems(ix).Key
    Case "mapdbstructure"
            
        tmrDBMapper.Enabled = True
            
    Case "mapcat"

    
        Dim rsCatMap As adodb.Recordset
        'Dim xNode As Node
        'Count
        'MsgBox MySQL.virtualisp("select * from exp_categories where SecLevel <= '" & Login.lLevel & "' order by RecID", "exp_categories", False, False)
        
        GUI.mapCategory.Clear
        
        Call MySQL.OpenTable(directConn, rsCatMap, , MySQL.virtualisp("select distinct exp_categories.* from exp_categories where SecLevel <= '" & Login.lLevel & "'", "exp_categories", False, False))
        
        If rsCatMap.State = adStateOpen Then
            If rsCatMap.RecordCount > 0 Then
                Do
                    With GUI.mapCategory.Add("r" & rsCatMap!RecID, rsCatMap!RecID, rsCatMap!SubRecID, rsCatMap!VirtualID, rsCatMap!SysopID, rsCatMap!Icon, IIf(IsNull(rsCatMap!Description), "not set (null)", rsCatMap!Description), rsCatMap!formcode, rsCatMap!SecLevel, "r" & rsCatMap!RecID)
                        lvSchedule.ListItems(ix).SubItems(3) = ((GUI.mapCategory.Count / rsCatMap.RecordCount) * 100) & "%"
                        lvSchedule.ListItems(ix).SubItems(4) = "Category Found: " & .Description
                        lvSchedule.Refresh
                    End With
                    
                    gSleep 23
                    rsCatMap.MoveNext
                Loop Until rsCatMap.EOF
                rsCatMap.Close
            End If
        End If
    
        lvSchedule.ListItems(ix).SubItems(3) = "100%"
        lvSchedule.ListItems(ix).SubItems(4) = GUI.mapCategory.Count & " Total Categories Found"

    Case "mapvisp"
    
        On Error Resume Next
        
        Me.Caption = "Searching for Reseller MAP"
    
        Dim rsPriMap As adodb.Recordset
                
        Call MySQL.OpenTable(directConn, rsPriMap, , "select distinct vispa.RecID as RecIDa, vispb.RecID as RecIDb, vispb.Description from accountviewer, virtualisp as vispa, virtualisp as vispb where vispa.RecID = vispb.VirtualID  and  vispb.VirtualID = '" & Login.lVirtualID & "'")
        gSleep 23
        ViSPMAP.Clear
        INClause = "("
        If rsPriMap.State = adStateOpen Then
        
            If rsPriMap.RecordCount > 0 Then
                While Not rsPriMap.EOF And Err.Number = 0
                    If rsPriMap!RecIDb <> Login.lVirtualID Then ViSPMAP.Add "r" & ViSPMAP.Count + 1, Val(rsPriMap!RecIDa), Val(rsPriMap!RecIDb), IIf(IsNull(rsPriMap!Description), "\-\", rsPriMap!Description), "r" & ViSPMAP.Count + 1
                    lvSchedule.ListItems(ix).SubItems(4) = "Found Reseller: " & IIf(IsNull(rsPriMap!Description), "[NULL]", rsPriMap!Description)
                    If InStr(INClause, "'" & rsPriMap!RecIDb & "'") = 0 Then INClause = INClause & "'" & rsPriMap!RecIDb & "',"
                    gSleep 23
                    rsPriMap.MoveNext
                Wend
            End If
            
         End If
        
        rsPriMap.Close
        
        
        If ViSPMAP.Count > 0 Then
            Do
                lvSchedule.ListItems(ix).SubItems(4) = "Reseller Map [ " & ResellerMAP.Count & " resellers found]"
                lvSchedule.Refresh
                allchecked = allchecked + 1
                
                Call MySQL.OpenTable(directConn, rsPriMap, , "select distinct vispa.RecID as RecIDa, vispb.RecID as RecIDb, vispb.Description from accountviewer, virtualisp as vispa, virtualisp as vispb where vispa.RecID = vispb.VirtualID and vispb.VirtualID = '" & Val(ViSPMAP("r" & allchecked).RecIDb) & "'")
                gSleep 23
                'ViSPMAP.Clear
                
                If rsPriMap.State = adStateOpen Then
                
                    If rsPriMap.RecordCount > 0 Then
                        While Not rsPriMap.EOF And Err.Number = 0
                            If rsPriMap!RecIDb <> Login.lVirtualID Then ViSPMAP.Add "r" & ViSPMAP.Count + 1, Val(rsPriMap!RecIDa), Val(rsPriMap!RecIDb), IIf(IsNull(rsPriMap!Description), "\-\", rsPriMap!Description), "r" & ViSPMAP.Count + 1
                            lvSchedule.ListItems(ix).SubItems(4) = "Found Reseller: " & IIf(IsNull(rsPriMap!Description), "[NULL]", rsPriMap!Description)
                            lvSchedule.Refresh
                            If InStr(INClause, "'" & rsPriMap!RecIDb & "'") = 0 Then INClause = INClause & "'" & rsPriMap!RecIDb & "',"
                            rsPriMap.MoveNext
                            gSleep 23
                        Wend
                    End If
                    
                 End If
                
                rsPriMap.Close
                gSleep 23
                lvSchedule.ListItems(ix).SubItems(3) = ((allchecked / ViSPMAP.Count) * 100) & "%"
                lvSchedule.Refresh
            Loop Until ViSPMAP.Count = allchecked Or Err.Number <> 0
        End If
            
        Call MySQL.OpenTable(directConn, rsPriMap, , "select RecID as RecIDa, RecID as RecIDb, Description from virtualisp where RecID = '" & Login.lVirtualID & "'")
        If rsPriMap.State = adStateOpen Then
            If rsPriMap.RecordCount >= 1 Then
                lvSchedule.ListItems(ix).SubItems(4) = "Found Reseller: " & IIf(IsNull(rsPriMap!Description), "[NULL]", rsPriMap!Description)
                lvSchedule.Refresh
                gSleep 23
                ViSPMAP.Add "r0", Val(rsPriMap!RecIDa), Val(rsPriMap!RecIDb), IIf(IsNull(rsPriMap!Description), "\-\", rsPriMap!Description), "r0"
                If InStr(INClause, "'" & rsPriMap!RecIDb & "'") = 0 Then INClause = INClause & "'" & rsPriMap!RecIDb & "',"
            End If
         End If
        rsPriMap.Close
        
        Me.Caption = "The Nexus - version [" & App.Major & "." & App.Minor & ".0." & App.Revision & "]"
        INClause = Left(INClause, Len(INClause) - 1) & ")"
        MySQL.Execute directConn, "update sysops set INClause = '" & MySQL.ESC(CStr(INClause)) & "' where RecID = '" & Login.lSysopID & "'", False
        lvSchedule.ListItems(ix).SubItems(3) = "100%"
        lvSchedule.Refresh
        lvSchedule.ListItems(ix).SubItems(4) = ViSPMAP.Count & " Total Resellers Nodes Found"
                
    Case "mapproducts"
    
        Me.Caption = "Splining Query Reference Chart - 0 refs"
        NDECHRT.Clear
        Dim rsSecMap As adodb.Recordset
        Dim svrNode As Long
        Dim subIndx As Long
        
        lvSchedule.ListItems(ix).SubItems(4) = "Mapping Products"
        lvSchedule.Refresh
        Call MySQL.OpenTable(directConn, rsSecMap, , "select servicetypes.RecID, servicetypes.ServiceKey, servicetypes.Description, servicetypes.SubofRecID, servicetypes.ListOnRadius, servicetypes.HasUID, servicetypes.HasSysUID, servicetypes.BillImmediately, servicetypes_matrix.SubServiceID  from servicetypes inner join servicetypes_matrix on servicetypes.RecID = servicetypes_matrix.ServiceID and servicetypes_matrix.VirtualID = '" & Login.lVirtualID & "'")
        
        For allchecked = 0 To ViSPMAP.Count - 1
            
            lvSchedule.ListItems(ix).SubItems(3) = ((allchecked / ViSPMAP.Count) * 50) & "%"
            
            Call MySQL.OpenTable(directConn, rsPriMap, , "select plantypes.RecID, servicetypes.description as svrDesc, servicetypes.SubofRecID, plantypes.Description, plantypes.ServiceID, vendors.vName, plantypes.VendorID, plantypes.CatNo, plantypes.VirtualID from plantypes inner join servicetypes on plantypes.ServiceID = servicetypes.RecID inner join vendors on plantypes.VendorID = vendors.RecID where plantypes.VirtualID = '" & ViSPMAP("r" & allchecked).RecIDb & "' order by SubofRecID, svrDesc")
                
            If rsPriMap.State = adStateOpen Then
                If rsPriMap.RecordCount > 0 Then
                    While Not rsPriMap.EOF And Err.Number = 0
                        svrNode = 0
                        svrNode = NDECHRT.FindKey("svr_" & "r" & ViSPMAP("r" & allchecked).RecIDb & "_" & rsPriMap!ServiceID)
                        If svrNode = 0 Then
                            NDECHRT.Add "r" & ViSPMAP("r" & allchecked).RecIDb, IIf(IsNull(rsPriMap!svrDesc), "/--/", rsPriMap!svrDesc), "select distinct accountinfo.* from accountinfo, acci_services, plantypes Where accountinfo.RecID = acci_services.acci_RecID And acci_services.ptRecID = plantypes.RecID and plantypes.ServiceID = '" & rsPriMap!ServiceID & "' and plantypes.VirtualID = '" & ViSPMAP("r" & allchecked).RecIDb & "'", "select count(distinct accountinfo.*) as RecCount from accountinfo, acci_services, plantypes Where accountinfo.RecID = acci_services.acci_RecID And acci_services.ptRecID = plantypes.RecID and plantypes.ServiceID = '" & rsPriMap!ServiceID & "' and plantypes.VirtualID = '" & ViSPMAP("r" & allchecked).RecIDb & "'", 48, rsPriMap!VirtualID, Me.Count + 1, Val(svrNode), "svr_" & "r" & ViSPMAP("r" & allchecked).RecIDb & "_" & rsPriMap!ServiceID
                            ssvrNode = NDECHRT.FindKey("svr_" & "r" & ViSPMAP("r" & allchecked).RecIDb & "_" & rsPriMap!ServiceID)
                        End If
                        
                        lvSchedule.ListItems(ix).SubItems(4) = "Found Product: " & IIf(IsNull(rsPriMap!Description), "[NULL]", rsPriMap!Description)
                        lvSchedule.Refresh
                        NDECHRT.Add "svr_" & "r" & ViSPMAP("r" & allchecked).RecIDb & "_" & rsPriMap!ServiceID, IIf(IsNull(rsPriMap!Description), "/--/", rsPriMap!Description), "select distinct accountinfo.* from accountinfo inner join acci_services on accountinfo.RecID = acci_services.acci_RecID Where acci_services.ptRecID = '" & rsPriMap!RecID & "'", "select count (distinct accountinfo.*) from accountinfo, acci_services Where accountinfo.RecID = acci_services.acci_RecID And acci_services.ptRecID = '" & rsPriMap!RecID & "'", 12, rsPriMap!VirtualID, Me.Count + 1, Val(svrNode), "p" & rsPriMap!RecID & ""
                        subIndx = NDECHRT.FindKey("p" & rsPriMap!RecID)
                        NDECHRT.Add "p" & rsPriMap!RecID, IIf(IsNull(rsPriMap!Description), "/--/", "[Active] - " & rsPriMap!Description), "select distinct accountinfo.* from accountinfo, acci_services Where accountinfo.RecID = acci_services.acci_RecID And acci_services.ptRecID = '" & rsPriMap!RecID & "' and accountinfo.Cancelled = 0 ", "select count (distinct accountinfo.*) from accountinfo, acci_services Where accountinfo.RecID = acci_services.acci_RecID And acci_services.ptRecID = '" & rsPriMap!RecID & "' and accountinfo.Cancelled = 0 ", 18, rsPriMap!VirtualID, Me.Count + 1, Val(subIndx), "p" & rsPriMap!RecID & "_sub1"
                        NDECHRT.Add "p" & rsPriMap!RecID, IIf(IsNull(rsPriMap!Description), "/--/", "[Cancelled] - " & rsPriMap!Description), "select distinct accountinfo.* from accountinfo, acci_services Where accountinfo.RecID = acci_services.acci_RecID And acci_services.ptRecID = '" & rsPriMap!RecID & "' and accountinfo.Cancelled <> 0 ", "select count (distinct accountinfo.*) from accountinfo, acci_services Where accountinfo.RecID = acci_services.acci_RecID And acci_services.ptRecID = '" & rsPriMap!RecID & "' and accountinfo.Cancelled <> 0 ", 18, rsPriMap!VirtualID, Me.Count + 1, Val(subIndx), "p" & rsPriMap!RecID & "_sub2"
                        rsPriMap.MoveNext
                    Wend
                End If
            End If
            
            Me.Caption = "Splining Query Reference Chart - " & NDECHRT.Count & " refs"
        Next
    
        lvSchedule.ListItems(ix).SubItems(4) = "Mapping Templates"
        lvSchedule.Refresh
        
        NDECHRT.Add "troot", "Browse By Templates", MySQL.virtualisp("select distinct accountinfo.* from accountinfo, acci_services Where accountinfo.RecID = acci_services.acci_RecID", "accountinfo", True, Login.bMaster), "", 5, Login.lVirtualID, Me.Count + 1, 0, "troot"
        
        For allchecked = 0 To ViSPMAP.Count - 1
            
            lvSchedule.ListItems(ix).SubItems(3) = (50 + (allchecked / ViSPMAP.Count) * 50) & "%"
            lvSchedule.Refresh
            
            Call MySQL.OpenTable(directConn, rsPriMap, , "select plantemplates.RecID, servicetypes.description as svrDesc, servicetypes.SubofRecID, plantemplates.Description, plantemplates.ServiceID, vendors.vName, plantemplates.VendorID, plantypes.VirtualID, virtualisp.Description as ViSPDesc from virtualisp inner join plantypes on virtualisp.RecID = plantypes.VirtualID inner join servicetypes on plantypes.ServiceID = servicetypes.RecID inner join vendors on plantypes.VendorID = vendors.RecID inner join plantemplates on plantemplates.RecID = plantypes.TemplateID where plantypes.VirtualID = '" & ViSPMAP("r" & allchecked).RecIDb & "' order by SubofRecID, svrDesc")
            
            
            svrNode = 0
            
            If rsPriMap.State = adStateOpen Then
                
                If rsPriMap.RecordCount > 0 Then
                    While Not rsPriMap.EOF And Err.Number = 0
                        svrNode = NDECHRT.FindKey("troot_" & "visp" & rsPriMap!VirtualID)
                        If svrNode = 0 Then
                            svrNode = NDECHRT.FindKey("troot")
                            NDECHRT.Add "troot_" & "visp" & rsPriMap!VirtualID, IIf(IsNull(rsPriMap!ViSPDesc), "/--/", rsPriMap!ViSPDesc), "select distinct accountinfo.* from accountinfo, acci_services, plantypes, plantemplates Where plantypes.TemplateID = plantemplates.RecID and accountinfo.RecID = acci_services.acci_RecID And acci_services.ptRecID = plantypes.RecID and plantypes.VirtualID = '" & ViSPMAP("r" & allchecked).RecIDb & "'", "", 33, rsPriMap!VirtualID, Me.Count + 1, Val(svrNode), "troot_" & "visp" & rsPriMap!VirtualID
                            svrNode = NDECHRT.FindKey("troot_" & "visp" & rsPriMap!VirtualID)
                        End If
                    
                        
                        svrNode = NDECHRT.FindKey("TMPSVR" & ViSPMAP("r" & allchecked).RecIDb & "_SVR" & rsPriMap!ServiceID & "_VID" & rsPriMap!VirtualID)
                        If svrNode = 0 Then
                            svrNode = NDECHRT.FindKey("troot_" & "visp" & rsPriMap!VirtualID)
                            NDECHRT.Add "TMPSVR" & ViSPMAP("r" & allchecked).RecIDb & "_SVR" & rsPriMap!ServiceID & "_VID" & rsPriMap!VirtualID, IIf(IsNull(rsPriMap!svrDesc), "/--/", rsPriMap!svrDesc), "select distinct accountinfo.* from accountinfo, acci_services, plantypes, plantemplates Where plantypes.TemplateID = plantemplates.RecID and accountinfo.RecID = acci_services.acci_RecID And acci_services.ptRecID = plantypes.RecID and plantypes.ServiceID = '" & rsPriMap!ServiceID & "' and plantypes.VirtualID = '" & ViSPMAP("r" & allchecked).RecIDb & "'", "", 38, rsPriMap!VirtualID, Me.Count + 1, Val(svrNode), "TMPSVR" & ViSPMAP("r" & allchecked).RecIDb & "_SVR" & rsPriMap!ServiceID & "_VID" & rsPriMap!VirtualID
                            svrNode = NDECHRT.FindKey("TMPSVR" & ViSPMAP("r" & allchecked).RecIDb & "_SVR" & rsPriMap!ServiceID & "_VID" & rsPriMap!VirtualID)
                        End If
                        
                        rndNo = Rnd * 28713287632971#
                        NDECHRT.Add "tmp_" & "t" & ViSPMAP("r" & allchecked).RecIDb & "_" & rsPriMap!ServiceID & "_" & rsPriMap!RecID & "_" & rsPriMap!VirtualID & "_" & "TMPSVR" & ViSPMAP("r" & allchecked).RecIDb & "_SVR" & rsPriMap!ServiceID & "_VID" & rsPriMap!VirtualID, IIf(IsNull(rsPriMap!Description), "/--/", rsPriMap!Description), "select distinct accountinfo.* from accountinfo inner join acci_services on accountinfo.RecID = acci_services.acci_RecID inner join plantypes on plantypes.RecID = acci_services.ptRecID inner join plantemplates on plantemplates.RecID = plantypes.TemplateID Where plantemplates.RecID = '" & rsPriMap!RecID & "' and accountinfo.VirtualID = '" & rsPriMap!VirtualID & "'", "", 10, rsPriMap!VirtualID, Me.Count + 1, Val(svrNode), "tmp_" & "t" & ViSPMAP("r" & allchecked).RecIDb & "_" & rsPriMap!ServiceID & "_" & rsPriMap!RecID & "_" & rsPriMap!VirtualID & "_" & "TMPSVR" & ViSPMAP("r" & allchecked).RecIDb & "_SVR" & rsPriMap!ServiceID & "_VID" & rsPriMap!VirtualID & "_" & rndNo
                        subIndx = NDECHRT.FindKey("tmp_" & "t" & ViSPMAP("r" & allchecked).RecIDb & "_" & rsPriMap!ServiceID & "_" & rsPriMap!RecID & "_" & rsPriMap!VirtualID & "_" & "TMPSVR" & ViSPMAP("r" & allchecked).RecIDb & "_SVR" & rsPriMap!ServiceID & "_VID" & rsPriMap!VirtualID & "_" & rndNo)
                        NDECHRT.Add "p" & rsPriMap!RecID, IIf(IsNull(rsPriMap!Description), "/--/", "[Active] - " & rsPriMap!Description), "select distinct accountinfo.* from accountinfo inner join acci_services on accountinfo.RecID = acci_services.acci_RecID inner join plantypes on plantypes.RecID = acci_services.ptRecID inner join plantemplates on plantemplates.RecID = plantypes.TemplateID Where plantemplates.RecID = '" & rsPriMap!RecID & "' and accountinfo.Cancelled = 0  and accountinfo.VirtualID = '" & rsPriMap!VirtualID & "'", "", 18, rsPriMap!VirtualID, Me.Count + 1, Val(subIndx), "tmp_" & "t" & ViSPMAP("r" & allchecked).RecIDb & "_" & rsPriMap!ServiceID & "_" & rsPriMap!RecID & "_" & rsPriMap!VirtualID & "_" & "TMPSVR" & ViSPMAP("r" & allchecked).RecIDb & "_SVR" & rsPriMap!ServiceID & "_VID" & rsPriMap!VirtualID & "_sub1" & "_" & rndNo
                        NDECHRT.Add "p" & rsPriMap!RecID, IIf(IsNull(rsPriMap!Description), "/--/", "[Cancelled] - " & rsPriMap!Description), "select distinct accountinfo.* from accountinfo inner join acci_services on accountinfo.RecID = acci_services.acci_RecID inner join plantypes on plantypes.RecID = acci_services.ptRecID inner join plantemplates on plantemplates.RecID = plantypes.TemplateID Where plantemplates.RecID = '" & rsPriMap!RecID & "' and accountinfo.Cancelled <> 0  and accountinfo.VirtualID = '" & rsPriMap!VirtualID & "'", "", 18, rsPriMap!VirtualID, Me.Count + 1, Val(subIndx), "tmp_" & "t" & ViSPMAP("r" & allchecked).RecIDb & "_" & rsPriMap!ServiceID & "_" & rsPriMap!RecID & "_" & rsPriMap!VirtualID & "_" & "TMPSVR" & ViSPMAP("r" & allchecked).RecIDb & "_SVR" & rsPriMap!ServiceID & "_VID" & rsPriMap!VirtualID & "_sub2" & "_" & rndNo
                        
                        lvSchedule.ListItems(ix).SubItems(4) = "Found Used Template: " & IIf(IsNull(rsPriMap!Description), "[NULL]", rsPriMap!Description)
                        lvSchedule.Refresh
                        
                        rsPriMap.MoveNext
                        
                        
                    Wend
                End If
            End If
            
            Me.Caption = "Splining Query Reference Chart - " & NDECHRT.Count & " refs"
        Next
    
        Me.Caption = "The Nexus - version [" & App.Major & "." & App.Minor & ".0." & App.Revision & "]"
        lvSchedule.ListItems(ix).SubItems(4) = NDECHRT.Count & " Total Referenced Nodes, Products & Templates"
        lvSchedule.ListItems(ix).SubItems(3) = "100%"
        lvSchedule.Refresh
        
    Case Else
        
        If Login.bRunMaintenance = True Then
            Select Case Left(lvSchedule.ListItems(ix).Key, 1)
            Case "r"
                Load frmBot
                frmBot.cmdRadiusUpdate.Tag = lvSchedule.ListItems(ix).Key
                Call frmBot.cmdRadiusUpdate_Click
            Case "b"
                Load frmBot
                frmBot.cmdProcessBills.Tag = lvSchedule.ListItems(ix).Key
                Call frmBot.cmdProcessBills_Click
            Case "u"
                Load frmBot
                frmBot.cmdUploadUsers.Tag = lvSchedule.ListItems(ix).Key
                Call frmBot.cmdUploadUsers_Click
            Case "q"
                Load frmBot
                frmBot.Command1.Tag = lvSchedule.ListItems(ix).Key
                Call frmBot.Command1_Click
            Case "k"
                Load frmBot
                frmBot.cmdUnpaid.Tag = lvSchedule.ListItems(ix).Key
                Call frmBot.cmdUnpaid_Click
            End Select
        End If
        
    End Select
    
    lvSchedule.ListItems(ix).SubItems(3) = "100%"
    
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

Public Sub SetFormAccess()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "SetFormAccess"
    Const ContainerName = "frmMDIMain"
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


    lvSchedule.ListItems.Clear
    
    mnuSettings_Stationary.Visible = Login.bMaster
    
    
    mnuInt.Enabled = Login.bMaster
    
    mnuSettings_Vendors.Enabled = Login.bVendors
    
    'picColumn.Visible = Login.bMaster
    
    mnuOptions_Maintenance.Visible = Login.bRunMaintenance
    mnuAction_Sep1.Visible = Login.bRunMaintenance
    mnuSettings_Radius.Visible = Login.bMaster
    tmSchedule.Enabled = Login.bMaster
    mnuOptions_View_Tasks.Visible = Login.bMaster
    mnuOptions_settings.Enabled = True
    mnuAction_Accounts.Enabled = Login.bRecievables
    mnuAction_InvoiceSystem.Enabled = Login.bInvoice
    mnuAction_Money_Exp.Enabled = Login.bExpenditure
    mnuAction_AccountList.Enabled = Login.bHoldings
    mnuAction_Commission.Enabled = Login.bComm
    mnuNewCust.Enabled = Login.bAddCust
    
    'mnuOptions_View_Debug.Visible = Login.bMaster
    mnuAction_VISPS.Visible = Login.bVISP
    mnuSettings_Templates.Visible = Login.bTemplates
    'mnuAction_Sep3.Visible = Login.bMaster
    mnuAction_Refund.Enabled = Login.bRefund

    mnuWindows_Sysops.Enabled = Login.bMaster
    
    If Login.bCreateSysop = True Then mnuOptions.Enabled = True
    If Login.bCreateSysop = True Then mnuSettings_Vendors.Enabled = True
    
    mnuOptions_setting_AccountType.Enabled = Login.bAccSettings
    
    sbFooter.Panels(4).Text = Login.ViSPDesc
    
    ToolBar.Buttons("State").Enabled = mnuAction_InvoiceSystem.Enabled
    ToolBar.Buttons("comm").Enabled = mnuAction_Commission.Enabled
    ToolBar.Buttons("AccHold").Enabled = mnuAction_AccountList.Enabled
    ToolBar.Buttons("Recieve").Enabled = mnuAction_Accounts.Enabled
    
    Call LoadSchedule
    
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

