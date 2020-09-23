VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmTemplates 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Plan Templates"
   ClientHeight    =   8820
   ClientLeft      =   7560
   ClientTop       =   2880
   ClientWidth     =   10740
   Icon            =   "frmTemplates.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   10740
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvAccounts 
      Height          =   3075
      Left            =   90
      TabIndex        =   1
      Top             =   930
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   5424
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ilSML"
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "Description<No Format"
         Text            =   "Description"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "PeriodFee<Currency"
         Text            =   "Fee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "MBPerPeriod<No Format"
         Text            =   "Period Data"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "HoursPerPeriod<No Format"
         Text            =   "Period Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "SessionTimeout<No Format"
         Text            =   "Session Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "IdleTimeout<No Format"
         Text            =   "Idle"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ilSML 
      Left            =   4290
      Top             =   150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   110
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":075C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":0BAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":1000
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":1452
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":18A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":1CF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":2148
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":259A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":29EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":2E3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":3290
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":36E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":39FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":3E4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":42A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":46F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":4B44
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":4F96
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":53E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":583A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":5C8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":60DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":6530
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":6982
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":6DD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":7226
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":7678
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":7ACA
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":7F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":836E
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":87C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":8C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":9064
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":94B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":9908
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":9D5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":A1AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":A5FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":AA50
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":AEA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":B2F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":B746
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":BB98
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":BFEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":C43C
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":C88E
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":CCE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":D132
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":D584
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":D9D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":DE28
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":E27A
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":E6CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":EB1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":EF70
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":F3C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":F814
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":FC66
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":100B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":1050A
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":1095C
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":10DAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":11200
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":11652
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":11AA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":11EF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":12348
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":1279A
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":12BEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":1303E
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":13490
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":138E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":14034
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":14486
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":148D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":14D2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":1517C
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":155CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":158E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":15C02
            Key             =   ""
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":15F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":16236
            Key             =   ""
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":16550
            Key             =   ""
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":1686A
            Key             =   ""
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":16B84
            Key             =   ""
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":16E9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":171B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":174D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":177EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":17B06
            Key             =   ""
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":17E20
            Key             =   ""
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":1813A
            Key             =   ""
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":18454
            Key             =   ""
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":1876E
            Key             =   ""
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":18A88
            Key             =   ""
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":18DA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":190BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":193D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":196F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":19A0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":19D24
            Key             =   ""
         EndProperty
         BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":1A03E
            Key             =   ""
         EndProperty
         BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":1A358
            Key             =   ""
         EndProperty
         BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":1A7AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":1ABFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":1B04E
            Key             =   ""
         EndProperty
         BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":1B4A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":1B8F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplates.frx":1BD44
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTS 
      BorderStyle     =   0  'None
      Height          =   4035
      Index           =   1
      Left            =   180
      ScaleHeight     =   4035
      ScaleWidth      =   10395
      TabIndex        =   3
      Top             =   4560
      Width           =   10395
      Begin VB.Frame Frame2 
         Caption         =   "Vendor"
         Height          =   2175
         Left            =   90
         TabIndex        =   23
         Top             =   60
         Width           =   5175
         Begin VB.Frame Frame4 
            Caption         =   "Part Code"
            Height          =   765
            Left            =   120
            TabIndex        =   26
            Top             =   1290
            Width           =   4935
            Begin VB.TextBox txtField 
               BackColor       =   &H00859CFA&
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
               Index           =   3
               Left            =   2460
               Locked          =   -1  'True
               TabIndex        =   28
               Tag             =   "SubPartID"
               Top             =   270
               Width           =   2295
            End
            Begin VB.TextBox txtField 
               BackColor       =   &H004FD2F9&
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
               Index           =   2
               Left            =   90
               Locked          =   -1  'True
               TabIndex        =   27
               Tag             =   "VendorPartID"
               Top             =   270
               Width           =   2295
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Description"
            Height          =   1065
            Left            =   120
            TabIndex        =   24
            Top             =   210
            Width           =   4935
            Begin VB.TextBox txtField 
               Height          =   765
               Index           =   1
               Left            =   90
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   25
               Tag             =   "Description"
               Top             =   210
               Width           =   4755
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Costing"
         Height          =   2175
         Left            =   5400
         TabIndex        =   4
         Top             =   60
         Width           =   4875
         Begin VB.Line Line6 
            X1              =   1770
            X2              =   1770
            Y1              =   120
            Y2              =   1320
         End
         Begin VB.Line Line5 
            X1              =   3240
            X2              =   3240
            Y1              =   120
            Y2              =   1320
         End
         Begin VB.Line Line4 
            X1              =   30
            X2              =   4890
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Line Line3 
            X1              =   0
            X2              =   4830
            Y1              =   780
            Y2              =   780
         End
         Begin VB.Line Line2 
            X1              =   30
            X2              =   4830
            Y1              =   1620
            Y2              =   1620
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "10 Hours Limit"
            ForeColor       =   &H00EC7A71&
            Height          =   195
            Index           =   1
            Left            =   330
            TabIndex        =   22
            Top             =   1890
            Width           =   1005
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Inclusive of Tax"
            Height          =   195
            Index           =   11
            Left            =   3495
            TabIndex        =   21
            Top             =   1860
            Width           =   1125
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Exclusive of Tax"
            Height          =   195
            Index           =   10
            Left            =   1860
            TabIndex        =   20
            Top             =   1860
            Width           =   1170
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "$ 0.00"
            Height          =   195
            Index           =   9
            Left            =   4170
            TabIndex        =   19
            Top             =   1650
            Width           =   450
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "$ 0.00"
            Height          =   195
            Index           =   8
            Left            =   2580
            TabIndex        =   18
            Top             =   1650
            Width           =   450
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Extra Per Hour after limit:"
            Height          =   195
            Index           =   2
            Left            =   330
            TabIndex        =   17
            Top             =   1680
            Width           =   1740
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "1000Mb's Limit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00EC7A71&
            Height          =   195
            Index           =   0
            Left            =   1800
            TabIndex        =   16
            Top             =   1380
            Width           =   1305
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "100Mb Block Size"
            ForeColor       =   &H00EC7A71&
            Height          =   225
            Index           =   0
            Left            =   60
            TabIndex        =   15
            Top             =   1050
            Width           =   1605
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Inclusive of Tax"
            Height          =   195
            Index           =   7
            Left            =   3495
            TabIndex        =   14
            Top             =   1020
            Width           =   1125
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Exclusive of Tax"
            Height          =   195
            Index           =   6
            Left            =   1860
            TabIndex        =   13
            Top             =   1020
            Width           =   1170
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "$ 0.00"
            Height          =   195
            Index           =   5
            Left            =   4170
            TabIndex        =   12
            Top             =   810
            Width           =   450
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "$ 0.00"
            Height          =   195
            Index           =   4
            Left            =   2580
            TabIndex        =   11
            Top             =   810
            Width           =   450
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Cost Per MB Block:"
            Height          =   195
            Index           =   1
            Left            =   330
            TabIndex        =   10
            Top             =   840
            Width           =   1380
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Inclusive of Tax"
            Height          =   195
            Index           =   3
            Left            =   3495
            TabIndex        =   9
            Top             =   510
            Width           =   1125
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Exclusive of Tax"
            Height          =   195
            Index           =   2
            Left            =   1860
            TabIndex        =   8
            Top             =   510
            Width           =   1170
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "$ 0.00"
            Height          =   195
            Index           =   1
            Left            =   4170
            TabIndex        =   7
            Top             =   300
            Width           =   450
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "$ 0.00"
            Height          =   195
            Index           =   0
            Left            =   2580
            TabIndex        =   6
            Top             =   300
            Width           =   450
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Cost Per Cycle:"
            Height          =   195
            Index           =   0
            Left            =   330
            TabIndex        =   5
            Top             =   330
            Width           =   1080
         End
      End
   End
   Begin MSComctlLib.TabStrip ts 
      Height          =   4545
      Left            =   120
      TabIndex        =   2
      Top             =   4140
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   8017
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Product Text"
            Object.ToolTipText     =   "Here is the product text of this item"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Details"
            Object.ToolTipText     =   "Here is the details for the templates window"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTS 
      BorderStyle     =   0  'None
      Height          =   4095
      Index           =   0
      Left            =   60
      ScaleHeight     =   4095
      ScaleWidth      =   10515
      TabIndex        =   29
      Top             =   4440
      Width           =   10515
      Begin RichTextLib.RichTextBox rtfProd 
         Height          =   3885
         Left            =   150
         TabIndex        =   30
         Tag             =   "ProductText"
         Top             =   120
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   6853
         _Version        =   393217
         ReadOnly        =   -1  'True
         TextRTF         =   $"frmTemplates.frx":1C196
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0066B0DD&
      BorderWidth     =   2
      X1              =   30
      X2              =   11970
      Y1              =   510
      Y2              =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Plan Template"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   990
      TabIndex        =   0
      Top             =   0
      Width           =   2670
   End
   Begin VB.Image Image1 
      Height          =   2010
      Left            =   0
      Picture         =   "frmTemplates.frx":1C218
      Stretch         =   -1  'True
      Top             =   -480
      Width           =   2460
   End
End
Attribute VB_Name = "frmTemplates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ServiceID As Variant
Public vRecID As Variant
Dim rsPlanType As ADODB.Recordset

Private Sub Form_Load()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Load"
    Const ContainerName = "frmTemplates"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    
    If Login.bMaster = True Then
        bResult = MySQL.OpenTable(ADOConn, rsPlanType, , "select * from plantemplates Where ServiceID = " & Me.ServiceID)
    ElseIf Login.lLevel > 75 Then
        bResult = MySQL.OpenTable(ADOConn, rsPlanType, , MySQL.virtualisp("select * from plantemplates Where ServiceID = " & Me.ServiceID & " and Hidden = 0", "plantemplates", True, Login.bMaster))
    Else
        bResult = MySQL.OpenTable(ADOConn, rsPlanType, , MySQL.virtualisp("select * from plantemplates Where ServiceID = " & Me.ServiceID & " and Hidden = 0", "plantemplates", False, Login.bMaster))
    End If

    If rsPlanType.RecordCount > 0 Then
        While Not rsPlanType.EOF And Err.Number = 0
            Set itmX = lvAccounts.ListItems.Add(, "r" & rsPlanType!RecID, IIf(IsNull(rsPlanType!Description), "{Not Set}", rsPlanType!Description))
            
            If Not IsNull(rsPlanType!CategoryID) Then
                If rsPlanType!CategoryID > 0 Then
                    itmX.SmallIcon = GUI.mapCategory(GUI.mapCategory.FindKey("r" & rsPlanType!CategoryID)).Icon
                End If
            End If
            
            For X = 1 To lvAccounts.ColumnHeaders.Count - 1
                If lvAccounts.ColumnHeaders(X + 1).Tag <> "" Then
                    Select Case Mid(lvAccounts.ColumnHeaders(X + 1).Tag, InStr(lvAccounts.ColumnHeaders(X + 1).Tag, "<") + 1, Len(lvAccounts.ColumnHeaders(X + 1).Tag) - InStr(lvAccounts.ColumnHeaders(X + 1).Tag, "<"))
                    Case "No Format"
                        itmX.SubItems(X) = rsPlanType(Left(lvAccounts.ColumnHeaders(X + 1).Tag, InStr(lvAccounts.ColumnHeaders(X + 1).Tag, "<") - 1))
                        If InStr(itmX.SubItems(X), "-1") > 0 Then itmX.SubItems(X) = sSTR.ReplaceString(itmX.SubItems(X), "-1", "Unlimited")
                    Case Else
                        If rsPlanType(Left(lvAccounts.ColumnHeaders(X + 1).Tag, InStr(lvAccounts.ColumnHeaders(X + 1).Tag, "<") - 1)) <> -1 Then
                            itmX.SubItems(X) = Format(rsPlanType(Left(lvAccounts.ColumnHeaders(X + 1).Tag, InStr(lvAccounts.ColumnHeaders(X + 1).Tag, "<") - 1)), Mid(lvAccounts.ColumnHeaders(X + 1).Tag, InStr(lvAccounts.ColumnHeaders(X + 1).Tag, "<") + 1, Len(lvAccounts.ColumnHeaders(X + 1).Tag) - InStr(lvAccounts.ColumnHeaders(X + 1).Tag, "<")))
                            If InStr(itmX.SubItems(X), "-1") > 0 Then itmX.SubItems(X) = sSTR.ReplaceString(itmX.SubItems(X), "-1", "Unlimited")
                        End If
                    End Select
                End If
                gSleep
            Next X
            rsPlanType.MoveNext
        Wend
    End If

    ts.Tabs(2).Selected = True
    
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

Private Sub lvAccounts_DblClick()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvAccounts_DblClick"
    Const ContainerName = "frmTemplates"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    If lvAccounts.Tag = True Then
        vRecID = Mid(lvAccounts.SelectedItem.Key, 2)
        Unload Me
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

Private Sub lvAccounts_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvAccounts_ItemClick"
    Const ContainerName = "frmTemplates"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    lvAccounts.Tag = True
    
    rsPlanType.Filter = "RecID = " & Mid(Item.Key, 2)
    
    If rsPlanType.RecordCount > 0 Then
    
        Dim bx As Byte
        
        For bx = txtField.LBound To txtField.UBound
            txtField(bx).Text = IIf(IsNull(rsPlanType(txtField(bx).Tag)), "", rsPlanType(txtField(bx).Tag))
        Next
    
        rtfProd.TextRTF = IIf(IsNull(rsPlanType!ProductText), "", rsPlanType!ProductText)
        
        Label3(0).Caption = Format(rsPlanType!PeriodFee, "Currency")
        Label3(1).Caption = Format(rsPlanType!PeriodFee * oTax(Login.TaxCode, Login.TaxCountry) + rsPlanType!PeriodFee, "Currency")
        Label3(4).Caption = Format(rsPlanType!FeePerBlock, "Currency")
        Label3(5).Caption = Format(rsPlanType!FeePerBlock * oTax(Login.TaxCode, Login.TaxCountry) + rsPlanType!FeePerBlock, "Currency")
        Label3(8).Caption = Format(rsPlanType!ExtraPerHour, "Currency")
        Label3(9).Caption = Format(rsPlanType!ExtraPerHour * oTax(Login.TaxCode, Login.TaxCountry) + rsPlanType!ExtraPerHour, "Currency")
        If rsPlanType!MBBlockSize <> -1 Then
            Label4(0).Caption = rsPlanType!MBBlockSize & "Mb's Block Size"
        Else
            Label4(0).Caption = "Unlimited Block Size"
        End If
        If rsPlanType!HoursPerPeriod <> -1 Then
            Label4(1).Caption = rsPlanType!HoursPerPeriod & " Hours Limit"
        Else
            Label4(1).Caption = "Unlimited Hours"
        End If
        If rsPlanType!MBPerPeriod <> -1 Then
            Label5(0).Caption = rsPlanType!MBPerPeriod & "Mb's per cycle"
        Else
            Label5(0).Caption = "Unlimited Mb's per cycle"
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

Private Sub ts_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "ts_Click"
    Const ContainerName = "frmTemplates"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    
    For ix = picTS.LBound To picTS.UBound
        picTS(ix).Visible = False
    Next ix
    
    picTS(ts.SelectedItem.Index - 1).Move ts.ClientLeft, ts.ClientTop, ts.ClientWidth, ts.ClientHeight
    picTS(ts.SelectedItem.Index - 1).Visible = True
    picTS(ts.SelectedItem.Index - 1).ZOrder 0
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

