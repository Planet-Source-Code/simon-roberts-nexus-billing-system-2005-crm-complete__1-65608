VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCustomerRec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Entry"
   ClientHeight    =   9750
   ClientLeft      =   4500
   ClientTop       =   3135
   ClientWidth     =   9720
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCustomerRec_refrehs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   650
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   648
   Begin VB.PictureBox picBut 
      BorderStyle     =   0  'None
      Height          =   1395
      Index           =   2
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   1665
      TabIndex        =   101
      Tag             =   "<b>paymentconfig<b/>"
      Top             =   1410
      Width           =   1665
      Begin VB.Frame Frame17 
         Caption         =   "Payment Settings"
         Height          =   7605
         Left            =   60
         TabIndex        =   102
         Top             =   90
         Width           =   9375
         Begin VB.Frame Frame20 
            Caption         =   "Payment Pool and Order"
            Height          =   3345
            Left            =   120
            TabIndex        =   129
            Top             =   4140
            Width           =   9135
            Begin VB.CommandButton Command5 
               Enabled         =   0   'False
               Height          =   675
               Index           =   1
               Left            =   8310
               Picture         =   "frmCustomerRec_refrehs.frx":030A
               Style           =   1  'Graphical
               TabIndex        =   132
               Top             =   990
               Width           =   675
            End
            Begin VB.CommandButton Command5 
               Enabled         =   0   'False
               Height          =   675
               Index           =   0
               Left            =   8310
               Picture         =   "frmCustomerRec_refrehs.frx":074C
               Style           =   1  'Graphical
               TabIndex        =   131
               Top             =   300
               Width           =   675
            End
            Begin MSComctlLib.ListView lvPayment 
               Height          =   2925
               Left            =   120
               TabIndex        =   130
               Top             =   240
               Width           =   8085
               _ExtentX        =   14261
               _ExtentY        =   5159
               View            =   3
               Arrange         =   1
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               Checkboxes      =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   3
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Type"
                  Object.Width           =   3528
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Percentile"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Description"
                  Object.Width           =   8819
               EndProperty
            End
         End
         Begin VB.Frame Frame18 
            Caption         =   "Statement Watcher"
            Height          =   1125
            Index           =   2
            Left            =   120
            TabIndex        =   121
            Top             =   2970
            Width           =   9135
            Begin VB.CommandButton cmdPaymentTypeAdd 
               Caption         =   "&Add to Statement Watcher Pool"
               Height          =   315
               Index           =   2
               Left            =   5520
               TabIndex        =   128
               Top             =   690
               Width           =   3465
            End
            Begin VB.TextBox txtswNumber 
               Height          =   315
               Left            =   1590
               TabIndex        =   123
               Top             =   690
               Width           =   3795
            End
            Begin VB.TextBox txtswWord 
               Height          =   315
               Left            =   1590
               TabIndex        =   122
               Top             =   270
               Width           =   7395
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Number to Find:"
               Height          =   210
               Index           =   2
               Left            =   300
               TabIndex        =   125
               Top             =   750
               Width           =   1125
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Name to Find:"
               Height          =   210
               Index           =   2
               Left            =   300
               TabIndex        =   124
               Top             =   330
               Width           =   975
            End
         End
         Begin VB.Frame Frame18 
            Caption         =   "Credit Card Settings"
            Height          =   1485
            Index           =   1
            Left            =   120
            TabIndex        =   112
            Top             =   1440
            Width           =   9135
            Begin VB.CommandButton cmdPaymentTypeAdd 
               Caption         =   "&Add to Credit Card Pool"
               Height          =   315
               Index           =   1
               Left            =   5520
               TabIndex        =   126
               Top             =   1050
               Width           =   3465
            End
            Begin VB.TextBox txtccCIC 
               Height          =   315
               Left            =   3930
               TabIndex        =   119
               Top             =   1050
               Width           =   1455
            End
            Begin VB.TextBox txtccCardName 
               Height          =   315
               Left            =   1590
               TabIndex        =   115
               Top             =   270
               Width           =   7395
            End
            Begin VB.TextBox txtccCardNumber 
               Height          =   315
               Left            =   1590
               TabIndex        =   114
               Top             =   660
               Width           =   7395
            End
            Begin VB.TextBox txtccCardExpiry 
               Height          =   315
               Left            =   1590
               TabIndex        =   113
               Top             =   1050
               Width           =   1455
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "CIC:"
               Height          =   210
               Left            =   3240
               TabIndex        =   120
               Top             =   1110
               Width           =   285
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Card Name:"
               Height          =   210
               Index           =   1
               Left            =   300
               TabIndex        =   118
               Top             =   330
               Width           =   840
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Card Number:"
               Height          =   210
               Index           =   1
               Left            =   300
               TabIndex        =   117
               Top             =   720
               Width           =   990
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Card Expiry:"
               Height          =   210
               Left            =   300
               TabIndex        =   116
               Top             =   1110
               Width           =   885
            End
         End
         Begin VB.Frame Frame18 
            Caption         =   "Direct Debit Settings"
            Height          =   1125
            Index           =   0
            Left            =   120
            TabIndex        =   105
            Top             =   270
            Width           =   9135
            Begin VB.CommandButton cmdPaymentTypeAdd 
               Caption         =   "&Add to Direct Debit Pool"
               Height          =   315
               Index           =   0
               Left            =   5520
               TabIndex        =   127
               Top             =   660
               Width           =   3465
            End
            Begin VB.TextBox txtddAccountNumber 
               Height          =   315
               Left            =   3930
               TabIndex        =   110
               Top             =   660
               Width           =   1455
            End
            Begin VB.TextBox txtddBSB 
               Height          =   315
               Left            =   1590
               MaxLength       =   6
               TabIndex        =   108
               Top             =   660
               Width           =   1185
            End
            Begin VB.TextBox txtddAccountName 
               Height          =   315
               Left            =   1590
               TabIndex        =   106
               Top             =   270
               Width           =   7395
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Acc. Number:"
               Height          =   210
               Left            =   2910
               TabIndex        =   111
               Top             =   720
               Width           =   990
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "BSB:"
               Height          =   210
               Index           =   0
               Left            =   300
               TabIndex        =   109
               Top             =   720
               Width           =   360
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Account Name:"
               Height          =   210
               Index           =   0
               Left            =   300
               TabIndex        =   107
               Top             =   330
               Width           =   1110
            End
         End
      End
   End
   Begin MSComctlLib.ImageList iltb32x32 
      Left            =   2760
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerRec_refrehs.frx":0B8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerRec_refrehs.frx":0FE0
            Key             =   "resend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerRec_refrehs.frx":18BA
            Key             =   "audit"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerRec_refrehs.frx":1D0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerRec_refrehs.frx":215E
            Key             =   "info"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerRec_refrehs.frx":25B0
            Key             =   "network"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerRec_refrehs.frx":2E8A
            Key             =   "payment"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerRec_refrehs.frx":3764
            Key             =   "housing"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerRec_refrehs.frx":403E
            Key             =   "<b>FileDB<b/>"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tb 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   1376
      ButtonWidth     =   2408
      ButtonHeight    =   1376
      Style           =   1
      ImageList       =   "iltb32x32"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Customer Info"
            Description     =   "Here is where you can add and change the customer information"
            Object.Tag             =   "<b>customerinfo<b/>"
            ImageKey        =   "info"
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "File Manager"
            Object.Tag             =   "<b>FileDB<b/>"
            ImageKey        =   "<b>FileDB<b/>"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Payment Settings"
            Object.Tag             =   "<b>paymentconfig<b/>"
            ImageKey        =   "payment"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Audit History"
            Description     =   "This is the audit trail of the customer."
            Object.Tag             =   "<b>audithistory<b/>"
            ImageKey        =   "audit"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCC 
      Caption         =   "Set Credit Card"
      Height          =   315
      Left            =   2970
      TabIndex        =   20
      Top             =   9300
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7560
      TabIndex        =   96
      Top             =   9210
      Width           =   2085
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save And Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   90
      TabIndex        =   94
      Top             =   9210
      Width           =   2715
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   9510
      Top             =   1830
   End
   Begin VB.PictureBox picBut 
      BorderStyle     =   0  'None
      Height          =   2955
      Index           =   0
      Left            =   90
      ScaleHeight     =   2955
      ScaleWidth      =   2385
      TabIndex        =   22
      Tag             =   "<b>customerinfo<b/>"
      Top             =   5220
      Width           =   2385
      Begin VB.PictureBox picTSContainer 
         BorderStyle     =   0  'None
         Height          =   4455
         Index           =   1
         Left            =   120
         ScaleHeight     =   4455
         ScaleWidth      =   9405
         TabIndex        =   85
         Top             =   2790
         Visible         =   0   'False
         Width           =   9405
         Begin VB.CommandButton cmdAddEmail 
            Caption         =   "&Add e-Mail"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   87
            Top             =   1980
            Width           =   1545
         End
         Begin VB.CommandButton cmdAddPhone 
            Caption         =   "&Add Phone"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   86
            Top             =   4050
            Width           =   1515
         End
         Begin MSComctlLib.ListView lvEmail 
            Height          =   1905
            Left            =   90
            TabIndex        =   88
            Top             =   30
            Width           =   9195
            _ExtentX        =   16219
            _ExtentY        =   3360
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
               Name            =   "Arial"
               Size            =   8.25
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
            Picture         =   "frmCustomerRec_refrehs.frx":4D18
         End
         Begin MSComctlLib.ListView lvPhone 
            Height          =   1605
            Left            =   90
            TabIndex        =   89
            Top             =   2370
            Width           =   9195
            _ExtentX        =   16219
            _ExtentY        =   2831
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
               Name            =   "Arial"
               Size            =   8.25
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
            Picture         =   "frmCustomerRec_refrehs.frx":580A
         End
      End
      Begin VB.PictureBox SMTP1 
         Height          =   480
         Left            =   8880
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   134
         Top             =   2490
         Width           =   1200
      End
      Begin VB.PictureBox picHeader 
         BorderStyle     =   0  'None
         Height          =   2235
         Index           =   0
         Left            =   120
         ScaleHeight     =   2235
         ScaleWidth      =   9345
         TabIndex        =   30
         Top             =   450
         Width           =   9345
         Begin VB.TextBox txtAccountName 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2190
            TabIndex        =   0
            Top             =   60
            Width           =   7035
         End
         Begin VB.Frame Frame2 
            Caption         =   "How did you hear about Us?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Left            =   3540
            TabIndex        =   4
            Top             =   480
            Width           =   5715
            Begin VB.OptionButton optHearAbout 
               Caption         =   "Other Means"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Index           =   11
               Left            =   3780
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   1230
               Width           =   1815
            End
            Begin VB.OptionButton optHearAbout 
               Caption         =   "Just Walked In"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Index           =   10
               Left            =   3780
               Style           =   1  'Graphical
               TabIndex        =   15
               Top             =   930
               Width           =   1815
            End
            Begin VB.OptionButton optHearAbout 
               Caption         =   "Business"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Index           =   9
               Left            =   3780
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   630
               Width           =   1815
            End
            Begin VB.OptionButton optHearAbout 
               Caption         =   "A Friend"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Index           =   8
               Left            =   3780
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   330
               Width           =   1815
            End
            Begin VB.OptionButton optHearAbout 
               Caption         =   "News Paper"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Index           =   7
               Left            =   1950
               Style           =   1  'Graphical
               TabIndex        =   12
               Top             =   1230
               Width           =   1815
            End
            Begin VB.OptionButton optHearAbout 
               Caption         =   "Magazine"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Index           =   6
               Left            =   1950
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   930
               Width           =   1815
            End
            Begin VB.OptionButton optHearAbout 
               Caption         =   "Search Engine"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Index           =   5
               Left            =   1950
               Style           =   1  'Graphical
               TabIndex        =   10
               Top             =   630
               Width           =   1815
            End
            Begin VB.OptionButton optHearAbout 
               Caption         =   "Word of Mouth"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Index           =   0
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   330
               Width           =   1815
            End
            Begin VB.OptionButton optHearAbout 
               Caption         =   "Brochure"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Index           =   1
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   630
               Width           =   1815
            End
            Begin VB.OptionButton optHearAbout 
               Caption         =   "TV"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Index           =   2
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   7
               Top             =   930
               Width           =   1815
            End
            Begin VB.OptionButton optHearAbout 
               Caption         =   "Radio"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Index           =   3
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   8
               Top             =   1230
               Width           =   1815
            End
            Begin VB.OptionButton optHearAbout 
               Caption         =   "Internet"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Index           =   4
               Left            =   1950
               Style           =   1  'Graphical
               TabIndex        =   9
               Top             =   330
               Value           =   -1  'True
               Width           =   1815
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Web Access Logon Details"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Left            =   90
            TabIndex        =   1
            Top             =   480
            Width           =   3375
            Begin VB.TextBox txtgUID 
               Alignment       =   2  'Center
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   180
               TabIndex        =   2
               Top             =   330
               Width           =   3045
            End
            Begin VB.TextBox txtgPWD 
               Alignment       =   2  'Center
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               IMEMode         =   3  'DISABLE
               Left            =   180
               PasswordChar    =   "¤"
               TabIndex        =   3
               Top             =   990
               Width           =   3015
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "Username"
               Height          =   225
               Index           =   0
               Left            =   180
               TabIndex        =   32
               Top             =   720
               Width           =   3045
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "Password"
               Height          =   255
               Index           =   1
               Left            =   180
               TabIndex        =   31
               Top             =   1380
               Width           =   2985
            End
         End
         Begin VB.Label lblDescr 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Name:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   60
            TabIndex        =   33
            Top             =   90
            Width           =   1935
         End
      End
      Begin VB.PictureBox picHeader 
         BorderStyle     =   0  'None
         Height          =   2145
         Index           =   3
         Left            =   150
         ScaleHeight     =   2145
         ScaleWidth      =   9315
         TabIndex        =   49
         Top             =   480
         Width           =   9315
         Begin VB.Frame Frame11 
            Caption         =   "Period for Bill Payment"
            Height          =   2055
            Left            =   3720
            TabIndex        =   54
            Top             =   60
            Width           =   2295
            Begin VB.OptionButton optBillingCycle 
               Caption         =   "Days:"
               Height          =   285
               Index           =   0
               Left            =   300
               TabIndex        =   60
               Tag             =   "d"
               Top             =   360
               Value           =   -1  'True
               Width           =   1005
            End
            Begin VB.TextBox txtBillingCycle 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   0
               Left            =   1440
               MaxLength       =   5
               TabIndex        =   59
               Text            =   "14"
               Top             =   360
               Width           =   675
            End
            Begin VB.OptionButton optBillingCycle 
               Caption         =   "Months:"
               Height          =   255
               Index           =   1
               Left            =   300
               TabIndex        =   58
               Tag             =   "m"
               Top             =   690
               Width           =   915
            End
            Begin VB.TextBox txtBillingCycle 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Index           =   1
               Left            =   1440
               MaxLength       =   5
               TabIndex        =   57
               Text            =   "1"
               Top             =   690
               Width           =   675
            End
            Begin VB.OptionButton optBillingCycle 
               Caption         =   "Hours:"
               Height          =   285
               Index           =   2
               Left            =   300
               TabIndex        =   56
               Tag             =   "h"
               Top             =   990
               Width           =   1005
            End
            Begin VB.TextBox txtBillingCycle 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Index           =   2
               Left            =   1440
               MaxLength       =   5
               TabIndex        =   55
               Text            =   "96"
               Top             =   990
               Width           =   675
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Edits"
            Height          =   2085
            Left            =   6090
            TabIndex        =   52
            Top             =   30
            Width           =   3225
            Begin VB.ListBox lstEdits 
               Height          =   1740
               Left            =   150
               TabIndex        =   53
               Top             =   240
               Width           =   2895
            End
         End
         Begin VB.Frame Frame13 
            Height          =   2055
            Left            =   120
            TabIndex        =   50
            Top             =   90
            Width           =   3525
            Begin VB.CheckBox chkCancelled 
               Caption         =   "Account &Cancelled"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   180
               TabIndex        =   51
               Top             =   240
               Width           =   3225
            End
         End
      End
      Begin VB.PictureBox picHeader 
         BorderStyle     =   0  'None
         Height          =   2265
         Index           =   4
         Left            =   180
         ScaleHeight     =   2265
         ScaleWidth      =   9405
         TabIndex        =   43
         Top             =   420
         Width           =   9405
         Begin VB.Frame Frame4 
            Caption         =   "Downloaded This Period"
            ForeColor       =   &H00C00000&
            Height          =   795
            Index           =   0
            Left            =   90
            TabIndex        =   47
            Top             =   1410
            Width           =   4485
            Begin VB.TextBox txtData 
               Alignment       =   2  'Center
               Height          =   360
               Index           =   0
               Left            =   180
               Locked          =   -1  'True
               TabIndex        =   48
               Text            =   "0 bytes"
               Top             =   270
               Width           =   4185
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Uploaded This Period"
            ForeColor       =   &H00C000C0&
            Height          =   795
            Index           =   1
            Left            =   4650
            TabIndex        =   45
            Top             =   1410
            Width           =   4695
            Begin VB.TextBox txtData 
               Alignment       =   2  'Center
               Height          =   360
               Index           =   2
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   46
               Text            =   "0 bytes"
               Top             =   270
               Width           =   4425
            End
         End
         Begin VB.PictureBox picDatawave 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            Height          =   1335
            Left            =   990
            ScaleHeight     =   85
            ScaleMode       =   0  'User
            ScaleWidth      =   600.218
            TabIndex        =   44
            Tag             =   "Data wave graph mesuring the history of data usage over time of the account creatation"
            Top             =   60
            Width           =   8325
            Begin VB.Line Line4 
               BorderColor     =   &H00FF00FF&
               BorderStyle     =   3  'Dot
               DrawMode        =   4  'Mask Not Pen
               X1              =   6.536
               X2              =   590.414
               Y1              =   40
               Y2              =   40
            End
         End
         Begin VB.Image Image3 
            Height          =   630
            Left            =   150
            Picture         =   "frmCustomerRec_refrehs.frx":63C2
            Stretch         =   -1  'True
            Top             =   90
            Width           =   630
         End
      End
      Begin VB.PictureBox picHeader 
         BorderStyle     =   0  'None
         Height          =   2235
         Index           =   5
         Left            =   90
         ScaleHeight     =   2235
         ScaleWidth      =   9405
         TabIndex        =   25
         Top             =   450
         Width           =   9405
         Begin VB.Frame Frame16 
            Caption         =   "ViSP Node that owns Account"
            Height          =   855
            Index           =   0
            Left            =   90
            TabIndex        =   28
            Top             =   60
            Width           =   9225
            Begin VB.ComboBox cmbVirtualID 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Left            =   150
               Style           =   2  'Dropdown List
               TabIndex        =   29
               Top             =   240
               Width           =   8955
            End
         End
         Begin VB.Frame Frame16 
            Caption         =   "Sysop that created, controls and get commission on the account"
            Height          =   855
            Index           =   1
            Left            =   90
            TabIndex        =   26
            Top             =   990
            Width           =   9255
            Begin VB.ComboBox cmbSysopID 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Left            =   150
               Style           =   2  'Dropdown List
               TabIndex        =   27
               Top             =   240
               Width           =   8985
            End
         End
      End
      Begin VB.PictureBox picHeader 
         BorderStyle     =   0  'None
         Height          =   2235
         Index           =   6
         Left            =   90
         ScaleHeight     =   2235
         ScaleWidth      =   9405
         TabIndex        =   23
         Top             =   390
         Width           =   9405
         Begin MSComctlLib.ListView lvEditLog 
            Height          =   2115
            Left            =   90
            TabIndex        =   24
            Top             =   60
            Width           =   9225
            _ExtentX        =   16272
            _ExtentY        =   3731
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Date Edit was Made"
               Object.Width           =   4939
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Username"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Edit text"
               Object.Width           =   17639
            EndProperty
         End
      End
      Begin MSComctlLib.TabStrip tsHeader 
         Height          =   2625
         Left            =   0
         TabIndex        =   42
         Top             =   90
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   4630
         MultiRow        =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   7
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Account &Information"
               Object.ToolTipText     =   "Here is where the account information is found."
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Record &Details"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "&Domain"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "&Billing Details"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Data &Usage"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "&Ownership"
               Object.ToolTipText     =   "Here is where you can adjust the ownership of the customer, ie. Sysop and ViSP"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Edit Logs"
               Object.ToolTipText     =   "Here is the edit logs"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.PictureBox picHeader 
         BorderStyle     =   0  'None
         Height          =   2235
         Index           =   2
         Left            =   30
         ScaleHeight     =   2235
         ScaleWidth      =   9405
         TabIndex        =   61
         Top             =   30
         Width           =   9405
         Begin MSComctlLib.ListView lvDomains 
            Height          =   2055
            Left            =   1110
            TabIndex        =   62
            Top             =   90
            Width           =   8205
            _ExtentX        =   14473
            _ExtentY        =   3625
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   14
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Domain"
               Object.Width           =   6438
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Admin's Email Address"
               Object.Width           =   5292
            EndProperty
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Domain's:"
            Height          =   240
            Left            =   180
            TabIndex        =   63
            Top             =   120
            Width           =   705
         End
      End
      Begin VB.PictureBox picHeader 
         BorderStyle     =   0  'None
         Height          =   2175
         Index           =   1
         Left            =   90
         ScaleHeight     =   2175
         ScaleWidth      =   9345
         TabIndex        =   34
         Top             =   150
         Width           =   9345
         Begin VB.Frame Frame6 
            Caption         =   "Record Flags"
            Height          =   2025
            Left            =   90
            TabIndex        =   37
            Top             =   90
            Width           =   5115
            Begin VB.Frame Frame10 
               Caption         =   "Flag A (Primary)"
               Height          =   795
               Left            =   150
               TabIndex        =   40
               Top             =   210
               Width           =   4875
               Begin VB.ComboBox cmbFlagA 
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
                  Left            =   150
                  Style           =   2  'Dropdown List
                  TabIndex        =   41
                  Top             =   270
                  Width           =   4605
               End
            End
            Begin VB.Frame Frame12 
               Caption         =   "Flag B (Secondary)"
               Height          =   795
               Left            =   150
               TabIndex        =   38
               Top             =   1020
               Width           =   4875
               Begin VB.ComboBox cmbFlagB 
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
                  Left            =   150
                  Style           =   2  'Dropdown List
                  TabIndex        =   39
                  Top             =   270
                  Width           =   4575
               End
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "&Classification"
            Height          =   2025
            Left            =   5280
            TabIndex        =   35
            Top             =   90
            Width           =   3975
            Begin MSComctlLib.ListView lvClassification 
               Height          =   1605
               Left            =   120
               TabIndex        =   36
               Top             =   270
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   2831
               View            =   3
               LabelEdit       =   1
               Sorted          =   -1  'True
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
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   1
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Class"
                  Object.Width           =   4674
               EndProperty
            End
         End
      End
      Begin VB.PictureBox picTSContainer 
         BorderStyle     =   0  'None
         Height          =   1755
         Index           =   3
         Left            =   60
         ScaleHeight     =   1755
         ScaleWidth      =   2955
         TabIndex        =   64
         Top             =   2790
         Visible         =   0   'False
         Width           =   2955
         Begin VB.CommandButton cmdChangePlan 
            Caption         =   "&Change Plan"
            Enabled         =   0   'False
            Height          =   285
            Left            =   3510
            TabIndex        =   66
            Top             =   4050
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.CommandButton cmdAddService 
            Caption         =   "&Add Subscription or Product"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   60
            TabIndex        =   65
            Top             =   3990
            Width           =   2985
         End
         Begin MSComctlLib.ListView lvPlans 
            Height          =   3795
            Left            =   90
            TabIndex        =   67
            Top             =   60
            Width           =   9225
            _ExtentX        =   16272
            _ExtentY        =   6694
            View            =   3
            Arrange         =   1
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            PictureAlignment=   1
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
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
            Picture         =   "frmCustomerRec_refrehs.frx":6804
         End
      End
      Begin VB.PictureBox picTSContainer 
         BorderStyle     =   0  'None
         Height          =   3075
         Index           =   5
         Left            =   3390
         ScaleHeight     =   3075
         ScaleWidth      =   4905
         TabIndex        =   103
         Top             =   3210
         Width           =   4905
         Begin SHDocVwCtl.WebBrowser wwwFileDB 
            Height          =   4365
            Left            =   60
            TabIndex        =   104
            Top             =   60
            Width           =   9345
            ExtentX         =   16484
            ExtentY         =   7699
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "http:///"
         End
      End
      Begin VB.PictureBox picTSContainer 
         BorderStyle     =   0  'None
         Height          =   795
         Index           =   4
         Left            =   2850
         ScaleHeight     =   795
         ScaleWidth      =   1755
         TabIndex        =   90
         Top             =   3810
         Visible         =   0   'False
         Width           =   1755
         Begin VB.CommandButton cmdBill 
            Caption         =   "Add &Tranaction"
            Height          =   285
            Left            =   90
            TabIndex        =   93
            Top             =   3990
            Width           =   1605
         End
         Begin VB.CommandButton Command1 
            Height          =   285
            Left            =   1680
            TabIndex        =   92
            Top             =   3990
            Visible         =   0   'False
            Width           =   1605
         End
         Begin MSComctlLib.ListView lvTransactions 
            Height          =   3855
            Left            =   60
            TabIndex        =   91
            Top             =   60
            Width           =   9225
            _ExtentX        =   16272
            _ExtentY        =   6800
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
            BackColor       =   -2147483643
            BorderStyle     =   1
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
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Description"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   1
               Text            =   "Total Due"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "GST"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Payment Due"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Amount Paid"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "All Paid"
               Object.Width           =   1411
            EndProperty
            Picture         =   "frmCustomerRec_refrehs.frx":3EC56
         End
      End
      Begin MSComctlLib.TabStrip tsDetails 
         Height          =   4845
         Left            =   0
         TabIndex        =   95
         Top             =   2760
         Width           =   9525
         _ExtentX        =   16801
         _ExtentY        =   8546
         Placement       =   1
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   6
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Address' && DOB"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Phone No's && &Email"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "&Referals and Dates"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Subscriptions And &Products"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "&Transactions"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "&File Database"
               Object.ToolTipText     =   "This is where you can store any email or files relating to this customer."
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.PictureBox picTSContainer 
         BorderStyle     =   0  'None
         Height          =   4395
         Index           =   2
         Left            =   120
         ScaleHeight     =   4395
         ScaleWidth      =   9375
         TabIndex        =   71
         Top             =   2820
         Visible         =   0   'False
         Width           =   9375
         Begin VB.Frame Frame3 
            Caption         =   "Refered By"
            Height          =   4335
            Left            =   3690
            TabIndex        =   77
            Top             =   0
            Width           =   5505
            Begin VB.Frame Frame5 
               Caption         =   "Search Options"
               Height          =   1395
               Left            =   120
               TabIndex        =   78
               Top             =   210
               Width           =   5265
               Begin VB.TextBox txtSearchName 
                  Height          =   285
                  Left            =   180
                  TabIndex        =   83
                  Top             =   300
                  Width           =   4905
               End
               Begin VB.CheckBox chkSearchAccountNames 
                  Caption         =   "Search Account Names"
                  Height          =   255
                  Left            =   180
                  TabIndex        =   82
                  Top             =   660
                  Value           =   1  'Checked
                  Width           =   2295
               End
               Begin VB.CheckBox chkSearchContacts 
                  Caption         =   "Search All Contact Names"
                  Height          =   255
                  Left            =   180
                  TabIndex        =   81
                  Top             =   930
                  Width           =   2265
               End
               Begin VB.CheckBox chkSearchUsername 
                  Caption         =   "Search Username"
                  Height          =   240
                  Left            =   2520
                  TabIndex        =   79
                  Top             =   660
                  Width           =   1875
               End
               Begin MSComctlLib.ProgressBar pb1 
                  Height          =   105
                  Left            =   30
                  TabIndex        =   80
                  Top             =   1200
                  Width           =   5205
                  _ExtentX        =   9181
                  _ExtentY        =   185
                  _Version        =   393216
                  Appearance      =   0
                  Max             =   6
               End
            End
            Begin MSComctlLib.ListView lvReferedBy 
               Height          =   2505
               Left            =   120
               TabIndex        =   84
               Top             =   1680
               Width           =   5265
               _ExtentX        =   9287
               _ExtentY        =   4419
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               Checkboxes      =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   3
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Account Name"
                  Object.Width           =   3528
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Contact Name"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Short Note"
                  Object.Width           =   7056
               EndProperty
            End
         End
         Begin VB.Frame Frame15 
            Caption         =   "Next Billing Date"
            Height          =   3585
            Left            =   120
            TabIndex        =   74
            Top             =   0
            Width           =   3405
            Begin MSComCtl2.MonthView mvBilling 
               Height          =   2820
               Left            =   1080
               TabIndex        =   76
               Top             =   1350
               Width           =   3120
               _ExtentX        =   5503
               _ExtentY        =   4974
               _Version        =   393216
               ForeColor       =   -2147483630
               BackColor       =   -2147483633
               Appearance      =   1
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               StartOfWeek     =   66912258
               CurrentDate     =   37705
            End
            Begin MSComCtl2.DTPicker dtpBilling 
               Height          =   315
               Left            =   150
               TabIndex        =   75
               Top             =   3150
               Width           =   3135
               _ExtentX        =   5530
               _ExtentY        =   556
               _Version        =   393216
               Format          =   66912258
               CurrentDate     =   37753
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Date of Birth/Registration of Organisation"
            Height          =   705
            Left            =   120
            TabIndex        =   72
            Top             =   3630
            Width           =   3405
            Begin MSComCtl2.DTPicker dtpDOB 
               Height          =   345
               Left            =   120
               TabIndex        =   73
               Top             =   210
               Width           =   3195
               _ExtentX        =   5636
               _ExtentY        =   609
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   66912256
               CurrentDate     =   37720
            End
         End
      End
      Begin VB.PictureBox picTSContainer 
         BorderStyle     =   0  'None
         Height          =   4395
         Index           =   0
         Left            =   90
         ScaleHeight     =   4395
         ScaleWidth      =   9435
         TabIndex        =   68
         Top             =   2820
         Visible         =   0   'False
         Width           =   9435
         Begin VB.CommandButton cmdAddAddress 
            Caption         =   "&Add Address"
            Height          =   375
            Left            =   60
            TabIndex        =   69
            Top             =   3990
            Width           =   1995
         End
         Begin MSComctlLib.ListView lvAddresses 
            Height          =   3795
            Left            =   60
            TabIndex        =   70
            Top             =   120
            Width           =   9075
            _ExtentX        =   16007
            _ExtentY        =   6694
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
               Name            =   "Arial"
               Size            =   8.25
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
            Picture         =   "frmCustomerRec_refrehs.frx":588A4
         End
      End
   End
   Begin VB.PictureBox picBut 
      BorderStyle     =   0  'None
      Height          =   2235
      Index           =   1
      Left            =   120
      ScaleHeight     =   2235
      ScaleWidth      =   2295
      TabIndex        =   97
      Tag             =   "<b>audithistory<b/>"
      Top             =   2910
      Width           =   2295
      Begin MSDataGridLib.DataGrid dgAudit 
         Bindings        =   "frmCustomerRec_refrehs.frx":8834E
         Height          =   5325
         Left            =   150
         TabIndex        =   100
         Top             =   1830
         Width           =   9345
         _ExtentX        =   16484
         _ExtentY        =   9393
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Audit History for Client or Company"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3081
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3081
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame14 
         Caption         =   "Statistics"
         Height          =   1695
         Left            =   150
         TabIndex        =   98
         Top             =   90
         Width           =   9375
         Begin MSComctlLib.ListView lvAuditStats 
            Height          =   1335
            Left            =   120
            TabIndex        =   99
            Top             =   240
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   2355
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   0
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Description"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Result"
               Object.Width           =   15875
            EndProperty
         End
      End
      Begin MSAdodcLib.Adodc adoAudit 
         Height          =   330
         Left            =   150
         Top             =   7230
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
   Begin VB.PictureBox picBut 
      BorderStyle     =   0  'None
      Height          =   7725
      Index           =   3
      Left            =   90
      ScaleHeight     =   7725
      ScaleWidth      =   9525
      TabIndex        =   133
      Tag             =   "<b>FileDB<b/>"
      Top             =   1410
      Width           =   9525
   End
   Begin VB.Label lblCustCreated 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Created: 00/00/0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   210
      Left            =   6570
      TabIndex        =   19
      Top             =   870
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.Label lblCustomerCreated 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Created: 00/00/0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   240
      Left            =   6570
      TabIndex        =   18
      Top             =   1170
      Width           =   3105
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Record"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   345
      Left            =   60
      TabIndex        =   17
      Top             =   810
      Width           =   2475
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      BorderWidth     =   3
      X1              =   2
      X2              =   764
      Y1              =   76
      Y2              =   76
   End
End
Attribute VB_Name = "frmCustomerRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const SubPath = "Clients"

Public UNCPath As String

Public bEditMade As Boolean
Public FormLoad As Date
Public osub As New clsSubscriber

Public FormState As enumFormState

Dim Saved As Boolean

Dim bsmtpCreation As Boolean
Public SESSION As String
Dim sTMPShortNote As String



Private Sub chkCancelled_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "chkCancelled_Click"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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



Private Sub cmbFlagA_GotFocus()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmbFlagA_GotFocus"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    
    If bPlay = False And picHeader(1).Visible = True Then
            
        L = GetCursorPos(oXY)
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y

        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "This is the primary flag on the account, it controls access to the service as Not Set it allow full use of Radius access and account privileges."
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

Private Sub cmdAddAddress_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdAddAddress_Click"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
        
        With osub.colSnailMail.Add("NEW" & SESSION & osub.colSnailMail.Count + 1, osub.fRecID, ffrmSnailMail.FlagID, _
                                sysnow, ffrmSnailMail.sContactName, ffrmSnailMail.sStreetLine1, ffrmSnailMail.sStreetLine2, _
                                ffrmSnailMail.sCountry, ffrmSnailMail.sState, ffrmSnailMail.sPostcode, ffrmSnailMail.sSuburb, _
                                0, True, "NEW" & SESSION & osub.colSnailMail.Count + 1)
        
            Dim itmX As ListItem
            Set itmX = lvAddresses.ListItems.Add(1, .Key, .ContactName)
            itmX.SubItems(1) = .Street1
            itmX.SubItems(2) = .Street2
            itmX.SubItems(3) = .Suburb
            itmX.SubItems(4) = .State
            itmX.SubItems(5) = .PostCode
            itmX.SubItems(6) = .Country
            itmX.Checked = .Checked
            bEditMade = True
        End With
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
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
        ffrmEmail.sContactName = txtAccountName.Text
        ffrmEmail.sEmailAddress = "@" & IIf(txtRealm.Text <> "", txtRealm, "ep.net.au")
    End If
    ffrmEmail.Show 1
    
    If ffrmEmail.iCloseState = frmCloseSave Then
        With osub.col_subEmails.Add("NEW" & SESSION & osub.col_subEmails.Count + 1, 0, osub.fRecID, ffrmEmail.FlagID, sysnow, ffrmEmail.sEmailAddress, ffrmEmail.sContactName, False, True, "NEW" & SESSION & osub.col_subEmails.Count + 1)
        
        
            Dim itmX As ListItem
            Set itmX = lvEmail.ListItems.Add(, .Key, .ContactName)
            itmX.SubItems(1) = .EmailAddress
            itmX.Checked = .Checked
            bEditMade = True
        End With
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
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
        ffrmPhoneNo.sContactName = txtAccountName.Text
    End If
    
    ffrmPhoneNo.Show 1
    
    If ffrmPhoneNo.iCloseState = frmCloseSave Then
        
        With osub.col_subPhoneNo.Add("NEW" & SESSION & osub.col_subPhoneNo.Count + 1, 0, osub.fRecID, ffrmPhoneNo.sFlagID, sysnow, ffrmPhoneNo.sPhonenumber, ffrmPhoneNo.sExtension, ffrmPhoneNo.sContactName, False, True, ffrmPhoneNo.sNote, "NEW" & SESSION & osub.col_subPhoneNo.Count + 1)
        
            Dim itmX As ListItem
            Set itmX = lvPhone.ListItems.Add(1, .Key, .ContactName)
            itmX.SubItems(1) = .PhoneNumber
            itmX.SubItems(2) = .Extension
            itmX.SubItems(3) = .ShortNote
            itmX.Checked = .Checked
            bEditMade = True
        End With
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



Private Sub cmdAddService_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdAddService_Click"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    If bDebug = True Then On Error GoTo 0 Else On Error Resume Next
    
    Dim fAddPlan As New frmAddPlan
    Dim frmPayments As New frmAccPayment
    Dim fDomain As New frmDNS
    Dim fRadius As New frmRadiusAccount
    
    Dim fDSL As New frmBroadband
    Dim fAlias As New frmAlias
    
    
    Dim rsTMP As adodb.Recordset
    
    fAddPlan.Show 1
    
    Dim noFDSL As Boolean
    
    Dim ShippingIDX As Long
    Dim SubRecID As Long
    Dim acciRecID As Long
    Dim ContractID As Long
    Dim ContractExpiry As Date
    Dim JoiningFee As Single
    Dim PerHour As Single
    Dim PeriodFee As Single
    Dim PerMB As Single
    Dim Activation As Date
    Dim NextCycle As Date
    Dim ServiceID As Long
    Dim ptRecID As Long
    Dim MBQuota As Integer
    Dim Username As String
    Dim Password As String
    Dim BaseURL As String
    Dim DynamicField1 As String
    Dim DynamicField2 As String
    Dim DynamicField3 As String
    Dim ContactName As String
    Dim DynamicField4 As String
    Dim DynamicField5 As String
        
    Dim rsRadius As adodb.Recordset
    Dim rsCheck As adodb.Recordset
    Dim rsinvoiceout As adodb.Recordset
    Dim rsAccInfo As adodb.Recordset
    Dim rsDomains As adodb.Recordset
    Dim fPOP3 As frmPOP3Account
    Dim fFTP As frmFTPAccount
    Dim bx As Byte
    Dim sPassword As String
    Dim l5meoDMT As Variant
    Dim FlagID As Integer
    Dim bSkip As Boolean
    
    Set fFTP = New frmFTPAccount
    Set fPOP3 = New frmPOP3Account
    Set rsinvoiceout = New adodb.Recordset
    Set rsAccInfo = New adodb.Recordset
    
    Dim InvRecID  As Variant
    Dim bResult As Boolean
    Dim lSubRecID As Long
    Dim invout As invoiceout
    Dim itmy As ListItem
    Dim oSvCount As Long
    
    'Dim acci_services As type_acci_services
    
    'FlagID = 8.43874378437844E+22
    
    If fAddPlan.ptRecID <> 0 Then
                
        Dim frmShip As New frmShipping
        Dim tmpSession As String
        tmpSession = GetSessionChar(tmpSession, Me.hWnd, 14)
        
        ' gets shipping information
        Set frmShip.LinkdForm = Me
        frmShip.picTSContainer(0).Visible = True
        frmShip.Show 1
        ShippingIDX = frmShip.ShippingID
        
        For oSvCount = 1 To oService.Count
            
            On Error GoTo 0
            bSkip = False
            
            
            acciRecID = osub.fRecID
            ContractID = oService.ContractID
            ContractExpiry = DateAdd(IIf(oService.IntervalType = "", "d", oService.IntervalType), oService.IntervalLength, sysnow)
            JoiningFee = oService(oSvCount).JoiningFee
            PerHour = oService(oSvCount).PerHour
            PeriodFee = oService(oSvCount).PeriodFee
            PerMB = oService(oSvCount).PerMBBlock
            Activation = oService.Activated
            'NextCycle = DateAdd(oService(oSvCount).cycType, oService(oSvCount).cycInterval, Sysnow)
            ServiceID = oService(oSvCount).ServiceID
            ptRecID = oService(oSvCount).ptRecID
            MBQuota = oService(oSvCount).MBQuota
            
            radiusset = False
            
            If oService(oSvCount).ListedOnRadius = True Then
                Set fRadius = New frmRadiusAccount
                fRadius.p_ContactName = osub.fAccountName
                fRadius.p_Password = osub.fgPassword
                fRadius.p_Username = osub.fgUsername
                fRadius.p_SessionTimeOut = oService(oSvCount).SessionTimeout
                fRadius.p_IdleTimeout = oService(oSvCount).IdleTimeout
                fRadius.p_Sessions = oService(oSvCount).SessionsAllowed
                
                If oService(oSvCount).svrCode = "ADSL" Or oService(oSvCount).svrCode = "SHDSL" Then
                    If noFDSL = False Then
                        noFDSL = True
                        fDSL.Show 1
                    End If
                    
                    fRadius.p_Username = fDSL.sAreaCode & "" & fDSL.sPhoneNum
                    fRadius.cmdUsername.Enabled = False
                    fRadius.txtUID.Locked = True
                End If
                
                Screen.MousePointer = vbDefault
                fRadius.p_ContactName = osub.fAccountName
                fRadius.Show 1
                Screen.MousePointer = vbHourglass
                
                If fRadius.iCloseState = frmCloseSave Then
                    
                    Call osub.col_subRadius.Add("NEW" & tmpSession & osub.col_subRadius.Count + 1, 0, fRadius.p_Username, fRadius.p_Password, fRadius.p_Sessions, _
                                        fRadius.p_AutoFlag, fRadius.p_Activate, fRadius.p_Deactivate, fRadius.p_SessionTimeOut, fRadius.p_IdleTimeout, _
                                        osub.fRecID, oService(oSvCount).svrCode, "", IIf(DateDiff("s", oService.Activated, sysnow) > 0, True, False), sysnow, sysnow, sysnow, 0, 0, 0, "", fAddPlan.ptRecID, Login.lVirtualID, _
                                        "", "", sysnow, "NEW" & tmpSession & osub.col_subRadius.Count + 1)
                    
                    Username = fRadius.p_Username
                    Password = fRadius.p_Password
                    ContactName = fRadius.p_ContactName
                    radiusset = True
                    'MySQL.Execute directConn, "UPDATE radiusaccounts SET Password=AES_ENCRYPT('" & fRadius.p_Password & "','" & odb.colSalts.ReturnSalt("md5Password") & "') Where RecID = " & RadiusID
                    'mysql.execute directConn,  "UPDATE radiusaccounts SET Password=MD5('" & fradius.p_Password & "') Where RecID = " & rsSave3!RadiusID
                                        
                Else
                    Call osub.ClearSession(tmpSession, Me.lvPlans, Me.lvDomains, Me.lvTransactions)
                    Exit Sub
                End If
            End If
                   
            FlagID = 0
            Select Case oService(oSvCount).svrCode
            Case "ALIAS"
                
                If lvDomains.ListItems.Count = 0 Then
                    
                    MsgBox "The customer must have there own domain to have email aliases!", vbCritical, "No domain Found"
                    bSkip = True
                    
                Else
                    
                    Screen.MousePointer = vbDefault
                    'fAlias.sContactName = txtAccountName.Text
                    Set fAlias.oCNT = Me
                    fAlias.Show 1
                        Screen.MousePointer = vbHourglass
                    If fAlias.iCloseState = frmCloseCancel Then
                        bSkip = True
                    Else
                        Call osub.ClearSession(fAlias.SESSION, Me.lvPlans, Me.lvDomains, Me.lvTransactions)
                        Call osub.ClearSession(tmpSession, Me.lvPlans, Me.lvDomains, Me.lvTransactions)
                        Exit Sub
                    End If
                        
                End If
                
            Case "ADSL", "SHDSL"
                    
                Screen.MousePointer = vbDefault
                If noFDSL = False Then fDSL.Show 1
                With osub.col_subPhoneNo.Add("NEW" & tmpSession & osub.col_subPhoneNo.Count + 1, 0, osub.fRecID, 0, sysnow, _
                    fDSL.sAreaCode & " " & fDSL.sPhoneNum, "", fDSL.sName, 0, False, "DSL CONNECTION - " & fDSL.sName, "NEW" & tmpSession & osub.col_subPhoneNo.Count + 1)
                    
                    Set itmy = lvPhone.ListItems.Add(, , .ContactName)
                    itmy.SubItems(1) = .PhoneNumber
                End With
                
                Call osub.col_subDSLLink.Add("NEW" & tmpSession & osub.col_subDSLLink.Count + 1, fDSL.sName, fDSL.sAreaCode, fDSL.sPhoneNum, fDSL.sEmail, _
                     osub.fRecID, 0, True, fDSL.UnitNumber, fDSL.StreetNo, fDSL.StreetName, fDSL.StreetType, fDSL.Suburb, fDSL.Country, fDSL.PostCode, _
                     fDSL.State, sysnow, fDSL.Churn, osub.col_subRadius.Count, 0, 0, "NEW" & tmpSession & osub.col_subDSLLink.Count + 1)
                     
                Screen.MousePointer = vbHourglass
                FlagID = 3
                
            Case "FTP"
            
DoFTPAgain2:
    
                fFTP.sContactName = txtAccountName.Text
                fFTP.sUsername = osub.fgUsername
                fFTP.sPassword = osub.fgPassword
                fFTP.sBaseDIR = "/home/" & Login.sVISPDomain & "/" & fRadius.p_Username
                Screen.MousePointer = vbDefault
                fFTP.Show 1
                Screen.MousePointer = vbHourglass
                If fFTP.sUsername <> "Not Set" Then
                    bResult = MySQL.OpenTable(directConn, rsCheck, , "select * from acci_services Where ServiceID = " & oService(oSvCount).ServiceID & " and Username = '" & MySQL.ESC(fFTP.sUsername) & "' Limit 1")
                    If rsCheck.RecordCount > 0 Then
                        MsgBox "Username for this service already exist on schemer.", vbCritical, "Username Exists"
                        GoTo DoFTPAgain2
                    End If
                End If
                
                If fFTP.iCloseState = frmCloseSave Then
                    Username = fFTP.sUsername
                    Password = fFTP.sPassword
                    If fRadius.p_Username <> Username Then
                        fFTP.sBaseDIR = "/home/" & Login.sVISPDomain & "/" & Username
                    End If
                    If fFTP.sBaseDIR = "" Then fFTP.sBaseDIR = "/home/" & Login.sVISPDomain & "/" & Username
    
                    BaseURL = fFTP.sBaseDIR
                    DynamicField1 = fFTP.sSessions
                    DynamicField2 = fFTP.byteBandwidth
                    DynamicField3 = fFTP.byteBWUpload
                    ContactName = fFTP.sContactName
                    DynamicField4 = "21"
                    DynamicField5 = "21"
                Else
                    Call osub.ClearSession(tmpSession, Me.lvPlans, Me.lvDomains, Me.lvTransactions)
                    Exit Sub
                End If
                    
               ' rsSave3.Update
            Case "POP3"
                
DoPOP3Again2:
                Set fPOP3.oCNT = osub
                Set fPOP3.frm = Me
                fPOP3.PeriodFee = oService(oSvCount).PeriodFee
                fPOP3.PerHour = oService(oSvCount).PerHour
                fPOP3.PerMB = oService(oSvCount).PerMBBlock
                fPOP3.JoiningFee = oService(oSvCount).JoiningFee
                fPOP3.DefShippingID = frmShip.ShippingID
                fPOP3.Description = oService(oSvCount).Description
                fPOP3.ptRecID = oService(oSvCount).ptRecID
                fPOP3.ServiceID = oService(oSvCount).ServiceID
                fPOP3.NumAdd = IIf(oService(oSvCount).NumOf = 0, 1, oService(oSvCount).NumOf)
                
                Screen.MousePointer = vbDefault
                fPOP3.Show 1
                If fPOP3.iCloseState <> frmCloseSave Then
                    
                    Call osub.ClearSession(fPOP3.SESSION, Me.lvPlans, Me.lvDomains, Me.lvTransactions)
                    Call osub.ClearSession(tmpSession, Me.lvPlans, Me.lvDomains, Me.lvTransactions)
                    Exit Sub

                End If
                Screen.MousePointer = vbHourglass
                
                bSkip = True
            Case "DOMAIN"
    
doDomainAgain2:
                domainset = False
                Screen.MousePointer = vbDefault
                Set fDomain.oCNT = osub
                fDomain.Show 1
                If fDomain.iCloseState = frmCloseCancel Then
                    Call osub.ClearSession(fDomain.SESSION, Me.lvPlans, Me.lvDomains, Me.lvTransactions, Me.lvDomains)
                    Call osub.ClearSession(tmpSession, Me.lvPlans, Me.lvDomains, Me.lvTransactions)
                    Exit Sub
                End If
                
                Screen.MousePointer = vbHourglass
                
'                If fDomain.iCloseState = 2 Then
'
'                    On Error Resume Next
'                    Do
'                        Err.Clear
'                        DomainID = MySQL.GetTMPRecID("domainlist", directConn)
'                        directConn.Execute "insert into domainlist (TechPass,TechName, acci_RecID,AdminEmail,Domain,RecID,SysopID,VirtualID) " + _
'                                                                            "VALUES (AES_ENCRYPT('" & MySQL.ESC(fDomain.sAdminEmail) & "','" & _
'                                                                            odb.colSalts.ReturnSalt(PWSalt) & "'),'" & MySQL.ESC(fDomain.sTechName) & "','" & acciRecID & "','" & _
'                                                                            MySQL.ESC(fDomain.sAdminEmail) & "','" & MySQL.ESC(fDomain.sDomain) & "','" & DomainID & "','" & _
'                                                                            Login.lSysopID & "','" & Login.lVirtualID & "')"
'                    Loop Until Err.Number = 0
'
'
'                    MySQL.Execute directConn, "update domainlist set Checked = '-1' where RecID = '" & DomainID
'
'                    If oIP.colDNSDocs.Count > 0 Then
'                        Dim Xnt As Integer
'                        For Xnt = oIP.colDNSDocs.Count To 1 Step -1
'                            MySQL.Execute directConn, "INSERT INTO domaindocs (DomainID, DocType, DocText, Icon, Description, ItemText) " + _
'                                                   "VALUES ('" & DomainID & "','" & oIP.colDNSDocs(Xnt).DocType & "','" + _
'                                                   MySQL.ESC(oIP.colDNSDocs(Xnt).DocText) & "','" & oIP.colDNSDocs(Xnt).bIcon & "','" + _
'                                                   MySQL.ESC(oIP.colDNSDocs(Xnt).Description) & "','" & MySQL.ESC(oIP.colDNSDocs(Xnt).ItemText) & "')"
'                            oIP.colDNSDocs.Remove Xnt
'                        Next
'                    End If
'
'                    Set itmX = lvDomains.ListItems.Add(, "r" & DomainID, fDomain.sDomain)
'                    itmX.SubItems(1) = fDomain.sAdminEmail
'
'                    'domainlist 11
'
'                    If fDomain.sKey <> "" Then MySQL.Execute directConn, "Update domainlist Set vKey = AES_ENCRYPT('" & fDomain.sKey & "','" & odb.colSalts.ReturnSalt("md5Password") & "') Where RecID = " & DomainID
'
'                Else
'                    oService(oSvCount).BillNow = False
'                    bSkip = True
'                    BaseURL = "ep.net.au"
'                    DynamicField1 = "support@ep.net.au"
'                End If
'
'
            End Select
            
            'bResult = MySQL.OpenTable(directConn, rsSave3, , "select * from acci_services Limit 1")
            
            On Error Resume Next
            If bSkip = False Then
                
            If ContactName = "" Then ContactName = txtAccountName.Text
                
                
'                Do
'                    Err.Clear
'                    oSub.fRecID_acci_services = MySQL.GetTMPRecID("acci_services", directConn)
'
'                    If oSvCount = 1 Then lSubRecID = oSub.fRecID_acci_services
'
'
'                    directConn.Execute "Insert into acci_services (RecID, ContractID, PreviousCycle, ContractExpiry, DefaultShippingID, AgencyID, RadiusID, ServiceID, ptRecID, DomainID, VirtualID, AccI_RecID, SysopID, NextCycle, Checked, ContactName, Username, Password, BaseURL, DynamicField1, DynamicField2, DynamicField3, DynamicField4, DynamicField5, SubRecID, MBQuota, Activation, PeriodFee, JoiningFee, PerMB, PerHour) " + _
'                                    "VALUES (" & oSub.fRecID_acci_services & ",'" & ContractID & "','" & Format(Activation, "yyyy-mm-dd ttttt") & "','" & Format(ContractExpiry, "yyyy-mm-dd ttttt") & "','" & frmShip.ShippingID & "'," & Login.lAgencyID & "," & RadiusID & "," & ServiceID & "," & ptRecID & "," & DomainID & "," & Login.lVirtualID & ",'" & acciRecID & "'," & Login.lSysopID & ",'" & "" & "', " & _
'                                    IIf(Checked = True, "-1", "0") & ",'" & MySQL.ESC(ContactName) & "','" & MySQL.ESC(Username) & "',AES_ENCRYPT('" & MySQL.ESC(Password) & "','" & odb.colSalts.ReturnSalt("md5Password") & "'),'" & BaseURL & "','" & DynamicField1 & "','" & DynamicField2 & _
'                                    "','" & DynamicField3 & "','" & DynamicField4 & "','" & DynamicField5 & "'," & lSubRecID & "," & MBQuota & ",'" & Format(Activation, "yyyy-mm-dd Hh:Nn:Ss") & "'," & PeriodFee & "," & JoiningFee & "," & PerMB & "," & PerHour & ")"
'                    If Err.Number <> 0 Then cDebug Err.Description
'
'                Loop Until Err.Number = 0
                
            If SubRecID = 0 Then SubRecID = osub.col_subServices.Count + 1
            
            On Error GoTo 0
            
            If oService.IntervalType = "" Then
                oService.IntervalType = "s"
                oService.IntervalLength = 1
            End If
            With osub.col_subServices.Add("NEW" & tmpSession & osub.col_subServices.Count + 1, 0, oService(oSvCount).ptRecID, oService(oSvCount).ServiceID, ContactName, Username, _
                    Password, Format(oService.Activated, "yyyy-mm-dd ttttt"), BaseURL, IIf(radiusset = True, osub.col_subRadius.Count, 0), sysnow, osub.fRecID, DynamicField1, DynamicField2, _
                    DynamicField3, DynamicField4, DynamicField5, IIf(DateDiff("s", oService.Activated, sysnow) > 0, True, False), Login.lVirtualID, sysnow, IIf(domainset = True, osub.col_Domains.Count, 0), SubRecID, oService(oSvCount).MBQuota, _
                    oService.Activated, oService(oSvCount).PeriodFee, oService(oSvCount).PerHour, oService(oSvCount).PerMBBlock, oService(oSvCount).JoiningFee, _
                    Login.lAgencyID, ShippingIDX, oService.ContractID, DateAdd(oService.IntervalType, oService.IntervalLength, oService.Activated), IIf(DateDiff("s", oService.Activated, sysnow) > 10, True, False), "NEW" & tmpSession & osub.col_subServices.Count + 1)
                
                Screen.MousePointer = vbDefault
                
                Set itmX = lvPlans.ListItems.Add(1, .Key, .ContactName)
                itmX.Tag = .ServiceID
                itmX.Checked = .Checked
                itmX.SubItems(1) = IIf(IsNull(oService(oSvCount).Description), "", oService(oSvCount).Description)
                itmX.SubItems(2) = .Username
                itmX.SubItems(3) = .Password
                itmX.SubItems(4) = .BaseURL
                itmX.SubItems(5) = .DynamicField1
                itmX.SubItems(6) = .DynamicField2
                itmX.SubItems(7) = .DynamicField3
                itmX.SubItems(8) = .DynamicField4
                itmX.SubItems(9) = .DynamicField5
            
            End With
            
'            Dim bRefered As Long
'            Dim oCaption As String
'            oCaption = Me.Caption
'
'            Dim rsCount As ADODB.Recordset
'            Dim rsBonus As ADODB.Recordset
'            Dim rsSpiff As ADODB.Recordset
'
'            If bSkip = False Then
'                Me.Caption = "Searching for Bonuses"
'                Call MySQL.OpenTable(directConn, rsBonus, , "select distinct AwardTo from bonus_awards where ptRecID = " & ptRecID)
'                If rsBonus.RecordCount > 0 Then
'                    While Not rsBonus.EOF And Err.Number = 0
'                        Select Case rsBonus!AwardTo
'                        Case 0 ' To Sysop
'                            Call MySQL.OpenTable(directConn, rsCount, , "select Count(*) as RecCount from acci_services where ptRecID = " & ptRecID & " and SysopID = " & Login.lSysopID)
'                        Case 1 ' TO Site
'                            Call MySQL.OpenTable(directConn, rsCount, , "select Count(*) as RecCount from acci_services where ptRecID = " & ptRecID & " and VirtualID = " & Login.lVirtualID)
'                        Case 2 ' To Agency
'                            Call MySQL.OpenTable(directConn, rsCount, , "select Count(*) as RecCount from acci_services where ptRecID = " & ptRecID & " and AgencyID = " & Login.lAgencyID)
'                        Case 3 ' To Developer
'                            Call MySQL.OpenTable(directConn, rsCount, , "select Count(*) as RecCount from acci_services where ptRecID = " & ptRecID)
'                        End Select
'
'                        If rsCount.RecordCount > 0 Then
'                            Call MySQL.OpenTable(directConn, rsSpiff, , "select bonus_awards.*, bonus_units.FieldName from bonus_awards, bonus_units where bonus_units.RecID = bonus_awards.UnitType and ptRecID = " & ptRecID & " AND Min <= " & rsCount!RecCount & " and Max >= " & rsCount!RecCount)
'                            If rsSpiff.RecordCount > 0 Then
'                                Me.Caption = "Awarding Bonuses and Spiff"
'                                While Not rsSpiff.EOF And Err.Number = 0
'                                    Call MySQL.AddBonus(directConn, rsSpiff!FieldName, "Bonus for sale awarded", rsSpiff!units, rsSpiff!AwardTo, Login.lSysopID, Login.lAgencyID, Login.lVirtualID, 1, ptRecID, acciRecID, oSub.fRecID_acci_services)
'                                    rsSpiff.MoveNext
'                                Wend
'                            End If
'                        End If
'                        rsBonus.MoveNext
'                    Wend
'                End If
'            End If
'
            Me.Caption = oCaption
            
            JoiningFee = 0
            PerHour = 0
            PeriodFee = 0
            PerMB = 0
            'acciRecID = osub.fRecID
            BaseURL = ""
            
            ContactName = ""
            DynamicField1 = ""
            DynamicField2 = ""
            DynamicField3 = ""
            DynamicField4 = ""
            DynamicField5 = ""
            'NextCycle = Null
            Password = ""
            ptRecID = 0
            ServiceID = 0
            Username = ""
            DomainID = 0
            RadiusID = 0
            MBQuota = 0
            
            
            If oService(oSvCount).BillNow = True Then
            
                Dim BillCycle As Byte
                For bx = optBillingCycle.LBound To optBillingCycle.UBound
                    If optBillingCycle(bx).Value = True Then
                        BillCycle = bx
                        Exit For
                    End If
                Next
        
                If DateDiff("s", fAddPlan.Activation, sysnow) < 0 Then
                    mvBilling.Value = fAddPlan.Activation
                ElseIf oService(oSvCount).BillNow = True Then
                    Select Case oService(oSvCount).BillNow
                    Case -1
                    
                        If lvPlans.ListItems.Count > 0 Then bRefered = vbNo
                        
                        If bRefered = 0 Then
                            Screen.MousePointer = vbDefault
                            Select Case vbNo 'MsgBox("Was this transaction refered by another customer? If so the Joining fee is obmitted.", vbQuestion + vbYesNo, "Was customer refered")
                            Case vbYes
                                invout.AmountDue = oService(oSvCount).PeriodFee
                                invout.Description = oService(oSvCount).Description + " - Setup Free"
                                bRefered = vbYes
                            Case vbNo
                                invout.AmountDue = oService(oSvCount).PeriodFee + oService(oSvCount).JoiningFee
                                invout.Description = oService(oSvCount).Description + " - Setup Fee Included [" + Format(oService(oSvCount).JoiningFee, "Currency") + "]"
                                bRefered = vbNo
                            End Select
                            Screen.MousePointer = vbHourglass
                        Else
                            Select Case bRefered
                            Case vbYes
                                invout.AmountDue = oService(oSvCount).PeriodFee
                                invout.Description = oService(oSvCount).Description + " - Setup Free"
                                bRefered = vbYes
                            Case vbNo
                                invout.AmountDue = oService(oSvCount).PeriodFee + oService(oSvCount).JoiningFee
                                invout.Description = oService(oSvCount).Description + " - Setup Fee Included [" + Format(oService(oSvCount).JoiningFee, "Currency") + "]"
                                bRefered = vbNo
                            End Select
                        End If
                        
                        invout.PlanServiceID = osub.col_subServices.Count
                        invout.GSTCharged = invout.AmountDue * oTax(Login.TaxCode, Login.TaxCountry)
                        invout.TotalDue = invout.AmountDue + invout.GSTCharged
                        invout.AmountPaid = 0
                        invout.SysopID = Login.lSysopID
                        invout.VirtualID = Login.lVirtualID
                        invout.FlagID = FlagID
                        
                        invout.acci_RecID = osub.fRecID
                        
                        For bx = optBillingCycle.LBound To optBillingCycle.UBound
                            If optBillingCycle(bx).Value = True Then
                                invout.PaymentDue = DateAdd(optBillingCycle(bx).Tag, Val(txtBillingCycle(bx).Text), sysnow)
                                Exit For
                            End If
                        Next
                    
'
'                            On Error Resume Next
'                            Do
'                                Err.Clear
'                                invout.RecID = MySQL.GetTMPRecID("invoiceout", directConn)
'                                If invout.SubRecID = 0 Then invout.SubRecID = invout.RecID
'                                'Clipboard.Clear
'                                'Clipboard.SetText "INSERT INTO invoiceout (RecID, Description, PaidWhen, AmountDue, GSTCharged, " + _
'                                                        "TotalDue, AmountRefunded, GSTRefunded, AmountPaid, acci_RecID, SysopID, PlanServiceID, SubRecID, StartCycle, EndCycle, PaymentDue, VirtualID, StatementID, sfCycle_Download, sfCycle_Upload, sfCycle_Mins) " + _
'                                                        "VALUES ('" & invout.RecID & "','" & MySQL.ESC(invout.Description) & "','" & Format(invout.PaidWhen, "YYYY-MM-DD ttttt") & "','" & invout.AmountDue & "','" & invout.GSTCharged & _
'                                                        "','" & invout.TotalDue & "','" & invout.AmountRefunded & "','" & invout.GSTRefunded & "','" & invout.AmountPaid & "','" & invout.acci_RecID & "','" & invout.SysopID & _
'                                                        "','" & invout.PlanServiceID & "','" & invout.SubRecID & "','" & Format(invout.StartCycle, "YYYY-MM-DD ttttt") & "','" & Format(invout.EndCycle, "YYYY-MM-DD ttttt") & "','" & Format(invout.PaymentDue, "YYYY-MM-DD ttttt") & "','" & invout.VirtualID & "','" & invout.StatementID & "','" & invout.sfCycle_Download & "','" & invout.sfCycle_Upload & "','" & invout.sfCycle_Mins & "')"
'
'                                'Stop
'                                directConn.Execute "INSERT INTO invoiceout (RecID, Description, PaidWhen, AmountDue, GSTCharged, " + _
'                                                        "TotalDue, AmountRefunded, GSTRefunded, AmountPaid, acci_RecID, SysopID, PlanServiceID, SubRecID, StartCycle, EndCycle, PaymentDue, VirtualID, StatementID, sfCycle_Download, sfCycle_Upload, sfCycle_Mins) " + _
'                                                        "VALUES ('" & invout.RecID & "','" & MySQL.ESC(invout.Description) & "','" & Format(invout.PaidWhen, "YYYY-MM-DD ttttt") & "','" & invout.AmountDue & "','" & invout.GSTCharged & _
'                                                        "','" & invout.TotalDue & "','" & invout.AmountRefunded & "','" & invout.GSTRefunded & "','" & invout.AmountPaid & "','" & invout.acci_RecID & "','" & invout.SysopID & _
'                                                        "','" & invout.PlanServiceID & "','" & invout.SubRecID & "','" & Format(invout.StartCycle, "YYYY-MM-DD ttttt") & "','" & Format(invout.EndCycle, "YYYY-MM-DD ttttt") & "','" & Format(invout.PaymentDue, "YYYY-MM-DD ttttt") & "','" & invout.VirtualID & "','" & invout.StatementID & "','" & invout.sfCycle_Download & "','" & invout.sfCycle_Upload & "','" & invout.sfCycle_Mins & "')"
'
'                                  If Err.Number <> 0 Then cDebug Err.Description
'                            Loop Until Err.Number = 0
'
'                            On Error GoTo ErrorOccur
                            On Error GoTo 0
                        With osub.col_subTrans.Add("NEW" & tmpSession & osub.col_subTrans.Count + 1, 0, osub.fRecID, invout.AmountDue, invout.GSTCharged, DateAdd(optBillingCycle(BillCycle).Tag, Val(txtBillingCycle(BillCycle).Text), sysnow), invout.AmountPaid, sysnow, True, 0, invout.TotalDue, 0, _
                           0, 0, Login.lAgencyID, Login.lVirtualID, invout.Description, 0, 0, osub.col_subServices.Count, invout.AmountRefunded, invout.GSTRefunded, Login.lSysopID, sysnow, SubRecID, 0, 0, 0, oService.Activated, DateAdd(optBillingCycle(BillCycle).Tag, Val(txtBillingCycle(BillCycle).Text), sysnow), _
                            oService(oSvCount).ptRecID, oService(oSvCount).ServiceID, IIf(domainset = True, osub.col_Domains.Count, 0), IIf(radiuset = True, osub.col_subRadius.Count, 0), "NEW" & tmpSession & osub.col_subTrans.Count + 1)
                            
                            Set itmX = lvTransactions.ListItems.Add(, .Key, .Description)
                            itmX.SubItems(1) = Format(.TotalDue, "Currency")
                            itmX.SubItems(2) = Format(.GSTCharged, "Currency")
                            itmX.SubItems(3) = .PaymentDue
                            itmX.SubItems(4) = Format(.AmountPaid, "Currency")
                            itmX.SubItems(5) = IIf(.AmountPaid + .AmountRefunded >= .TotalDue, "Paid", IIf(.AmountPaid + .AmountRefunded > 0, "Partial", "No Payment"))
                            itmX.ForeColor = IIf(.AmountPaid + .AmountRefunded >= .TotalDue, clrPaid, IIf(.AmountPaid + .AmountRefunded > 0, clrPartial, clrNoPay))
                            itmX.Checked = .Checked
                        End With
                    End Select
                End If
            End If
        End If
            Next
        End If
    
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

Private Sub cmdBill_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdBill_Click"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    Dim ffrmAddPlan As New frmAddPlan
    
    ffrmAddPlan.Show 1
    
    If ffrmAddPlan.ptRecID = 0 Then Exit Sub
    
    Dim rsinvoiceout As adodb.Recordset
    Set rsinvoiceout = New adodb.Recordset
    Dim rsload As adodb.Recordset
    Set rsload = New adodb.Recordset
    Dim rsAccInfo As adodb.Recordset
    Set rsAccInfo = New adodb.Recordset
    
    Dim bResult As Boolean
    
    bResult = MySQL.OpenTable(directConn, rsinvoiceout, , "select * from invoiceout Limit 1")
    bResult = MySQL.OpenTable(directConn, rsload, , "select * from plantypes where RecID = " & ffrmAddPlan.ptRecID & " Limit 1")
        
    rsinvoiceout.AddNew
    
    Dim fNote As frmShortNote
    
    If Login.bMaster = True Then
        Set fNote = New frmShortNote
        fNote.sDescription = "Transaction Description"
        fNote.sShortNote = rsload!Description
        fNote.Show 1
    End If
    
    rsinvoiceout!Description = IIf(fNote.sShortNote = "", rsload!Description, fNote.sShortNote)
    rsinvoiceout!AmountDue = rsload!PeriodFee
    rsinvoiceout!GSTCharged = rsload!PeriodFee * oTax(Login.TaxCode, Login.TaxCountry)
    rsinvoiceout!TotalDue = rsinvoiceout!AmountDue + rsinvoiceout!GSTCharged
    rsinvoiceout!SysopID = Login.lSysopID
    rsinvoiceout!PlanServiceID = rsload!RecID
    rsinvoiceout!AgencyID = Login.lAgencyID
    rsinvoiceout!AmountPaid = 0
    
    If osub.fRecID <> 0 And osub.fRecID = 0 Then
        rsinvoiceout!acci_RecID = osub.fRecID
    ElseIf osub.fRecID = 0 And osub.fRecID <> 0 Then
        rsinvoiceout!acci_RecID = osub.fRecID
    End If
    
    Dim bx As Byte
    
    For bx = optBillingCycle.LBound To optBillingCycle.UBound
        If optBillingCycle(bx).Value = True Then
            rsinvoiceout!PaymentDue = DateAdd(optBillingCycle(bx).Tag, Val(txtBillingCycle(bx).Text), sysnow)
            Exit For
        End If
    Next
    
    Dim frmPayments As New frmAccPayment
    
    frmPayments.acci_RecID = rsinvoiceout!acci_RecID
    frmPayments.c_Description = rsload!Description
    frmPayments.l_RecID = MySQL.SetRecID(rsinvoiceout, "invoiceout", directConn)
    frmPayments.c_TotalDue = rsinvoiceout!TotalDue
    frmPayments.c_TotalPaid = 0
    
    If Login.bMaster = True Then frmPayments.Show 1
    

    bResult = MySQL.OpenTable(directConn, rsload, , "select * from invoiceout where RecID = " & frmPayments.l_RecID & " Limit 1")
    
    If rsload.RecordCount > 0 Then
        rsload.MoveFirst
        While Not rsload.EOF And Err.Number = 0
            Set itmX = lvTransactions.ListItems.Add(, "r" & rsload!RecID, rsload!Description)
            itmX.SubItems(1) = Format(rsload!TotalDue, "Currency")
            itmX.SubItems(2) = Format(rsload!GSTCharged, "Currency")
            itmX.SubItems(3) = rsload!PaymentDue
            itmX.SubItems(4) = Format(rsload!AmountPaid, "Currency")
            itmX.SubItems(5) = IIf(rsload!AmountPaid + rsload!AmountRefunded >= rsload!TotalDue, "Yes", "No")
            itmX.Checked = IIf(rsload!Checked <> 0, True, False)
            rsload.MoveNext
        Wend
    
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

Private Sub cmdCancel_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdCancel_Click"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    Select Case MsgBox("Are you sure you wish to cancel this entire edit?", vbQuestion & vbYesNo, "Cancel?")
    Case vbYes
        If osub.fRecID = 0 Then
            FormState = Deleted
        Else
            FormState = Cancelled
        End If
        Unload Me
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

Private Sub cmdCC_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdCC_Click"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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



Dim fFrmCC As New frmCreditCard
    Dim iY As Variant
    
    bselected = False
                
                '¤¤¤ Query Local DB for Previous History
                '
                ' --> HERE
                '
                '¤
    
    If osub.fRecID = 0 Then
        osub.fRecID = IIf(osub.fRecID = 0, MySQL.GetTMPRecID("accountinfo", directConn), osub.fRecID)
    End If
    
    Dim rs_creditcard As adodb.Recordset
    Dim rs_cc_Receipt As adodb.Recordset
    
    bResult = MySQL.OpenTable(directConn, rs_creditcard, , "select RecID, AccI_RecID as acci_RecID, bType, AES_DECRYPT(CardNumber,'" + odb.colSalts.ReturnSalt("CCSalt") + "" & odb.colSalts.ReturnSalt("md5Password") & "') as CardNumber, SecurityNumber, ExpiryDate, Name from creditcard where AccI_RecID = " & osub.fRecID)
    bResult = MySQL.OpenTable(directConn, rs_cc_Receipt, , "select * from cc_Receipt limit 1")
    
    
    If rs_creditcard.RecordCount > 0 Then
        rs_creditcard.MoveFirst
        While Not rs_creditcard.EOF And Err.Number = 0
            fFrmCC.cmdName.AddItem rs_creditcard!Name
            fFrmCC.cmdName.ItemData(fFrmCC.cmdName.ListCount - 1) = rs_creditcard!RecID
            rs_creditcard.MoveNext
        Wend
        fFrmCC.cmdName.ListIndex = 0
        rs_creditcard.Filter = ""
    Else
        rs_creditcard.Filter = ""
    End If
    
    fFrmCC.Show 1
    If fFrmCC.iCloseState = frmCloseCancel Then Exit Sub
    bSaveCC = True

    '¤
    If bSaveCC = True Then
        If osub.fRecID <> 0 Then
            rs_creditcard.Filter = "AccI_RecID = " & osub.fRecID & " AND CardNumber = '" & MySQL.NumCrypt(fFrmCC.screditcard) & "'"
        Else
            If osub.fRecID <> 0 Then rs_creditcard.Filter = "AccI_RecID = " & osub.fRecID & " AND CardNumber = '" & MySQL.NumCrypt(fFrmCC.screditcard) & "'"
        End If
        If rs_creditcard.RecordCount = -1 Or rs_creditcard.RecordCount = 0 Then
            Dim RecID As Long
            
            On Error Resume Next
            
            Do
                Err.Clear
                RecID = MySQL.GetTMPRecID("creditcard", directConn)
                'CreditCard 151
                MySQL.Execute directConn, "INSERT INTO creditcard (RecID, acci_RecID, SecurityNumber, Name, bType, ExpiryDate) VALUES " + _
                                        "('" & RecID & "','" & osub.fRecID & "','" & fFrmCC.sSecurityNo & "','" & fFrmCC.sCardName & "','" & fFrmCC.bType & "','" & fFrmCC.sExpiry & "')"
                                        
            Loop Until Err.Number = 0
            MySQL.Execute directConn, "UPDATE creditcard SET CardNumber=AES_ENCRYPT('" & MySQL.NumCrypt(fFrmCC.screditcard) & "','" + odb.colSalts.ReturnSalt("CCSalt") + "" & odb.colSalts.ReturnSalt("md5Password") & "') where RecID = " & RecID
            If fFrmCC.bDefault = True Then
                MySQL.Execute directConn, "UPDATE creditcard SET bDefault=0 where acci_RecID = " & RecID
                MySQL.Execute directConn, "UPDATE creditcard SET bDefault=-1 where RecID = " & RecID
            End If
        Else
            
            MySQL.Execute directConn, "UPDATE creditcard SET SecurityNumber = '" & fFrmCC.sSecurityNo & "', Name = '" & fFrmCC.sCardName & "', bType = '" & fFrmCC.bType & "', ExpiryDate = '" & fFrmCC.sExpiry & "' where RecID = " & rs_creditcard!RecID
            
            MySQL.Execute directConn, "UPDATE creditcard SET CardNumber=AES_ENCRYPT('" & MySQL.NumCrypt(fFrmCC.screditcard) & "','" + odb.colSalts.ReturnSalt("CCSalt") + "" & odb.colSalts.ReturnSalt("md5Password") & "') where RecID = " & rs_creditcard!RecID
            If fFrmCC.bDefault = True Then
                MySQL.Execute directConn, "UPDATE creditcard SET bDefault=0 where acci_RecID = " & rs_creditcard!acci_RecID
                MySQL.Execute directConn, "UPDATE creditcard SET bDefault=-1 where RecID = " & rs_creditcard!RecID
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

Private Sub cmdChangePlan_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdChangePlan_Click"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    
    If lvPlans.SelectedItem Is Nothing Then Exit Sub
    
    Dim lvselect As ListItem
    
    Dim ffrmAddPlan As New frmAddPlan
    Set ListItem = lvPlans.SelectedItem
    ffrmAddPlan.lServiceID = ListItem.Tag
    ffrmAddPlan.Show 1
    
    Dim bResult As Boolean
    Dim rsSave As adodb.Recordset
    Dim rsload As adodb.Recordset

    Dim rsinvoicein As adodb.Recordset
    Dim rsinvoiceout As adodb.Recordset
    Dim rsinvoiceout2 As adodb.Recordset
    Dim rsPlanType As adodb.Recordset
    Dim rsAccInfo As adodb.Recordset
    
    
    If ffrmAddPlan.ptRecID <> 0 Then
    
        bResult = MySQL.OpenTable(directConn, rsSave, , "select * from acci_services where RecID = " + Mid(ListItem.Key, 2) + " Limit 1")
        bResult = MySQL.OpenTable(directConn, rsload, , "select * from plantypes where RecID = " & ffrmAddPlan.ptRecID & " Limit 1")
        bResult = MySQL.OpenTable(directConn, rsPlanType, , "select * from plantypes where RecID = " & rsSave!ptRecID & " Limit 1")
        bResult = MySQL.OpenTable(directConn, rsinvoicein, , "select * from invoicein Limit 1")
        bResult = MySQL.OpenTable(directConn, rsinvoiceout, , "select * from invoiceout Limit 1")
        
        Dim iTotalHours As Integer
        Dim iDiffeHours As Integer
        
        iTotalHours = DateDiff("h", CDate(IIf(IsNull(rsSave!PreviousCycle), rsSave!NextCycle, rsSave!PreviousCycle)), CDate(rsSave!NextCycle))
        iDiffeHours = DateDiff("h", sysnow, CDate(rsSave!NextCycle))
        
              
        If iTotalHours <> 0 Then
            
            rsinvoicein.AddNew
            
            rsinvoicein!AmountPaid = rsSave!PeriodFee - ((rsSave!PeriodFee / iTotalHours) * iDiffeHours)
            rsinvoicein!GSTCharged = (rsSave!PeriodFee - ((rsSave!PeriodFee / iTotalHours) * iDiffeHours)) * oTax(Login.TaxCode, Login.TaxCountry)
            rsinvoicein!Sub = 0
            rsinvoicein!TotalPaid = rsinvoicein!AmountPaid + rsinvoicein!GSTCharged
            rsinvoicein!VirtualID = Login.lVirtualID
            rsinvoicein!AmountUsed = 0
        
        
    
            If osub.fRecID <> 0 And osub.fRecID = 0 Then
                rsinvoicein!acci_RecID = osub.fRecID
            ElseIf osub.fRecID = 0 And osub.fRecID <> 0 Then
                rsinvoicein!acci_RecID = osub.fRecID
            End If
            
            rsinvoicein.Update
            
        End If
               
        Dim rsRadius As adodb.Recordset
                
        If rsSave!RadiusID <> 0 Then
            
            bResult = MySQL.OpenTable(directConn, rsRadius, , "select RecID, sfCycle_Upload , sfCycle_Download, sfCycle_Mins from radiusaccounts Where RecID = " & rsSave!RadiusID)
        
            cCharge = 0
            
            If rsPlanType.RecordCount > 0 Then
                
                bResult = MySQL.OpenTable(directConn, rsAccInfo, , "select * from accountinfo Where RecID = " & rsSave!acci_RecID)
                
                rsinvoiceout.AddNew
                                    
                rsinvoiceout!Description = rsPlanType!Description
                
                If rsPlanType!MBPerPeriod <> -1 Then
                    If rsRadius!sfCycle_Download / lBytesPerMB > rsPlanType!MBPerPeriod Then
                        rsinvoiceout!Description = rsinvoiceout!Description + " " & (rsRadius!sfCycle_Download / lBytesPerMB) - rsPlanType!MBPerPeriod & "MB's Over threshold"
                        
                        cCharge = cCharge + ((rsRadius!sfCycle_Download / 1048576 - rsPlanType!MBPerPeriod) / rsPlanType!MBBlockSize) * rsSave!PerMB
                    End If
                End If
                
                rsinvoiceout!sfCycle_Download = rsRadius!sfCycle_Download
                If Not rsAccInfo.EOF And Not rsAccInfo.BOF Then rsAccInfo!sfCycle_Download = IIf(IsNull(rsAccInfo!sfCycle_Download), 0, Val(rsAccInfo!sfCycle_Download)) - IIf(IsNull(rsinvoiceout!sfCycle_Download), 0, Val(rsinvoiceout!sfCycle_Download))
                
                rsinvoiceout!sfCycle_Upload = rsRadius!sfCycle_Upload
                If Not rsAccInfo.EOF And Not rsAccInfo.BOF Then rsAccInfo!sfCycle_Upload = rsAccInfo!sfCycle_Upload - rsinvoiceout!sfCycle_Upload
                
                directConn.Execute "Insert Into history_radius_datausage (RadiusID, Uploaded, Downloaded, NumMins) VALUES(" & rsRadius!RecID & ", " & rsRadius!sfCycle_Upload & ", " & rsRadius!sfCycle_Download & ", " & rsRadius!sfCycle_Mins & ")"
                
                rsRadius!sfCycle_Download = 0
                rsRadius!sfCycle_Upload = 0
                
                If rsPlanType!HoursPerPeriod <> -1 Then
                    If rsRadius!sfCycle_Mins / 60 > rsPlanType!HoursPerPeriod Then
                        cCharge = cCharge + (rsRadius!sfCycle_Mins / 60) * rsSave!PerHour
                    End If
                End If
                
                rsinvoiceout!sfCycle_Mins = rsRadius!sfCycle_Mins
                If Not rsAccInfo.EOF And Not rsAccInfo.BOF Then rsAccInfo!sfCycle_Mins = rsAccInfo!sfCycle_Mins - rsinvoiceout!sfCycle_Mins
                
                rsRadius!sfCycle_Mins = 0
                cCharge = cCharge '+ rsSave!PeriodFee
                Call MySQL.SetNextCycle(rsSave, IIf(IsNull(rsPlanType!chgIntervalType), "m", rsPlanType!chgIntervalType), IIf(IsNull(rsPlanType!chgInterval), 1, rsPlanType!chgInterval))
                                                 
                ' Calculates any prepaid disposition with currency.
                bResult = MySQL.OpenTable(directConn, rsinvoicein, , "select * from invoicein Where AmountPaid > AmountUsed AND AccI_RecID = " & rsSave!acci_RecID)
                cTotalDue = 0
                cAmountPaid = 0
                cTotalDue = cCharge + cCharge * oTax(Login.TaxCode, Login.TaxCountry)
                If rsinvoicein.RecordCount > 0 Then
                    Do While Not rsinvoicein.EOF And Err.Number = 0
                        If rsinvoicein!AmountPaid - rsinvoicein!AmountUsed > 0 Then
                            cPrePaid = rsinvoicein!AmountPaid - rsinvoicein!AmountUsed
                            If cPrePaid < cTotalDue Then
                                cTotalDue = cTotalDue - cPrePaid
                                cAmountPaid = cAmountPaid + cPrePaid
                                rsinvoicein!AmountUsed = rsinvoicein!AmountPaid
                                rsinvoicein.Update
                                rsinvoiceout!PaidWhen = sysnow
                            ElseIf cPrePaid > cTotalDue Then
                                rsinvoicein!AmountUsed = rsinvoicein!AmountUsed + cTotalDue
                                cAmountPaid = cAmountPaid + cTotalDue
                                cTotalDue = 0
                                rsinvoicein.Update
                                rsinvoiceout!PaidWhen = sysnow
                            End If
                        End If
                        rsinvoicein.MoveNext
                        If cTotalDue = 0 Then Exit Do
                    Loop
                End If
                
                ' Saves the entry for the account Invoice going out
                rsinvoiceout!AmountDue = cCharge
                rsinvoiceout!GSTCharged = cCharge * oTax(Login.TaxCode, Login.TaxCountry)
                rsinvoiceout!TotalDue = cTotalDue
                        
                rsinvoiceout!AmountPaid = cAmountPaid
                rsinvoiceout!acci_RecID = rsSave!acci_RecID
                rsinvoiceout!PlanServiceID = rsPlanType!RecID
                bResult = MySQL.OpenTable(directConn, rsAccInfo, , "select * from accountinfo Where RecID = " & rsSave!acci_RecID)
                If rsAccInfo.RecordCount > 0 Then
                    rsinvoiceout!PaymentDue = DateAdd(IIf(IsNull(rsAccInfo!PayIntervalType), "d", rsAccInfo!PayIntervalType), IIf(IsNull(rsAccInfo!PayInterval), 14, rsAccInfo!PayInterval), sysnow)
                Else
                    rsinvoiceout!PaymentDue = DateAdd("d", 14, sysnow)
                End If
                rsinvoiceout!VirtualID = Login.lVirtualID
                
                Call MySQL.SetRecID(rsinvoiceout, "invoiceout", directConn)
                
            Else
            
            End If

        End If

        Dim bx As Byte
        
        For bx = optBillingCycle.LBound To optBillingCycle.UBound
            If optBillingCycle(bx).Value = True Then
                rsinvoiceout!PaymentDue = DateAdd(optBillingCycle(bx).Tag, Val(txtBillingCycle(bx).Text), sysnow)
                Exit For
            End If
        Next
        
        Call MySQL.SetNextCycle(rsSave, rsload!chgIntervalType, rsload!chgInterval)
        
        MySQL.Execute directConn, "UPDATE acci_services SET ServiceID=" & rsload!ServiceID & ",ptRecID=" & ffrmAddPlan.ptRecID & ",JoiningFee=" & oService(1).JoiningFee & ", PeriodFee = " & oService(1).PeriodFee & ", PerMB = " & oService(1).PerMBBlock & ", PerHour = " & oService(1).PerHour & " Where RecID = " + Mid(ListItem.Key, 2)
        bResult = MySQL.OpenTable(directConn, rsSave, , "select * from acci_services where RecID = " + Mid(ListItem.Key, 2) + " Limit 1")
        
        If rsSave!RadiusID <> 0 Then
            
            bResult = MySQL.OpenTable(directConn, rsRadius, , "select RecID, sfCycle_Upload , sfCycle_Download, sfCycle_Mins from radiusaccounts Where RecID = " & rsSave!RadiusID)
            bResult = MySQL.OpenTable(directConn, rsinvoiceout2, , "select * from invoiceout Limit 1")
            
            cCharge = 0
            
            If rsPlanType.RecordCount > 0 Then
                
                bResult = MySQL.OpenTable(directConn, rsAccInfo, , "select * from accountinfo Where RecID = " & rsSave!acci_RecID)
                
                rsinvoiceout2.AddNew
                                    
                rsinvoiceout2!Description = rsPlanType!Description
                
                If rsload!MBPerPeriod <> -1 Then
                    If rsRadius!sfCycle_Download / lBytesPerMB > rsload!MBPerPeriod Then
                        rsinvoiceout2!Description = rsinvoiceout2!Description + " " & (rsRadius!sfCycle_Download / lBytesPerMB) - rsload!MBPerPeriod & "MB's Over threshold"
                        cCharge = cCharge + ((rsRadius!sfCycle_Download / 1048576 - rsload!MBPerPeriod) / rsload!MBBlockSize) * rsload!FeePerBlock
                    End If
                End If
                
                rsinvoiceout2!sfCycle_Download = rsRadius!sfCycle_Download
                If Not rsAccInfo.EOF And Not rsAccInfo.BOF Then rsAccInfo!sfCycle_Download = rsAccInfo!sfCycle_Download - rsinvoiceout2!sfCycle_Download
                rsinvoiceout2!sfCycle_Upload = rsRadius!sfCycle_Upload
                If Not rsAccInfo.EOF And Not rsAccInfo.BOF Then rsAccInfo!sfCycle_Upload = rsAccInfo!sfCycle_Upload - rsinvoiceout2!sfCycle_Upload
                
                directConn.Execute "Insert Into history_radius_datausage (RadiusID, Uploaded, Downloaded, NumMins) VALUES(" & rsRadius!RecID & ", " & rsRadius!sfCycle_Upload & ", " & rsRadius!sfCycle_Download & ", " & rsRadius!sfCycle_Mins & ")"
                
                rsRadius!sfCycle_Download = 0
                rsRadius!sfCycle_Upload = 0
                
                If rsload!HoursPerPeriod <> -1 Then
                    If rsRadius!sfCycle_Mins / 60 > rsload!HoursPerPeriod Then
                        cCharge = cCharge + (rsRadius!sfCycle_Mins / 60) * rsload!ExtraPerHour
                    End If
                End If
                
                rsinvoiceout2!sfCycle_Mins = rsRadius!sfCycle_Mins
                If Not rsAccInfo.EOF And Not rsAccInfo.BOF Then rsAccInfo!sfCycle_Mins = rsAccInfo!sfCycle_Mins - rsinvoiceout2!sfCycle_Mins
                
                rsRadius!sfCycle_Mins = 0
                cCharge = cCharge + rsload!PeriodFee
                Call MySQL.SetNextCycle(rsSave, IIf(IsNull(rsload!chgIntervalType), "m", rsload!chgIntervalType), IIf(IsNull(rsload!chgInterval), 1, rsload!chgInterval))
                                                 
                ' Calculates any prepaid disposition with currency.
                bResult = MySQL.OpenTable(directConn, rsinvoicein, , "select * from invoicein Where AmountPaid > AmountUsed AND AccI_RecID = " & rsSave!acci_RecID)
                cTotalDue = 0
                cAmountPaid = 0
                cTotalDue = cCharge + cCharge * oTax(Login.TaxCode, Login.TaxCountry)
                If rsinvoicein.RecordCount > 0 Then
                    Do While Not rsinvoicein.EOF And Err.Number = 0
                        If rsinvoicein!AmountPaid - rsinvoicein!AmountUsed > 0 Then
                            cPrePaid = rsinvoicein!AmountPaid - rsinvoicein!AmountUsed
                            If cPrePaid < cTotalDue Then
                                cTotalDue = cTotalDue - cPrePaid
                                cAmountPaid = cAmountPaid + cPrePaid
                                rsinvoicein!AmountUsed = rsinvoicein!AmountPaid
                                rsinvoicein.Update
                                rsinvoiceout2!PaidWhen = sysnow
                            ElseIf cPrePaid > cTotalDue Then
                                rsinvoicein!AmountUsed = rsinvoicein!AmountUsed + cTotalDue
                                cAmountPaid = cAmountPaid + cTotalDue
                                cTotalDue = 0
                                rsinvoicein.Update
                                rsinvoiceout2!PaidWhen = sysnow
                            End If
                        End If
                        rsinvoicein.MoveNext
                        If cTotalDue = 0 Then Exit Do
                    Loop
                End If
                
                ' Saves the entry for the account Invoice going out
                rsinvoiceout2!AmountDue = cCharge
                rsinvoiceout2!GSTCharged = cCharge * oTax(Login.TaxCode, Login.TaxCountry)
                rsinvoiceout2!TotalDue = cTotalDue
                        
                rsinvoiceout2!AmountPaid = cAmountPaid
                rsinvoiceout2!acci_RecID = rsSave!acci_RecID
                rsinvoiceout2!PlanServiceID = rsload!RecID
                bResult = MySQL.OpenTable(directConn, rsAccInfo, , "select * from accountinfo Where RecID = " & rsSave!acci_RecID)
                If rsAccInfo.RecordCount > 0 Then
                    rsinvoiceout2!PaymentDue = DateAdd(IIf(IsNull(rsAccInfo!PayIntervalType), "d", rsAccInfo!PayIntervalType), IIf(IsNull(rsAccInfo!PayInterval), 14, rsAccInfo!PayInterval), sysnow)
                Else
                    rsinvoiceout2!PaymentDue = DateAdd("d", 14, sysnow)
                End If
                rsinvoiceout2!VirtualID = Login.lVirtualID
                
                Call MySQL.SetRecID(rsinvoiceout2, "invoiceout", directConn)
                
            Else
            
            End If
        End If
            
        
        Dim itmX As ListItem
        Set itmX = ListItem
        itmX.Tag = rsSave!ServiceID
        itmX.Text = IIf(IsNull(rsSave!ContactName), "", rsSave!ContactName)
        itmX.Checked = True
        itmX.SubItems(1) = rsload!Description
        itmX.SubItems(2) = IIf(IsNull(rsSave!Username), "", rsSave!Username)
        itmX.SubItems(3) = IIf(IsNull(rsSave!Password), "", rsSave!Password)
        itmX.SubItems(4) = IIf(IsNull(rsSave!BaseURL), "", rsSave!BaseURL)
        itmX.SubItems(5) = IIf(IsNull(rsSave!DynamicField1), "", rsSave!DynamicField1)
        itmX.SubItems(6) = IIf(IsNull(rsSave!DynamicField2), "", rsSave!DynamicField2)
        itmX.SubItems(7) = IIf(IsNull(rsSave!DynamicField3), "", rsSave!DynamicField3)
        itmX.SubItems(8) = IIf(IsNull(rsSave!DynamicField4), "", rsSave!DynamicField4)
        itmX.SubItems(9) = IIf(IsNull(rsSave!DynamicField5), "", rsSave!DynamicField5)
             
                
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

Private Sub cmdPaymentTypeAdd_Click(Index As Integer)


    Select Case Index
    Case 0 'Direct Debit
        If Len(txtddAccountName) = 0 Or Len(txtddBSB) = 0 Or Len(txtddAccountNumber) = 0 Then
            MsgBox "You have not completed all the details for Direct Debit. Please review and set again!"
            Exit Sub
        End If
        txtccCardName = ""
        txtccCardNumber = ""
        txtccCardExpiry = ""
        txtccCIC = ""
        txtswWord = ""
        txtswNumber = ""
        
    Case 1 'Credit card
    
        If Len(txtccCardName) = 0 Or Len(txtccCardNumber) < 9 Or Len(txtccCardExpiry) = 0 Or Len(txtccCIC) = 0 Then
            MsgBox "You have not completed all the details for a credit card transaction. Please review and set again!"
            Exit Sub
        End If
        
        txtccCardNumber = MySQL.ReplaceString(txtccCardNumber, " ", "")
        If CheckCard(txtccCardNumber) = False Then
            MsgBox "You have entered an invalid Credit Card Number. Please check your card number with the customer and enter it again!"
            Exit Sub
        End If
        
        txtddAccountName = ""
        txtddBSB = ""
        txtddAccountNumber = ""
        txtccCIC = ""
        txtswWord = ""
        txtswNumber = ""
    Case 2 'Statement Watcher
    
        If Len(txtswWord) = 0 And Len(txtswNumber) = 0 Then
            MsgBox "You have not completed all the details for a statement watcher transaction. Please review and set again!"
            Exit Sub
        End If
        
        txtccCardName = ""
        txtccCardNumber = ""
        txtccCardExpiry = ""
        txtccCIC = ""
        txtddAccountName = ""
        txtddBSB = ""
        txtddAccountNumber = ""
    End Select
    
    With osub.colPaymentSet.Add("NEW" & osub.colPaymentSet.Count, txtddAccountName, txtddBSB, txtddAccountNumber, _
        txtccCardName, txtccCardNumber, txtccCardExpiry, txtccCIC, _
        txtswWord, txtswNumber, 0, 0, "NEW" & osub.colPaymentSet.Count)
        .Checked = True
    End With

    txtswWord = ""
    txtswNumber = ""
    txtccCardName = ""
    txtccCardNumber = ""
    txtccCardExpiry = ""
    txtccCIC = ""
    txtddAccountName = ""
    txtddBSB = ""
    txtddAccountNumber = ""
    
    Dim oPay As clsPaymentSet
    Dim itmX As ListItem
    
    If osub.colPaymentSet.Count > 0 Then
        lvPayment.ListItems.Clear
        For Each oPay In osub.colPaymentSet
            If Len(oPay.ccCardName) > 0 Then
            
                Set itmX = lvPayment.ListItems.Add(, oPay.Key, "Credit Card")
                itmX.SubItems(1) = "100%"
                itmX.SubItems(2) = "Card Details: " & oPay.ccCardNumber & " - " & oPay.ccCardName & " - " & oPay.ccCardExpiry
                itmX.Tag = oPay.IDX
                itmX.Checked = oPay.Checked
                
            ElseIf Len(oPay.ddAccountName) > 0 Then
            
                Set itmX = lvPayment.ListItems.Add(, oPay.Key, "Direct Debit")
                itmX.SubItems(1) = "100%"
                itmX.SubItems(2) = "Direct Debit Details: " & oPay.ddAccountName & " - " & oPay.ddAcountNo & " - " & oPay.ddBSB
                itmX.Tag = oPay.IDX
                itmX.Checked = oPay.Checked
            Else
            
                Set itmX = lvPayment.ListItems.Add(, oPay.Key, "Statement Watcher")
                itmX.SubItems(1) = "100%"
                itmX.SubItems(2) = "Watcher Details: " & oPay.swNumber & " - " & oPay.swWord
                itmX.Tag = oPay.IDX
                itmX.Checked = oPay.Checked
                
            End If
            
        Next
    End If
        
End Sub

Private Sub cmdSave_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdSave_Click"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    If txtAccountName.Text = "" Then
        MsgBox "You must specify an account name for the connection of these services!", vbCritical, "Account Name Missing"
        tsHeader.Tabs(1).Selected = True
        txtAccountName.SetFocus
        Exit Sub
    End If
    
    
    If lvPhone.ListItems.Count = 0 Then
        MsgBox "You must specify at least one physical phone number for the connection of these services!", vbCritical, "Phone Number Missing"
        tsDetails.Tabs(2).Selected = True
        Exit Sub
    End If
    
    If lvEmail.ListItems.Count = 0 Then
        MsgBox "You must specify at least one physical eMail Address for the connection of these services!", vbCritical, "Email Address Missing"
        tsDetails.Tabs(2).Selected = True
        Exit Sub
    End If
       
    If bEditMade = True Then
        Dim ffrmShortNote As frmShortNote
        Set ffrmShortNote = New frmShortNote
        ffrmShortNote.sDescription = "Edit Log Entry/Description"
        ffrmShortNote.sShortNote = sTMPShortNote
        ffrmShortNote.Show 1
        osub.col_subEditLog.Add "NEW" & SESSION & osub.col_subEditLog.Count + 1, osub.fRecID, sysnow, Login.lSysopID, ffrmShortNote.sShortNote, "", 0, "NEW" & SESSION & osub.col_subEditLog.Count + 1
    End If
       
    osub.fRecID = SaveInformation()
    Const SendMail = True
    
    If osub.fRecID = 0 Then Exit Sub
        
    
    txtAccountName.Tag = osub.fRecID
    
    Saved = True
    FormState = Saved
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

Private Sub Command3_Click()

End Sub

Private Sub dtpBilling_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "dtpBilling_MouseMove"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    
    If bPlay = False And picTSContainer(2).Visible = True Then
            
        L = GetCursorPos(oXY)
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "This is where you can adjust the next billing time, if you set it in the past they will be billed immediately next time upkeep is done."
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

Private Sub dtpDOB_Change()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "dtpDOB_Change"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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

Private Sub dtpDOB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "dtpDOB_MouseMove"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "select the date of birth or the date of registration of the Organisation."
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

Function IsEven(M_num As Integer) As Boolean
    If M_num Mod 2 = 0 Then
        IsEven = True
    Else
        IsEven = False
    End If
    
End Function
Function CheckCard(CCNumber As String) As Boolean
    Dim Counter As Integer, TmpInt As Integer
    Dim Answer As Integer
    Counter = 1
    TmpInt = 0
    While Counter <= Len(CCNumber)
        If IsEven(Len(CCNumber)) Then
            TmpInt = Val(Mid$(CCNumber, Counter, 1))
            If Not IsEven(Counter) Then
                TmpInt = TmpInt * 2
                If TmpInt > 9 Then TmpInt = TmpInt - 9
            End If
            Answer = Answer + TmpInt
            'Debug.Print Counter, TmpInt, Answer
            Counter = Counter + 1
        Else
            TmpInt = Val(Mid$(CCNumber, Counter, 1))
            If IsEven(Counter) Then
                TmpInt = TmpInt * 2
                If TmpInt > 9 Then TmpInt = TmpInt - 9
            End If
            Answer = Answer + TmpInt
            'Debug.Print Counter, TmpInt, Answer
            Counter = Counter + 1
        End If
    Wend
    Answer = Answer Mod 10
    If Answer = 0 Then CheckCard = True Else CheckCard = False
End Function


Private Sub Form_Load()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Load"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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

    FormState = Loading
    
    Call tb_ButtonClick(tb.Buttons(1))
    
    Dim rsload As adodb.Recordset
    Call MySQL.OpenTable(directConn, rsload, , "select NOW() as formload")
    FormLoad = rsload!FormLoad
    rsload.Close
    
    
    'SavePicture lvPhone.Picture, "C:\Phone.bmp"
    'SavePicture lvEmail.Picture, "C:\email.bmp"
   
    LoadColumnWidths
        
    'Call LoadAccountTypes
    Call LoadVISPs
    Call LoadSysops
    
    Call LoadAccountClass
    Call LoadFlags
    'Call LoadHardware
    
    If osub.fRecID <> 0 Then
        
        Dim itmX As ListItem
        Dim rsLoadb As adodb.Recordset
        
        SQL = "Select " + _
              " AVG(audithistory.vartype) as AVGVartype, AVG(audithistory.varname) as AVGVarname , avg(audithistory.oldcur) as AVGoldcur , STD(audithistory.oldcur) as STDoldcur, AVG(audithistory.newcur) as AVGnewcur, STD(audithistory.newcur) as STDnewcur , AVG(sysops.Username) as AVGUsername, AVG(sysops.Firstname) as AVGFirstname, avg(sysops.Surname) as AVGSurname " + _
              " from audithistory, sysops where audithistory.sysopid = sysops.RecID " + _
              " and audithistory.acci_RecID = '" & osub.fRecID & "'"
                Dim aado2 As New adodb.Connection
                aado2.Open directConn.ConnectionString
                aado2.Execute "use projectalpha"
                
                
                If MySQL.OpenTable(aado2, rsLoadb, , CStr(SQL)) = True Then
                    If rsLoadb.State = adStateOpen Then
                        lvAuditStats.SmallIcons = fIcon.il16x16
                        With lvAuditStats.ListItems.Add(, , "Average Variable Type", , 1)
                            .SubItems(1) = IIf(IsNull(rsLoadb!AVGVartype), "None", rsLoadb!AVGVartype)
                        End With
                        With lvAuditStats.ListItems.Add(, , "Average Variable Name", , 3)
                            .SubItems(1) = IIf(IsNull(rsLoadb!AVGVarname), "None", rsLoadb!AVGVarname)
                        End With
                        With lvAuditStats.ListItems.Add(, , "Average old Currency", , 6)
                            .SubItems(1) = IIf(IsNull(rsLoadb!avgoldcur), "$ 0.00", rsLoadb!avgoldcur)
                        End With
                        With lvAuditStats.ListItems.Add(, , "Average new Currency", , 9)
                            .SubItems(1) = IIf(IsNull(rsLoadb!AVGnewcur), "$ 0.00", rsLoadb!AVGnewcur)
                        End With
                        With lvAuditStats.ListItems.Add(, , "Standard Deviation old Currency", , 16)
                            .SubItems(1) = IIf(IsNull(rsLoadb!stdnewcur), "$ 0.00", rsLoadb!stdnewcur)
                        End With
                        With lvAuditStats.ListItems.Add(, , "Standard Deviation new Currency", , 26)
                            .SubItems(1) = IIf(IsNull(rsLoadb!stdnewcur), "$ 0.00", rsLoadb!stdnewcur)
                        End With
                        'msfgAudit.Refresh
                    End If
                    
                End If
                
        
        SQL = "Select audithistory.RecID as 'Audit ID', audithistory.acci_RecID as 'Account Number', audithistory.Description  , audithistory.Checked, audithistory.systemstamp  , audithistory.sysnow , audithistory.localtime , audithistory.appname , audithistory.appversion , audithistory.apphdc , audithistory.formname , " + _
              " audithistory.formhwnd , audithistory.vartype , audithistory.varname , audithistory.oldcur , audithistory.newcur , audithistory.oldvalue  , audithistory.newvalue  , audithistory.oldpointer, sysops.Username, sysops.Firstname, sysops.Surname, virtualisp.Description " + _
              " from sysops inner join audithistory on sysops.RecID = audithistory.sysopid inner join virtualisp on audithistory.VirtualID = virtualisp.RecID where audithistory.sysopid = sysops.RecID  and audithistory.virtualid = virtualisp.RecID  " + _
              " and audithistory.acci_RecID = '" & osub.fRecID & "'"
              
'                MsgBox directConn.ConnectionString
                
                adoAudit.ConnectionString = directConn.ConnectionString
                adoAudit.CommandTimeout = directConn.CommandTimeout
                adoAudit.CursorLocation = adUseClient
                adoAudit.CursorType = adOpenStatic
                adoAudit.RecordSource = SQL
                
                '= rsLoadb
                adoAudit.Refresh
            
             
        Timer1.Tag = "Customer In Edit: "
        Call osub.GetClient(osub.fRecID, directConn)
        Call LoadInformation
    Else
        Timer1.Tag = "Customer Created: "
        'picPaymentFlag.BorderStyle = 1
    End If
    
    On Error Resume Next
    
    If Len(osub.fftpPathKey) = 0 Then
        Dim CharSTR As String
        Randomize Now
        CharSTR = Chr(Rnd * 255) + Chr(Rnd * 255) + Chr(Rnd * 255) + Chr(Rnd * 255) + Chr(Rnd * 255) + Chr(Rnd * 255) + Chr(Rnd * 255) + Chr(Rnd * 255) + Chr(Rnd * 255) + Chr(Rnd * 255) + Chr(Rnd * 255) + Chr(Rnd * 255) + Chr(Rnd * 255) + Chr(Rnd * 255) + Chr(Rnd * 255) + Chr(Rnd * 255) + Chr(Rnd * 255) + Chr(Rnd * 255)
        osub.fftpPathKey = MySQL.MD5(directConn, CharSTR)
    End If
    
    
    Dim ExtendedXML As String
    Dim IdentityXML As String
    
    IdentityXML = ""
    IdentityXML = IdentityXML & "<SysopID>" & Login.lSysopID & "</SysopID>"
    IdentityXML = IdentityXML & "<VirtualID>" & Login.lVirtualID & "</VirtualID>"
    IdentityXML = IdentityXML & "<AgencyID>" & Login.lAgencyID & "</AgencyID>"
    
    ExtendedXML = ExtendedXML & "<FileCategory>" & "Subscriber" & "</FileCategory>"
    


    
    Call dftMain.Startup(directConn, CStr(oResell(1).GetNetworkXML), IdentityXML, ExtendedXML, CStr(osub.fftpPathKey), CStr(oResell(1).ftpGroupingFolder), CStr(odb.colSalts.ReturnSalt("FileDB")))
    
    
       
    Call tsDetails_Click
    Call tsHeader_Click
    
    mvBilling.Value = Format(DateAdd("m", 1, sysnow), "dd/mm/yyyy")
    dtpBilling.Value = sysnow
    
    SetFormAccess

    FormState = Waiting
    
    Me.Show
    Me.Refresh
    DoEvents
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
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    
    If Saved = False And osub.fRecID = 0 Then
    
       ' MySQL.Execute directConn, "Delete from recidplacement where TableName = 'accountinfo' And RecID = " & osub.fRecID
        'MySQL.Execute directConn, "Delete from acci_services where acci_RecID = '" & osub.fRecID& "' and DateCreated >= '" & Format(FormLoad, "yyyy-mm-dd ttttt") & "'"
        'MySQL.Execute directConn, "Delete from acci_addresses where acci_RecID = '" & osub.fRecID& "' and DateCreated >= '" & Format(FormLoad, "yyyy-mm-dd ttttt") & "'"
        'MySQL.Execute directConn, "Delete from radiusaccounts where acci_RecID = '" & osub.fRecID& "' and DateCreated >= '" & Format(FormLoad, "yyyy-mm-dd ttttt") & "'"
        'MySQL.Execute directConn, "Delete from bonus_matrix where acci_RecID = '" & osub.fRecID& "' and Created >= '" & Format(FormLoad, "yyyy-mm-dd ttttt") & "'"
        
    End If
    
    SaveColumnWidths
    
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
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
        
        ffrmSnailMail.sContactName = itmX.Text
        ffrmSnailMail.sStreetLine1 = itmX.SubItems(1)
        ffrmSnailMail.sStreetLine2 = itmX.SubItems(2)
        ffrmSnailMail.sSuburb = itmX.SubItems(3)
        ffrmSnailMail.sState = itmX.SubItems(4)
        ffrmSnailMail.sPostcode = itmX.SubItems(5)
        ffrmSnailMail.sCountry = itmX.SubItems(6)
    
        ffrmSnailMail.Show 1
        
        If ffrmSnailMail.iCloseState = frmCloseSave Then
            
            osub.colSnailMail(itmX.Key).ContactName = ffrmSnailMail.sContactName
            osub.colSnailMail(itmX.Key).Street1 = ffrmSnailMail.sStreetLine1
            osub.colSnailMail(itmX.Key).Street2 = ffrmSnailMail.sStreetLine2
            osub.colSnailMail(itmX.Key).Country = ffrmSnailMail.sCountry
            osub.colSnailMail(itmX.Key).PostCode = ffrmSnailMail.sPostcode
            osub.colSnailMail(itmX.Key).Suburb = ffrmSnailMail.sSuburb
            osub.colSnailMail(itmX.Key).State = ffrmSnailMail.sState
            
            osub.colSnailMail(itmX.Key).Key = "UPDAT" & Right(osub.colSnailMail(itmX.Key).Key, Len(osub.colSnailMail(itmX.Key).Key) - 5)
            itmX.Key = "UPDAT" & Right(osub.colSnailMail(itmX.Key).Key, Len(osub.colSnailMail(itmX.Key).Key) - 5)
            
            
            itmX.Text = ffrmSnailMail.sContactName
            itmX.SubItems(1) = ffrmSnailMail.sStreetLine1
            itmX.SubItems(2) = ffrmSnailMail.sStreetLine2
            itmX.SubItems(3) = ffrmSnailMail.sSuburb
            itmX.SubItems(4) = ffrmSnailMail.sState
            itmX.SubItems(5) = ffrmSnailMail.sPostcode
            itmX.SubItems(6) = ffrmSnailMail.sCountry
            
            bEditMade = True
            
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
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    If Item.Key <> "" Then
        osub.colSnailMail(Item.Key).Checked = Item.Checked
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

Private Sub lvAddresses_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvAddresses_ItemClick"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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

Private Sub lvAddresses_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvAddresses_MouseMove"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "Here is the list of postal address and mail address such as po box's. If they are no longer active make sure you uncheck the box."
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

Private Sub lvClassification_ItemCheck(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvClassification_ItemCheck"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    Dim xs As Integer
    
    Item.Checked = True
    
    For xs = 1 To lvClassification.ListItems.Count
        If xs <> Item.Index Then lvClassification.ListItems(xs).Checked = False
    Next
    
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

Private Sub lvClassification_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvClassification_MouseMove"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    
    If bPlay = False And picHeader(1).Visible = True Then
            
        L = GetCursorPos(oXY)
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y

        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "select the classification of the subscribers account."
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

Private Sub lvDomains_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvDomains_MouseMove"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    
    If bPlay = False And picHeader(2).Visible = True Then
            
        L = GetCursorPos(oXY)
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y

        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "This is the static list of domain's on this subscriber's account."
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

Private Sub lvEmail_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvEmail_Click"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    
    If bPlay = False And picTSContainer(1).Visible = True Then
            
        L = GetCursorPos(oXY)
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "This is where the list of the subscribers email address is keep, to stop mail being sent to the email address take the tick out of the box, remember we need at least one email address to send bills to so leave a tick in at least one box."
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

Private Sub lvEmail_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvEmail_ColumnClick"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
        ffrmEmail.sContactName = lvEmail.SelectedItem.Text
        ffrmEmail.sEmailAddress = lvEmail.SelectedItem.SubItems(1)
        ffrmEmail.Show 1
        
        If ffrmEmail.iCloseState = frmCloseSave Then
            Dim itmX As ListItem
            Set itmX = lvEmail.SelectedItem
            itmX.Text = ffrmEmail.sContactName
            itmX.SubItems(1) = ffrmEmail.sEmailAddress
            
            osub.col_subEmails(itmX.Key).ContactName = ffrmEmail.sContactName
            osub.col_subEmails(itmX.Key).EmailAddress = ffrmEmail.sEmailAddress
            
            osub.col_subEmails(itmX.Key).Key = "UPDAT" & Right(itmX.Key, Len(itmX.Key) - 5)
            itmX.Key = "UPDAT" & Right(itmX.Key, Len(itmX.Key) - 5)
            
            bEditMade = True
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
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    If Item.Key <> "" Then
        osub.col_subEmails(Item.Key).Checked = Item.Checked
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

Private Sub lvEmail_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvEmail_ItemClick"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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

Private Sub lvPayment_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    If Left(Item.Key, 3) = "NEW" Then
        osub.colPaymentSet(Item.Key).Checked = Item.Checked
    Else
        
        osub.colPaymentSet(Item.Key).Checked = Item.Checked
        osub.colPaymentSet(Item.Key).Key = "UPDAT" & Right(Item.Key, Len(Item.Key) - 5)
        Item.Key = "UPDAT" & Right(Item.Key, Len(Item.Key) - 5)
    
    End If
        
End Sub

Private Sub lvPhone_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvPhone_ColumnClick"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
        
        ffrmPhoneNo.sContactName = itmX.Text
        ffrmPhoneNo.sPhonenumber = itmX.SubItems(1)
        ffrmPhoneNo.sExtension = itmX.SubItems(2)
        ffrmPhoneNo.sNote = itmX.SubItems(3)
        ffrmPhoneNo.Show 1
        
        If ffrmPhoneNo.iCloseState = frmCloseSave Then
            
            osub.col_subPhoneNo(itmX.Key).ContactName = ffrmPhoneNo.sContactName
            osub.col_subPhoneNo(itmX.Key).PhoneNumber = ffrmPhoneNo.sPhonenumber
            osub.col_subPhoneNo(itmX.Key).Extension = ffrmPhoneNo.sExtension
            osub.col_subPhoneNo(itmX.Key).ShortNote = ffrmPhoneNo.sNote
            
            osub.col_subPhoneNo(itmX.Key).Key = "UPDAT" & Right(itmX.Key, Len(itmX.Key) - 5)
            itmX.Key = "UPDAT" & Right(itmX.Key, Len(itmX.Key) - 5)
            
            itmX.Text = ffrmPhoneNo.sContactName
            itmX.SubItems(1) = ffrmPhoneNo.sPhonenumber
            itmX.SubItems(2) = ffrmPhoneNo.sExtension
            itmX.SubItems(3) = ffrmPhoneNo.sNote
            
            bEditMade = True
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
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    If Item.Key <> "" Then
        osub.col_subPhoneNo(Item.Key).Checked = Item.Checked
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

Private Sub lvPhone_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvPhone_ItemClick"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


Private Sub lvPhone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvPhone_MouseMove"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
        
    If bPlay = False And picTSContainer(1).Visible = True Then
            
        L = GetCursorPos(oXY)
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "If the subsciber has another phone number add it in or if a number is disconnected take a tick out of the box."
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


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvPlans_ColumnClick"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    Call GUI.ColumnSort(ColumnHeader, lvPlans)
    
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

'Private Sub lvPlans_DblClick()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
 '   Const RoutineName = "lvPlans_DblClick"
 '   Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
'    If bDebug = -1 Then
'        On Error GoTo 0
'    ElseIf bDebug = 1 Then
'        On Error Resume Next
'    Else
'        On Error GoTo ErrorOccur
'    End If
'
'
'    Dim rsServices As ADODB.Recordset
'    Dim rsService As ADODB.Recordset
'    Dim rsRadius As ADODB.Recordset
'    Dim itmx As ListItem
'    Dim bResult As Boolean
'
'
'    If Not lvPlans.SelectedItem Is Nothing Then
'        bResult = MySQL.OpenTable(directConn, rsServices, , "select * from servicetypes where RecID = " & lvPlans.SelectedItem.Tag & " Limit 1")
'        Select Case rsServices!ServiceKey
'        Case "ALIAS"
'            bResult = MySQL.OpenTable(directConn, rsService, , "select AccI_RecID, ContactName, RadiusID ,RecID, Username, ServiceID, Checked, AES_DECRYPT(Password,'" & odb.colSalts.ReturnSalt("md5Password") & "') as Password, BaseURL, DynamicField1, DynamicField2, DynamicField3, DynamicField4, DynamicField5 from acci_services where RecID = " & Mid(lvPlans.SelectedItem.Key, 2) & " Limit 1")
'            Dim fAlias As New frmAlias
'            Screen.MousePointer = vbDefault
'            fAlias.sContactName = IIf(IsNull(rsService!ContactName), "", rsService!ContactName)
'            fAlias.acciRecID = IIf(IsNull(rsService!acci_RecID), 0, rsService!acci_RecID)
'            fAlias.sBaseURL = rsService!BaseURL
'            fAlias.Show 1
'            Screen.MousePointer = vbHourglass
'
'
'            MySQL.Execute directConn, "UPDATE acci_services Set BaseURL = '" & MySQL.ESC(fAlias.sBaseURL) & "', ContactName = '" & MySQL.ESC(fAlias.sContactName) & "', Username = '" & MySQL.ESC(rsService!Username) & "', DynamicField1 = '" & MySQL.ESC(fAlias.DF1) & "', DynamicField2 = '" & MySQL.ESC(fAlias.DF2) & "', DynamicField3 = '" & MySQL.ESC(fAlias.DF3) & "', DynamicField4 = '" & MySQL.ESC(fAlias.DF4) & "', DynamicField5 = '" & MySQL.ESC(fAlias.DF5) & "' where RecID = " & rsService!RecID
'
'            Set itmx = lvPlans.SelectedItem
'            itmx.Tag = rsService!ServiceID
'            itmx.Checked = rsService!Checked
'            itmx.Text = fAlias.sContactName
'            itmx.SubItems(2) = IIf(IsNull(rsService!Username), "", rsService!Username)
'            itmx.SubItems(4) = fAlias.sBaseURL
'            itmx.SubItems(5) = fAlias.DF1
'            itmx.SubItems(6) = fAlias.DF2
'            itmx.SubItems(7) = fAlias.DF3
'            itmx.SubItems(8) = fAlias.DF4
'            itmx.SubItems(9) = fAlias.DF5
'
'        Case "DIALUP", "ADSL", "SHDSL"
'            bResult = MySQL.OpenTable(directConn, rsService, , "select ContactName,RadiusID ,RecID, Username, ServiceID, Checked, AES_DECRYPT(Password,'" & odb.colSalts.ReturnSalt("md5Password") & "') as Password, BaseURL, DynamicField1, DynamicField2, DynamicField3, DynamicField4, DynamicField5 from acci_services where RecID = " & Mid(lvPlans.SelectedItem.Key, 2) & " Limit 1")
'
'            Dim fRadius As New frmRadiusAccount
'            fRadius.p_ContactName = lvPlans.SelectedItem.Text
'            fRadius.lRadiusID = rsService!RadiusID
'            fRadius.Show 1
'
'            Select Case fRadius.iCloseState
'            Case frmCloseSave
'
'                MySQL.Execute directConn, "UPDATE radiusaccounts SET Username = '" & MySQL.ESC(fRadius.p_Username) & "', ContactName = '" & MySQL.ESC(fRadius.p_ContactName) & "', SessionsAllowed = '" & fRadius.p_Sessions & ", " & _
'                                       "AutoActivateFlag = '" & fRadius.p_AutoFlag & "', Activate = '" & Format(fRadius.p_Activate, "YYYY-MM-DD TTTTT") & "', Deactivate = '" & Format(fRadius.p_Deactivate, "yyyy-mm-dd ttttt") & "', " & _
'                                       "SessionTimeout = '" & fRadius.p_SessionTimeOut & "', IdleTimeout = '" & fRadius.p_IdleTimeout & "', Checked = '" & IIf(Val(rsService!Checked) = True, "-1", "0") & "' where RecID = " & rsService!RadiusID
'
'                If Trim(fRadius.p_Password) <> "" Then
'                    MySQL.Execute directConn, "UPDATE acci_services SET Password=AES_ENCRYPT('" & fRadius.p_Password & "','" & odb.colSalts.ReturnSalt("md5Password") & "') Where RecID = " & Mid(lvPlans.SelectedItem.Key, 2)
'                    MySQL.Execute directConn, "UPDATE radiusaccounts SET Password=AES_ENCRYPT('" & fRadius.p_Password & "','" & odb.colSalts.ReturnSalt("md5Password") & "') Where RecID = " & rsService!RadiusID
'                    MySQL.Execute directConn, "update radius.radiusradcheck set Attribute='Crypt-Password', Value=encrypt('" & fRadius.p_Password & "') where Username = '" & fRadius.p_Username & "'"
'                End If
'
'                Set itmx = lvPlans.SelectedItem
'                itmx.Tag = rsService!ServiceID
'                itmx.Checked = rsService!Checked
'                itmx.Text = IIf(IsNull(rsService!ContactName), "", rsService!ContactName)
'                itmx.SubItems(2) = IIf(IsNull(rsService!Username), "", rsService!Username)
'                If fRadius.p_Password <> "" Then itmx.SubItems(3) = fRadius.p_Password
'                itmx.SubItems(4) = IIf(IsNull(rsService!BaseURL), "", rsService!BaseURL)
'                itmx.SubItems(5) = IIf(IsNull(rsService!DynamicField1), "", rsService!DynamicField1)
'                itmx.SubItems(6) = IIf(IsNull(rsService!DynamicField2), "", rsService!DynamicField2)
'                itmx.SubItems(7) = IIf(IsNull(rsService!DynamicField3), "", rsService!DynamicField3)
'                itmx.SubItems(8) = IIf(IsNull(rsService!DynamicField4), "", rsService!DynamicField4)
'                itmx.SubItems(9) = IIf(IsNull(rsService!DynamicField5), "", rsService!DynamicField5)
'
'            End Select
'        Case "FTP"
'            Dim oFTP As frmFTPAccount
'            Set oFTP = New frmFTPAccount
'            bResult = MySQL.OpenTable(directConn, rsService, , "select ContactName, RecID, Username, ServiceID, Checked, AES_DECRYPT(Password,'" & odb.colSalts.ReturnSalt("md5Password") & "') as Password, BaseURL, DynamicField1, DynamicField2, DynamicField3, DynamicField4, DynamicField5 from acci_services where RecID = " & Mid(lvPlans.SelectedItem.Key, 2) & " Limit 1")
'
'            oFTP.osub.fRecID = rsService!RecID
'            oFTP.sBaseDIR = IIf(IsNull(rsService!BaseURL), "/home/" & rsService!Username, rsService!BaseURL)
'            oFTP.sContactName = IIf(IsNull(rsService!ContactName), "", rsService!ContactName)
'            oFTP.sUsername = rsService!Username
'            'oFTP.sPassword = rsService!Password
'            oFTP.sSessions = IIf(IsNull(rsService!DynamicField1), 4, rsService!DynamicField1)
'            oFTP.byteBandwidth = IIf(IsNull(rsService!DynamicField2), "64Kbit", rsService!DynamicField2)
'            oFTP.byteBWUpload = IIf(IsNull(rsService!DynamicField3), "64Kbit", rsService!DynamicField3)
'            oFTP.Show 1
'
'            Select Case oFTP.iCloseState
'            Case frmCloseSave
'
'                MySQL.Execute directConn, "UPDATE acci_services Set BaseURL = '" & MySQL.ESC(oFTP.sBaseDIR) & "', ContactName = '" & MySQL.ESC(oFTP.sContactName) & "', Username = '" & MySQL.ESC(oFTP.sUsername) & "', DynamicField1 = '" & oFTP.sSessions & "', DynamicField2 = '" & oFTP.byteBandwidth & " ', DynamicField3 = '" & oFTP.byteBWUpload & "', DynamicField4 = '21', DynamicField5 = '21' where RecID = " & rsService!RecID
'
'                If oFTP.sPassword <> "" Then
'                    MySQL.Execute directConn, "UPDATE acci_services SET Password=AES_ENCRYPT('" & oFTP.sPassword & "','" & odb.colSalts.ReturnSalt("md5Password") & "') Where RecID = " & Mid(lvPlans.SelectedItem.Key, 2)
'                End If
'
'                Set itmx = lvPlans.SelectedItem
'                itmx.Tag = rsService!ServiceID
'                itmx.Checked = rsService!Checked
'                itmx.Text = oFTP.sContactName
'                itmx.SubItems(2) = oFTP.sUsername
'                If oFTP.sPassword <> "" Then itmx.SubItems(3) = oFTP.sPassword
'                itmx.SubItems(4) = oFTP.sBaseDIR
'                itmx.SubItems(5) = oFTP.sSessions
'                itmx.SubItems(6) = oFTP.byteBandwidth
'                itmx.SubItems(7) = oFTP.byteBWUpload
'                itmx.SubItems(8) = "21"
'                itmx.SubItems(9) = "21"
'
'            End Select
'        Case "DOMAIN"
'
'            Dim fDomain As New frmDNS
'            Dim rsDomains As ADODB.Recordset
'            Dim rsCheck As New ADODB.Recordset
'            Dim rsDNSDocs As ADODB.Recordset
'
'            bResult = MySQL.OpenTable(directConn, rsService, , "select ContactName, RecID, Username, ServiceID, Checked, AES_DECRYPT(Password,'" & odb.colSalts.ReturnSalt("md5Password") & "') as Password, BaseURL, DynamicField1, DynamicField2, DynamicField3, DynamicField4, DynamicField5, DomainID from acci_services where RecID = " & Mid(lvPlans.SelectedItem.Key, 2) & "")
'
'doDomainAgain2:
'
'            bResult = MySQL.OpenTable(directConn, rsDomains, , "select RecID, Domain, AdminEmail, AES_DECRYPT(vKey,'" & odb.colSalts.ReturnSalt("md5Password") & "') as sKey, TechName, AES_DECRYPT(TechPass,'" & odb.colSalts.ReturnSalt(PWSalt) & "') as TechPass from domainlist where RecID = " & rsService!DomainID & " Limit 1")
'
'            fDomain.sContactName = IIf(IsNull(rsService!ContactName), "", rsService!ContactName) '
'            fDomain.sDomain = IIf(IsNull(rsService!BaseURL), "", rsService!BaseURL) '
'            fDomain.sAdminEmail = IIf(IsNull(rsService!DynamicField1), "", rsService!DynamicField1) '
'            fDomain.osub.fRecID = rsService!DomainID
'            If rsDomains.State = adStateOpen Then
'                If rsDomains.RecordCount > 0 Then
'                    fDomain.sKey = IIf(IsNull(rsDomains!sKey), "", rsDomains!sKey)
'                    fDomain.sTechName = IIf(IsNull(rsDomains!TechName), "", rsDomains!TechName)
'                    fDomain.sTechPass = IIf(IsNull(rsDomains!TechPass), "", rsDomains!TechPame)
'
'
'                    bResult = MySQL.OpenTable(directConn, rsDNSDocs, , "select * from domaindocs where DomainID = " & rsService!DomainID & "")
'                    If rsDNSDocs.State = adStateOpen Then
'                        If rsDNSDocs.RecordCount > 0 Then
'                            oIP.colDNSDocs.Clear
'                            While Not rsDNSDocs.EOF
'                                oIP.colDNSDocs.Add UCase(Left(IIf(IsNull(rsDNSDocs!DocType), "XML", rsDNSDocs!DocType), 3)), rsDNSDocs!RecID, rsDNSDocs!DomainID, IIf(IsNull(rsDNSDocs!DocType), "XML", rsDNSDocs!DocType), IIf(IsNull(rsDNSDocs!DocText), "", rsDNSDocs!DocText), IIf(Val(rsDNSDocs!Icon) = 0, 1, Val(rsDNSDocs!Icon)), IIf(IsNull(rsDNSDocs!Description), "", rsDNSDocs!Description), rsDNSDocs!ItemText
'                                rsDNSDocs.MoveNext
'                            Wend
'
'                        End If
'                    End If
'                    If rsDNSDocs.State = adStateOpen Then rsDNSDocs.Close
'                End If
'            End If
'
'            'bResult = MySQL.OpenTable(directConn, rsDomains, , "select RecID, Domain, AdminEmail, AES_DECRYPT(vKey,'" & odb.colSalts.ReturnSalt("md5Password") & "') as sKey, TechName, AES_DECRYPT(TechPass,'" & odb.colSalts.ReturnSalt(PWSalt) & "') as TechPass from domainlist where RecID = " & rsService!DomainID & " Limit 1")
'            fDomain.ZOrder 0
'            fDomain.Show 1
'
'
'            If fDomain.sDomain <> rsService!DynamicField1 Then
'                bResult = MySQL.OpenTable(directConn, rsCheck, , "select * from acci_services Where ServiceID = " & rsService!ServiceID & " and DynamicField1 = '" & fDomain.sDomain & "' Limit 1")
'                If rsCheck.RecordCount > 0 Then
'                    MsgBox "Domain for this service already exist on schemer.", vbCritical, "Domain Already Exists"
'                    GoTo doDomainAgain2
'                End If
'            End If
'
'            If fDomain.iCloseState = 2 Then
'                If rsDomains.RecordCount = 0 Then
'                    Dim DomainID As Long
'
'                                        On Error Resume Next
'                    Do
'                        Err.Clear
'                        DomainID = MySQL.GetTMPRecID("domainlist", directConn)
'                        Call MySQL.Execute(directConn, "insert into domainlist (RecID,SysopID,VirtualID) VALUES ('" & DomainID & "','" & Login.lSysopID & "','" & Login.lVirtualID & "')")
'                    Loop Until Err.Number = 0
'
'
'                    MySQL.Execute directConn, "update acci_services set DomainID = '" & DomainID & "' where RecID = '" & DomainID
'                    MySQL.Execute directConn, "update domainlist set Domain = '" & fDomain.sDomain & "' where RecID = '" & DomainID
'                    MySQL.Execute directConn, "update domainlist set AdminEmail = '" & fDomain.sAdminEmail & "' where RecID = '" & DomainID
'
'                    MySQL.Execute directConn, "update domainlist set Checked = '-1' where RecID = '" & DomainID
'                    MySQL.Execute directConn, "update domainlist set TechName = '" & fDomain.sTechName & "' where RecID = '" & DomainID
'                    MySQL.Execute directConn, "update domainlist set TechPass = AES_ENCRYPT('" & fDomain.sTechPass & "','" & odb.colSalts.ReturnSalt(PWSalt) & "') where RecID = '" & DomainID
'
'                Else
'                    DomainID = rsService!DomainID
'                    MySQL.Execute directConn, "update domainlist set Domain = '" & fDomain.sDomain & "' where RecID = '" & DomainID
'                    MySQL.Execute directConn, "update domainlist set AdminEmail = '" & fDomain.sAdminEmail & "' where RecID = '" & DomainID
'
'                    MySQL.Execute directConn, "update domainlist set Checked = '-1' where RecID = '" & DomainID
'                    MySQL.Execute directConn, "update domainlist set TechName = '" & fDomain.sTechName & "' where RecID = '" & DomainID
'                    If rsDomains!TechPass <> fDomain.sTechPass Then MySQL.Execute directConn, "update domainlist set TechPass = AES_ENCRYPT('" & fDomain.sTechPass & "','" & odb.colSalts.ReturnSalt(PWSalt) & "') where RecID = '" & DomainID
'
'                End If
'
'                    If oIP.colDNSDocs.Count > 0 Then
'                        Dim Xnt As Integer
'                        For Xnt = oIP.colDNSDocs.Count To 1 Step -1
'
'                            If oIP.colDNSDocs(Xnt).RecID = 0 Then
'                                MySQL.Execute directConn, "INSERT INTO domaindocs (DomainID, DocType, DocText, Icon, Description, ItemText) " + _
'                                                       "VALUES ('" & DomainID & "','" & oIP.colDNSDocs(Xnt).DocType & "','" + _
'                                                       MySQL.ESC(oIP.colDNSDocs(Xnt).DocText) & "','" & oIP.colDNSDocs(Xnt).bIcon & "','" + _
'                                                       MySQL.ESC(oIP.colDNSDocs(Xnt).Description) & "','" & MySQL.ESC(oIP.colDNSDocs(Xnt).ItemText) & "')"
'                                oIP.colDNSDocs.Remove Xnt
'                            Else
'                                MySQL.Execute directConn, "Update domaindocs Set DocText='" & MySQL.ESC(oIP.colDNSDocs(Xnt).DocText) & "', ItemText = '" & MySQL.ESC(oIP.colDNSDocs(Xnt).ItemText) & "' where RecID = '" & oIP.colDNSDocs(Xnt).RecID & "'"
'                                oIP.colDNSDocs.Remove Xnt
'                            End If
'                        Next
'                    End If
'
'                If fDomain.sKey <> "" Then MySQL.Execute directConn, "Update domainlist Set vKey = AES_ENCRYPT('" & fDomain.sKey & "','" & odb.colSalts.ReturnSalt("md5Password") & "') Where RecID = " & DomainID
'
'                MySQL.Execute directConn, "Update acci_services Set ContactName = '" & MySQL.ESC(fDomain.sContactName) + "', BaseURL = '" & MySQL.ESC(fDomain.sDomain) + "', DynamicField1 = '" + MySQL.ESC(fDomain.sAdminEmail) + "' where RecID = " & rsServices!RecID
'
'
'                Set itmx = lvPlans.SelectedItem
'                itmx.Tag = rsService!ServiceID
'                itmx.Checked = rsService!Checked
'                itmx.Text = IIf(IsNull(rsService!ContactName), "", fDomain.sContactName)
'                itmx.SubItems(2) = IIf(IsNull(rsService!Username), "", rsService!Username)
'                itmx.SubItems(3) = IIf(IsNull(rsService!Password), "", rsService!Password)
'                itmx.SubItems(4) = IIf(IsNull(fDomain.sDomain), "", fDomain.sDomain)
'                itmx.SubItems(5) = IIf(IsNull(fDomain.sAdminEmail), "", fDomain.sAdminEmail)
'                itmx.SubItems(6) = IIf(IsNull(rsService!DynamicField2), "", rsService!DynamicField2)
'                itmx.SubItems(7) = IIf(IsNull(rsService!DynamicField3), "", rsService!DynamicField3)
'                itmx.SubItems(8) = IIf(IsNull(rsService!DynamicField4), "", rsService!DynamicField4)
'                itmx.SubItems(9) = IIf(IsNull(rsService!DynamicField5), "", rsService!DynamicField5)
'
'            End If
'
'        Case "POP3"
'
'            Dim fPOP3 As frmPOP3Account
'            Set fPOP3 = New frmPOP3Account
'            bResult = MySQL.OpenTable(directConn, rsService, , "select * from acci_services where RecID = " & Mid(lvPlans.SelectedItem.Key, 2) & " Limit 1")
'
'            If rsService.RecordCount > 0 Then
'                oService.Clear
'                Call oService.Add(lvPlans.SelectedItem.SubItems(1), Val(rsService!DynamicField3), rsService!PeriodFee, rsService!JoiningFee, rsService!PerMB, rsService!PerHour, "m", 1, rsService!ServiceKey, -1, 0, 0, 0, 0, 0, 0, 0, 0, rsService!ServiceID, 0, rsService!ptRecID, "e1")
'                fPOP3.oSvrIndex = 1
'                fPOP3.acci_acciRecID = rsService!acci_RecID
'                fPOP3.acci_ptRecID = rsService!ptRecID
'                fPOP3.acci_Activation = rsService!Activation
'                fPOP3.acci_BaseURL = rsService!BaseURL
'                fPOP3.acci_Checked = rsService!Checked
'                fPOP3.acci_ContactName = rsService!ContactName
'                fPOP3.acci_DomainID = rsService!DomainID
'                fPOP3.acci_Description = lvPlans.SelectedItem.SubItems(1)
'                fPOP3.acci_DynamicField1 = rsService!DynamicField1
'                fPOP3.acci_DynamicField3 = rsService!DynamicField3
'                fPOP3.acci_DynamicField2 = rsService!DynamicField2
'                fPOP3.acci_DynamicField4 = rsService!DynamicField4
'                fPOP3.acci_DynamicField5 = rsService!DynamicField5
'                fPOP3.acci_MBQuota = rsService!MBQuota
'                fPOP3.acci_NextCycle = rsService!NextCycle
'                fPOP3.acci_Password = rsService!Password
'                fPOP3.acci_ptRecID = rsService!ptRecID
'                fPOP3.acci_RadiusID = rsService!RadiusID
'                fPOP3.acci_RecID = rsService!RecID
'                fPOP3.acci_ServiceID = rsService!ServiceID
'                fPOP3.acci_Username = rsService!Username
'                fPOP3.lSubRecID = rsService!SubRecID
'
'                Set fPOP3.frm = Me
'
'                fPOP3.sContactName = IIf(IsNull(rsService!ContactName), "", rsService!ContactName)
'                fPOP3.sUsername = rsService!Username
'                'fPOP3.sPassword = rsService!Password
'                fPOP3.acci_acciRecID = rsService!acci_RecID
'                fPOP3.acci_ptRecID = rsService!ptRecID
'
'                fPOP3.sDomain = IIf(IsNull(rsService!BaseURL), "ep.net.au", rsService!BaseURL)
'
'                fPOP3.Show 1
'
'            End If
'
'            Case Else
'
'        End Select
'
'    End If
'
'    Screen.MousePointer = vbDefault
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

Private Sub lvPlans_ItemCheck(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvPlans_ItemCheck"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    If Item.Key <> "" Then Item.Key = "x" & Mid(Item.Key, 2)
    
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

Private Sub lvPlans_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvPlans_ItemClick"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    cmdChangePlan.Enabled = True
    cmdChangePlan.Tag = Item.Key
    
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

Private Sub lvPlans_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvPlans_MouseDown"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    If Button = 2 Then
        If Not lvPlans.SelectedItem Is Nothing Then
                
            Set frmMDIMain.frmCust = Me
            PopupMenu frmMDIMain.mnuCust_lvPlans
            
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

Private Sub lvPlans_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvPlans_MouseMove"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    
    If bPlay = False And picTSContainer(3).Visible = True Then
            
        L = GetCursorPos(oXY)
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "This is where the list of service and plans that the account is subscribed to. To alter them you can right mouse click on the options and change the details."
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

Private Sub lvReferedBy_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvReferedBy_ColumnClick"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    Call GUI.ColumnSort(ColumnHeader, lvReferedBy)
    
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

Private Sub lvReferedBy_DblClick()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvReferedBy_DblClick"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    If lvReferedBy.Tag <> "" Then
        
        Dim ffrmShortNote As frmShortNote
        Dim itmX As ListItem
        Set ffrmShortNote = New frmShortNote
        Set itmX = lvReferedBy.SelectedItem
        
        ffrmShortNote.sShortNote = itmX.SubItems(2)
        ffrmShortNote.Show 1
        
        If ffrmShortNote.iCloseState = frmCloseSave Then
            itmX.SubItems(2) = ffrmShortNote.sShortNote
            Dim oRef As cls_subReferals
            
            For Each oRef In osub.col_subReferals
                If oRef.IDX = itmX.Tag Then
                    oRef.ShortNote = itmX.SubItems(2)
                    Exit For
                End If
            Next
            bEditMade = True
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

Private Sub lvReferedBy_ItemCheck(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvReferedBy_ItemCheck"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    If Item.Tag <> 0 Then
        Item.Checked = True
    End If
    Dim facciServiceID As New frmACCI_Services
    Dim xs As Integer
    
    If Left(Item.Key, 1) = "a" Then
        If Item.Checked = True Then
            
            Set facciServiceID.osub = Me.osub
            facciServiceID.Show 1
            
            With osub.col_subReferals.Add("NEW" & osub.col_subReferals.Count + 1, 0, 0, Val(Mid(Item.Key, 2)), 0, sysnow, Item.SubItems(1), "", 0, Item.Checked, facciServiceID.acciServiceID, "NEW" & osub.col_subReferals.Count + 1)
                .AccountName = Item.Text
                Item.Key = .Key
                Item.Tag = .IDX
            End With
            
        End If
    ElseIf Left(Item.Key, 3) = "NEW" Then
    
        
        Set facciServiceID.osub = Me.osub
        facciServiceID.Show 1
        
        With osub.col_subReferals(Item.Key)
            .acciServiceID = facciServiceID.acciServiceID
            Item.Tag = .IDX
        End With
    Else
    
        With osub.col_subReferals(Item.Key)
            .Checked = Item.Checked
            Item.Tag = .IDX
        End With
        
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

Private Sub lvReferedBy_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvReferedBy_ItemClick"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    lvReferedBy.Tag = True
    
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



Private Sub lvReferedBy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvReferedBy_MouseMove"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    
    If bPlay = False And picTSContainer(2).Visible = True Then
            
        L = GetCursorPos(oXY)
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "A list of people you have search for to select as a referee to the account holder/subcriber."
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

Private Sub lvTransactions_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvTransactions_ColumnClick"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    Call GUI.ColumnSort(ColumnHeader, lvTransactions)
    
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

Private Sub lvTransactions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvTransactions_MouseMove"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    
    If bPlay = False And picTSContainer(2).Visible = True Then
            
        L = GetCursorPos(oXY)
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "A list of the transaction that the account name subscriber has occured in upkeep."
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

Private Sub mvBilling_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mvBilling_MouseMove"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    
    If bPlay = False And picTSContainer(2).Visible = True Then
            
        L = GetCursorPos(oXY)
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "This is where you can adjust the next billing date, if you set it in the past they will be billed immediately next time upkeep is done."
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

Private Sub optBillingCycle_Click(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "optBillingCycle_Click"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    
    For bx = txtBillingCycle.LBound To txtBillingCycle.UBound
        If bx = Index Then
            txtBillingCycle(bx).Enabled = True
        Else
            txtBillingCycle(bx).Enabled = False
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

Private Sub picBut_Resize(Index As Integer)

    Select Case Index
    Case 3
        'dftMain.Move 0, 0, picBut(Index).ScaleWidth, picBut(Index).ScaleHeight
    End Select
    
End Sub

Private Sub picTSContainer_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "picTSContainer_MouseMove"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
'
'    If bPlay = False And picTSContainer(5).Visible = True Then
'
'        L = GetCursorPos(oXY)
'        if not frmAgent.oChar is nothing then frmAgent.oChar.StopAll
'        if not frmAgent.oChar is nothing then frmAgent.oChar.GestureAt oXY.X, oXY.Y
'        if not frmAgent.oChar is nothing then frmAgent.oChar.Speak "This is where the support information for hardware and system being connect to the subscribers account, this help diagnostics when analysing an error or system malfunction."
'        bPlay = True
'
'    End If
    
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

Private Sub tb_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim bbut As MSComctlLib.Button
    Dim Counter As Byte
    Dim counter2 As Byte
    
    For Each bbut In tb.Buttons
        If bbut.Index <> Button.Index Then
            If Not bbut.Value = tbrUnpressed Then bbut.Value = tbrUnpressed
        End If
    Next

    For Counter = 0 To tb.Buttons.Count
        For counter2 = 0 To picBut.UBound
            
            If picBut(counter2).Tag = Button.Tag Then
                picBut(counter2).Move 4, 94, 639, 513
                picBut(counter2).ZOrder 0
                picBut(counter2).Visible = True
            Else
                picBut(counter2).Visible = False
            End If
        Next
    Next
    
End Sub

Private Sub Timer1_Timer()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Timer1_Timer"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


   If Not lblCustomerCreated.Caption = Timer1.Tag & Format(sysnow, "dd/mmm/yyyy Hh:Nn:Ss") Then lblCustomerCreated.Caption = Timer1.Tag & Format(sysnow, "dd/mmm/yyyy Hh:Nn:Ss")
   
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

Private Sub tsDetails_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "tsDetails_Click"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
        If tsDetails.SelectedItem.Index - 1 <> X Then picTSContainer(X).Visible = False
    Next
    
    If tsDetails.SelectedItem.Index - 1 <= picTSContainer.UBound Then
        picTSContainer(tsDetails.SelectedItem.Index - 1).Move tsDetails.ClientLeft, tsDetails.ClientTop, tsDetails.ClientWidth, tsDetails.ClientHeight
        picTSContainer(tsDetails.SelectedItem.Index - 1).Visible = True
        picTSContainer(tsDetails.SelectedItem.Index - 1).ZOrder 0
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

Public Function LoadAccountTypes()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "LoadAccountTypes"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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



    rs_AccountType.MoveFirst
    If rs_AccountType.RecordCount > 0 Then
        Dim itmX As ListItem
        While Not rs_AccountType.EOF And Err.Number = 0
            Set itmX = lvAccountTypes.ListItems.Add(, "r" & rs_AccountType!RecID, rs_AccountType!AccountDesc)
            rs_AccountType.MoveNext
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

Public Function SaveInformation() As Variant


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "SaveInformation"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    
    
    
    'If DateDiff("s", Format(dtpActivation.Value, "HH:MM:SS"), Format(dtpDeactivation.Value, "HH:MM:SS")) < 0 Then
    '    MsgBox "Activation Time must be less than deactivation time in the Auto-activation section of Session Details.", vbInformation, "Time Variance Detected"
    '    Exit Function
    'End If
    
    Me.Hide
    
    If Not osub.fSysopID = IIf(cmbSysopID.ListIndex = -1, Login.lSysopID, cmbSysopID.ItemData(cmbSysopID.ListIndex)) Then
        osub.fSysopID = IIf(cmbSysopID.ListIndex = -1, Login.lSysopID, cmbSysopID.ItemData(cmbSysopID.ListIndex))
        If osub.fRecID <> 0 Then Call AddAudit(Client, vLong, "SysopID", Me, , , , , IIf(cmbSysopID.ListIndex = -1, Login.lSysopID, cmbSysopID.ItemData(cmbSysopID.ListIndex)), osub.fSysopID, "Sysop Ownership of Customer Record", osub.fRecID)
    End If
        
    If Not osub.fVirtualID = IIf(cmbVirtualID.ListIndex = -1, Login.lVirtualID, cmbVirtualID.ItemData(cmbVirtualID.ListIndex)) Then
        osub.fVirtualID = IIf(cmbVirtualID.ListIndex = -1, Login.lVirtualID, cmbVirtualID.ItemData(cmbVirtualID.ListIndex))
        If osub.fRecID <> 0 Then Call AddAudit(Client, vLong, "VirtualID", Me, , , , , IIf(cmbVirtualID.ListIndex = -1, Login.lVirtualID, cmbVirtualID.ItemData(cmbVirtualID.ListIndex)), osub.fVirtualID, "VISP Ownership of Customer Record", osub.fRecID)
    End If
    
    If Not osub.fAgencyID = Login.lAgencyID Then
        osub.fAgencyID = Login.lAgencyID
        If osub.fRecID <> 0 Then Call AddAudit(Client, vLong, "AgencyID", Me, , , , , osub.fAgencyID, Login.lAgencyID, "Agency Ownership of Customer Record", osub.fRecID)
    End If
   
    If Not osub.fAccountName = txtAccountName.Text Then
        osub.fAccountName = txtAccountName.Text
        If osub.fRecID <> 0 Then Call AddAudit(Client, vLong, "AccountName", Me, , , , , txtAccountName.Text, osub.fAccountName, "Account Name", osub.fRecID)
    End If
    
    If Not osub.fDOB = dtpDOB.Value Then
        osub.fDOB = dtpDOB.Value
        If osub.fRecID <> 0 Then Call AddAudit(Client, vLong, "DOB", Me, , , Format(dtpDOB.Value, "dd-mmm-yyyy ttttt"), Format(osub.fDOB, "dd-mmm-yyyy ttttt"), , , "Date of Registration or Birth", osub.fRecID)
    End If
    
    If Not osub.fCancelled = Val(chkCancelled.Value) Then
        osub.fCancelled = Val(chkCancelled.Value)
        If osub.fRecID <> 0 Then Call AddAudit(Client, vLong, "Cancelled", Me, , , , , Val(chkCancelled.Value), osub.fCancelled, "Cancelled Flag", osub.fRecID)
    End If
    
    
    osub.fProcessFlag = 1
    
    

    If Login.lLevel > 90 Or osub.fRecID = 0 Then
        If Not osub.fBillingDate = CDate(Format(mvBilling.Value, "yyyy-mm-dd ") + Format(dtpBilling.Value, "Hh:Nn:Ss")) Then
            osub.fBillingDate = CDate(Format(mvBilling.Value, "yyyy-mm-dd ") + Format(dtpBilling.Value, "Hh:Nn:Ss"))
            If osub.fRecID <> 0 Then Call AddAudit(Client, vLong, "DOB", Me, , , Format(osub.fBillingDate, "dd-mmm-yyyy ttttt"), Format(osub.fBillingDate, "dd-mmm-yyyy ttttt"), , , "Next Billing Date", osub.fRecID)
        End If
    End If
    
    If Not osub.fFlagA_RecID = cmbFlagA.ItemData(cmbFlagA.ListIndex) Then
        osub.fFlagA_RecID = cmbFlagA.ItemData(cmbFlagA.ListIndex)
        osub.fFlagASet = sysnow
    End If
    
    If Not osub.fFlagB_RecID = cmbFlagB.ItemData(cmbFlagB.ListIndex) Then
        osub.fFlagB_RecID = cmbFlagB.ItemData(cmbFlagB.ListIndex)
        osub.fFlagBSet = sysnow
    End If
       
    Dim bx As Byte
    
    For bx = optHearAbout.LBound To optHearAbout.UBound
        If optHearAbout(bx).Value = True Then
            osub.fAboutUs = bx
            Exit For
        End If
    Next
        
    For bx = txtBillingCycle.LBound To txtBillingCycle.UBound
        If txtBillingCycle(bx).Enabled = True Then
            osub.fPayIntervalType = optBillingCycle(bx).Tag
            osub.fPayInterval = Val(IIf(txtBillingCycle(bx).Text = "", "1", txtBillingCycle(bx).Text))
            Exit For
        End If
    Next
    
    Dim tx As Integer
    For tx = 1 To lvClassification.ListItems.Count
        If lvClassification.ListItems(tx).Checked = True Then
            osub.fClassification = Val(Mid(lvClassification.ListItems(tx).Key, 2))
            Exit For
        End If
    Next
    
    If txtgPWD.Text <> "" Then
        osub.fgPassword = txtgPWD.Text
        osub.fgUsername = txtgUID.Text
    End If
    
    Call osub.Commit(directConn)
    SaveInformation = osub.fRecID
    
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

'Public Sub SaveReferals(fRecID As Long)
'
'
'    '*[ Error Checking Variables ]**********************************************************************************
'
'
'    Const RoutineName = "SaveReferals"
'    Const ContainerName = "frmCustomerRec"
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
'    Dim itmX As ListItem
'    Dim sa As Integer
'    Dim rsSave As ADODB.Recordset
'    Dim acci_RecID2 As Long
'    If lvReferedBy.ListItems.Count > 0 Then
'
'        For sa = 1 To lvReferedBy.ListItems.Count
'            Set itmX = lvReferedBy.ListItems(sa)
'            If Left(itmX.Key, 1) <> "z" Then
'            If Left(itmX.Key, 1) = "x" Then
'
'            If itmX.Checked = True Then
'
'                If InStr(itmX.Key, "_") > 0 Then
'                    acci_RecID2 = CLng(Mid(itmX.Key, InStr(itmX.Key, "-") + 1))
'                Else
'                    acci_RecID2 = CLng(Mid(itmX.Key, 2))
'                End If
'
'                MySQL.Execute directConn, "INSERT INTO acci_referedby (acci_RecID, acci_RecID2, ContactName, ShortNote, Checked, acciServiceID) " & _
'                                        "VALUES('" & osub.fRecID& "','" & acci_RecID2 & "','" & MySQL.ESC(itmX.SubItems(1)) & "','" & MySQL.ESC(itmX.SubItems(2)) & "','" & IIf(itmX.Checked = True, -1, 0) & "','" & itmX.Tag & "')"
'            End If
'
'            ElseIf Left(itmX.Key, 1) = "j" Then
'
'                MySQL.Execute directConn, "UPDATE acci_referedby SET ContactName = '" & MySQL.ESC(itmX.SubItems(1)) & "', ShortNote = '" & MySQL.ESC(itmX.SubItems(2)) & "', Checked = '" & IIf(itmX.Checked = True, -1, 0) & "', acciServiceID = '" & itmX.Tag & "' Where RecID = " & Mid(itmX.Key, 2)
'
'            End If
'            End If
'        Next
'
'    End If
'
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
'Public Sub SaveAddresses(osub.fRecIDAs Variant)
'
'
'    '*[ Error Checking Variables ]**********************************************************************************
'
'
'    Const RoutineName = "SaveAddresses"
'    Const ContainerName = "frmCustomerRec"
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
'    Dim itmX As ListItem
'    Dim sa As Integer
'    Dim rsSave As ADODB.Recordset
'
'    If lvAddresses.ListItems.Count > 0 Then
'
'        For sa = 1 To lvAddresses.ListItems.Count
'            Set itmX = lvAddresses.ListItems(sa)
'            If itmX.Key = "" Then
'
'                MySQL.Execute directConn, "INSERT INTO acci_addresses (acci_RecID, ContactName, Street1, Street2, Suburb, State, PostCode, Country, Checked) VALUES ('" & osub.fRecID& "','" & MySQL.ESC(itmX.Text) & "','" & MySQL.ESC(itmX.SubItems(1)) & "','" & MySQL.ESC(itmX.SubItems(2)) & "','" & MySQL.ESC(itmX.SubItems(3)) & "','" & MySQL.ESC(itmX.SubItems(4)) & "','" & MySQL.ESC(itmX.SubItems(5)) & "','" & MySQL.ESC(itmX.SubItems(6)) & "','" & IIf(itmX.Checked = True, "-1", "0") & "')"
'
'            ElseIf Left(itmX.Key, 1) = "e" Then
'
'                MySQL.Execute directConn, "UPDATE acci_addresses set ContactName = '" & MySQL.ESC(itmX.Text) + "', Street1 = '" & MySQL.ESC(itmX.SubItems(1)) & "', Street2 = '" & MySQL.ESC(itmX.SubItems(2)) & "', Suburb = '" & MySQL.ESC(itmX.SubItems(3)) & "', State = '" & MySQL.ESC(itmX.SubItems(4)) & "', PostCode = '" & itmX.SubItems(5) & "', Country = '" & MySQL.ESC(itmX.SubItems(6)) & "', Checked = '" & IIf(itmX.Checked = True, "-1", "0") & "' where RecID = " & Mid(itmX.Key, 2)
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
'Public Sub SavePhoneNumbers(osub.fRecIDAs Variant)
'
'
'    '*[ Error Checking Variables ]**********************************************************************************
'
'
'    Const RoutineName = "SavePhoneNumbers"
'    Const ContainerName = "frmCustomerRec"
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
'    Dim itmX As ListItem
'    Dim sa As Integer
'    Dim rsSave As ADODB.Recordset
'
'    If lvPhone.ListItems.Count > 0 Then
'
'        For sa = 1 To lvPhone.ListItems.Count
'            Set itmX = lvPhone.ListItems(sa)
'            If itmX.Key = "" Then
'
'                MySQL.Execute directConn, "INSERT INTO acci_phonenumbers (acci_RecID, ContactName, PhoneNumber, Extension, ShortNote, Checked) " & _
'                                        "VALUES('" & osub.fRecID& "','" & MySQL.ESC(itmX.Text) & "','" & MySQL.ESC(itmX.SubItems(1)) & "','" & MySQL.ESC(itmX.SubItems(2)) & "','" & MySQL.ESC(itmX.SubItems(3)) & "','" & IIf(itmX.Checked = True, -1, 0) & "')"
'
'            ElseIf Left(itmX.Key, 1) = "e" Then
'
'                MySQL.Execute directConn, "UPDATE acci_phonenumbers SET ContactName = '" & MySQL.ESC(itmX.Text) & "', PhoneNumber= '" & MySQL.ESC(itmX.SubItems(1)) & "', Extension = '" & MySQL.ESC(itmX.SubItems(2)) & "', ShortNote = '" & MySQL.ESC(itmX.SubItems(3)) & ", Checked = '" & IIf(itmX.Checked = True, -1, 0) & " Where RecID = " & Mid(itmX.Key, 2)
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
'Public Sub SaveEmail(osub.fRecIDAs Variant)
'
'
'    '*[ Error Checking Variables ]**********************************************************************************
'
'
'    Const RoutineName = "SaveEmail"
'    Const ContainerName = "frmCustomerRec"
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
'    Dim itmX As ListItem
'    Dim sa As Integer
'    Dim rsSave As ADODB.Recordset
'
'    If lvEmail.ListItems.Count > 0 Then
'
'        For sa = 1 To lvEmail.ListItems.Count
'            Set itmX = lvEmail.ListItems(sa)
'            If itmX.Key = "" Then
'
'                MySQL.Execute directConn, "INSERT INTO acci_emailaddresses (acci_RecID, ContactName, Emailaddress, Checked) " + _
'                                       "VALUES('" & osub.fRecID& "','" & MySQL.ESC(itmX.Text) & "',AES_ENCRYPT('" & MySQL.ESC(itmX.SubItems(1)) & "','" & odb.colSalts.ReturnSalt(EMAILSalt) & "'),'" & IIf(itmX.Checked = True, -1, 0) & "')"
'
'            ElseIf Left(itmX.Key, 1) = "e" Then
'
'                MySQL.Execute directConn, "UPDATE acci_emailaddresses SET ContactName = '" & itmX.Text & "', Emailaddress=AES_ENCRYPT('" & MySQL.ESC(itmX.SubItems(1)) & "','" & odb.colSalts.ReturnSalt(EMAILSalt) & "'), acci_RecID=" & "'" & osub.fRecID& "', Checked='" & IIf(itmX.Checked = True, -1, 0) & "' Where RecID = " & Mid(itmX.Key, 2)
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
'Public Function SaveFTPAccounts(osub.fRecIDAs Variant)
'
'
'    '*[ Error Checking Variables ]**********************************************************************************
'
'
'    Const RoutineName = "SaveFTPAccounts"
'    Const ContainerName = "frmCustomerRec"
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
'    Dim itmX As ListItem
'    Dim sa As Integer
'    Dim rsSave As ADODB.Recordset
'
'    If lvFTPAccounts.ListItems.Count > 0 Then
'
'        For sa = 1 To lvFTPAccounts.ListItems.Count
'            Set itmX = lvFTPAccounts.ListItems(sa)
'            If itmX.Key = "" Then
'
'                MySQL.Execute directConn, "INSERT INTO acci_ftpaccounts (acci_RecID, ContactName, Username, Password, BaseDir, Checked) " & _
'                                        "VALUES('" & osub.fRecID& "','" & MySQL.ESC(itmX.Text) & "','" & MySQL.ESC(itmX.SubItems(1)) & "','" & MySQL.ESC(itmX.SubItems(2)) & "','" & MySQL.ESC(itmX.SubItems(3)) & "','" & IIf(itmX.Checked = True, -1, 0) & "')"
'
'            ElseIf Left(itmX.Key, 1) = "e" Then
'
'                MySQL.Execute directConn, "UPDATE acci_ftpaccounts SET ContactName = '" & MySQL.ESC(itmX.Text) & "', Username = '" & MySQL.ESC(itmX.SubItems(1)) & "', Password = '" & MySQL.ESC(itmX.SubItems(2)) & "', BaseDir = '" & MySQL.ESC(itmX.SubItems(3)) & ", Checked = '" & IIf(itmX.Checked = True, -1, 0) & " Where RecID = " & Mid(itmX.Key, 2)
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
'Public Function SavePOP3Accounts(osub.fRecIDAs Variant)
'
'
'    '*[ Error Checking Variables ]**********************************************************************************
'
'
'    Const RoutineName = "SavePOP3Accounts"
'    Const ContainerName = "frmCustomerRec"
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
'    Dim itmX As ListItem
'    Dim sa As Integer
'    Dim rsSave As ADODB.Recordset
'
'    If lvPOP3.ListItems.Count > 0 Then
'
'        For sa = 1 To lvPOP3.ListItems.Count
'            Set itmX = lvPOP3.ListItems(sa)
'            If itmX.Key = "" Then
'
'                MySQL.Execute directConn, "INSERT INTO acci_pop3accounts (acci_RecID, ContactName, Username, Password, Checked) " & _
'                            "VALUES('" & osub.fRecID& "','" & MySQL.ESC(itmX.Text) & "','" & MySQL.ESC(itmX.SubItems(1)) & "','" & MySQL.ESC(itmX.SubItems(2)) & "','" & IIf(itmX.Checked = True, "-1", "0") & "')"
'
'            ElseIf Left(itmX.Key, 1) = "e" Then
'
'                MySQL.Execute directConn, "UPDATE acci_pop3accounts SET ContactName = '" & MySQL.ESC(itmX.Text) & "', Username = '" & MySQL.ESC(itmX.SubItems(1)) & "', Password = '" & MySQL.ESC(itmX.SubItems(2)) & "', Checked = '" & IIf(itmX.Checked = True, -1, 0) & "' Where RecID = " & Mid(itmX.Key, 2)
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

Public Function LoadInformation()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "LoadInformation"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    
    Const lBytesPerMB = 1024 ^ 2
    
    txtAccountName.Text = osub.fAccountName
    txtgUID.Text = osub.fgUsername
    optHearAbout(osub.fAboutUs).Value = True
    chkCancelled.Value = osub.fCancelled
    dtpDOB.Value = osub.fDOB
    mvBilling.Value = osub.fBillingDate
    dtpBilling.Value = osub.fBillingDate
    
    Dim iCount As Long
    
    For iCount = 0 To cmbVirtualID.ListCount - 1
        If cmbVirtualID.ItemData(iCount) = osub.fVirtualID Then
            cmbVirtualID.ListIndex = iCount
        End If
    Next
    
    cmbVirtualID.Locked = IIf(Login.bOwnership = True, False, True)
    
    For iCount = 0 To cmbSysopID.ListCount - 1
        If cmbSysopID.ItemData(iCount) = osub.fSysopID Then
            cmbSysopID.ListIndex = iCount
        End If
    Next
    
    cmbSysopID.Locked = IIf(Login.bOwnership = True, False, True)
        
    txtData(0).Text = Format((osub.fsfCycle_Download), "###,###,###,###,###,###,###,###,###,###,###,### bytes")
    txtData(1).Text = Format((osub.fsfCycle_Download) / lBytesPerMB, "###,###,###,###,###,###,###,###,###,###,###,###.## MB")
    txtData(2).Text = Format(osub.fsfCycle_Upload, "###,###,###,###,###,###,###,###,###,###,###,### bytes")
    txtData(3).Text = Format(osub.fsfCycle_Upload / lBytesPerMB, "###,###,###,###,###,###,###,###,###,###,###,###.## MB")
    
    Dim bx As Byte
    
    'For bx = chkAppetite.LBound To chkAppetite.UBound
    '    chkAppetite(bx).Value = rsLoad(chkAppetite(bx).Tag)
    'Next

    Set lvClassification.SelectedItem = lvClassification.ListItems("k" & osub.fClassification)
    lvClassification.SelectedItem.Checked = True
    
    Dim ix As Variant
    For ix = 0 To cmbFlagA.ListCount - 1
        
        If cmbFlagA.ItemData(ix) = osub.fFlagA_RecID Then
            cmbFlagA.ListIndex = ix
            cmbFlagA.ToolTipText = "This flag was changed last on the " & Format(osub.fFlagASet, "dd mmm yyyy ttttt")
        End If
        
    Next ix
    
    For ix = 0 To cmbFlagB.ListCount - 1
        If cmbFlagB.ItemData(ix) = osub.fFlagB_RecID Then
            cmbFlagB.ListIndex = ix
            cmbFlagB.ToolTipText = "This flag was changed last on the " & Format(osub.fFlagBSet, "dd mmm yyyy ttttt")
        End If
    Next ix
    
   
    Dim oTrn As cls_subTrans
    If osub.col_subTrans.Count > 0 Then
        For Each oTrn In osub.col_subTrans
            Set itmX = lvTransactions.ListItems.Add(, oTrn.Key, oTrn.Description) ')
            itmX.SubItems(1) = Format(oTrn.TotalDue, "Currency")
            itmX.SubItems(2) = Format(oTrn.GSTCharged, "Currency")
            itmX.SubItems(3) = Format(oTrn.PaymentDue, "dddd dd-mm-yyyy Hh:Nn:Ss")
            itmX.SubItems(4) = Format(oTrn.AmountPaid, "Currency")
            If oTrn.AmountPaid + oTrn.AmountRefunded >= oTrn.TotalDue Then
                itmX.SubItems(5) = "Yes"
                itmX.ForeColor = RGB(0, 0, 255)
            ElseIf oTrn.AmountPaid > 10 Then
                itmX.SubItems(5) = "Partly"
                itmX.ForeColor = RGB(30, 128, 150)
            Else
                itmX.SubItems(5) = "no"
                itmX.ForeColor = RGB(234, 10, 15)
            End If
            itmX.Checked = oTrn.Checked
        Next
    End If
    
    Dim oSNL As clsSnailMail
    If osub.colSnailMail.Count > 0 Then
        For Each oSNL In osub.colSnailMail
            Set itmX = lvAddresses.ListItems.Add(1, oSNL.Key, oSNL.ContactName) '
            itmX.SubItems(1) = oSNL.Street1
            itmX.SubItems(2) = oSNL.Street2
            itmX.SubItems(3) = oSNL.Suburb
            itmX.SubItems(4) = oSNL.State
            itmX.SubItems(5) = oSNL.PostCode
            itmX.SubItems(6) = oSNL.Country
            itmX.Checked = oSNL.Checked
        Next
    End If
    
    Dim oPhn As cls_subPhoneNo
    If osub.col_subPhoneNo.Count > 0 Then
        For Each oPhn In osub.col_subPhoneNo
            Set itmX = lvPhone.ListItems.Add(, oPhn.Key, oPhn.ContactName)
            itmX.SubItems(1) = oPhn.PhoneNumber
            itmX.SubItems(2) = oPhn.Extension
            itmX.SubItems(3) = oPhn.ShortNote
            itmX.Checked = oPhn.Checked
        Next
    End If
    
    Dim oEml As cls_subEmails
    If osub.col_subEmails.Count > 0 Then
        For Each oEml In osub.col_subEmails
            Set itmX = lvEmail.ListItems.Add(, oEml.Key, oEml.ContactName)
            itmX.SubItems(1) = oEml.EmailAddress
        Next
    End If
    
    Dim oPln As cls_subServices
    If osub.col_subServices.Count > 0 Then
        For Each oPln In osub.col_subServices
                Set itmX = lvPlans.ListItems.Add(, oPln.Key, oPln.ContactName)
                itmX.Tag = oPln.ServiceID
                itmX.SubItems(1) = oPln.Description
                itmX.SubItems(2) = oPln.Username
                itmX.SubItems(3) = Left(oPln.Password, 2) + Right(oPln.Password, 2)
                itmX.SubItems(4) = oPln.BaseURL
                itmX.SubItems(5) = oPln.DynamicField1
                itmX.SubItems(6) = oPln.DynamicField2
                itmX.SubItems(7) = oPln.DynamicField3
                itmX.SubItems(8) = oPln.DynamicField4
                itmX.SubItems(9) = oPln.DynamicField5
                itmX.Checked = oPln.Checked
                If oPln.ActivationSet = 0 Then
                    itmX.ForeColor = RGB(255, 0, 0)
                Else
                    If oPln.Checked = False Then
                        itmX.ForeColor = RGB(0, 255, 0)
                    Else
                        itmX.ForeColor = RGB(0, 0, 255)
                    End If
                End If
        Next
    End If
    
    
    Dim oEdt As cls_subEditLog
    If osub.col_subEditLog.Count > 0 Then
        For Each oEdt In osub.col_subEditLog
            Set itmX = lvEditLog.ListItems.Add(, oEdt.Key, Format(oEdt.DateEditMade, "dddd, dd-mmm-yyyy ttttt"))
            itmX.SubItems(1) = oEdt.Username
            itmX.SubItems(2) = oEdt.EditTxt
            
            If oEdt.SysopID = Login.lSysopID Then
                itmX.ForeColor = RGB(128, 0, 128)
            Else
                itmX.ForeColor = RGB(128, 128, 128)
            End If
        Next
    End If
    
    

    Dim oPay As clsPaymentSet
    If osub.colPaymentSet.Count > 0 Then
        lvPayment.ListItems.Clear
        For Each oPay In osub.colPaymentSet
            If Len(oPay.ccCardName) > 0 Then
            
                Set itmX = lvPayment.ListItems.Add(, oPay.Key, "Credit Card")
                itmX.SubItems(1) = "100%"
                itmX.SubItems(2) = "Card Details: " & oPay.ccCardNumber & " - " & oPay.ccCardName & " - " & oPay.ccCardExpiry
                itmX.Tag = oPay.IDX
                itmX.Checked = oPay.Checked
                
            ElseIf Len(oPay.ddAccountName) > 0 Then
            
                Set itmX = lvPayment.ListItems.Add(, oPay.Key, "Direct Debit")
                itmX.SubItems(1) = "100%"
                itmX.SubItems(2) = "Direct Debit Details: " & oPay.ddAccountName & " - " & oPay.ddAcountNo & " - " & oPay.ddBSB
                itmX.Tag = oPay.IDX
                itmX.Checked = oPay.Checked
            Else
            
                Set itmX = lvPayment.ListItems.Add(, oPay.Key, "Statement Watcher")
                itmX.SubItems(1) = "100%"
                itmX.SubItems(2) = "Watcher Details: " & oPay.swNumber & " - " & oPay.swWord
                itmX.Tag = oPay.IDX
                itmX.Checked = oPay.Checked
                
            End If
            
        Next
    End If
    
    Dim oRff As cls_subReferals
    If osub.col_subReferals.Count > 0 Then
        For Each oRff In osub.col_subReferals
            Set itmX = lvReferedBy.ListItems.Add(, oRff.Key, oRff.AccountName)
            itmX.SubItems(1) = oRff.ContactName
            itmX.SubItems(2) = oRff.ShortNote
            itmX.Tag = oRff.IDX
            itmX.Checked = oRff.Checked
            rsload.MoveNext
            
            If oEdt.SysopID = Login.lSysopID Then
                itmX.ForeColor = RGB(128, 0, 128)
            Else
                itmX.ForeColor = RGB(128, 128, 128)
            End If
        Next
    End If
    
    
'    bResult = MySQL.OpenTable(directConn, rsLoad, , "select * from domainlist Where acci_RecID = " & oSub.fRecID)
'
'    If bResult = True Then
'        If rsLoad.RecordCount > 0 Then
'            rsLoad.MoveFirst
'            While Not rsLoad.EOF And Err.Number = 0
'                Set itmX = lvDomains.ListItems.Add(, "l" & oSUB.fRecID, oSUB.fDomain)
'                itmX.SubItems(1) = IIf(IsNull(oSUB.fAdminEmail), "", oSUB.fAdminEmai) 'oSUB.fContactName
'                rsLoad.MoveNext
'            Wend
'        End If
'    End If
'
'    Dim ib As Byte
    
    'bResult = MySQL.OpenTable(directConn, rsLoad, , "select * from acci_hardware Where acci_RecID = " & oSub.fRecID)
     
    'If bResult = True Then
    '    For ib = cmbField.LBound To cmbField.UBound
    '        If Not IsNull(rsLoad(cmbField(ib).Tag)) Then cmbField(ib).Text = rsLoad(cmbField(ib).Tag)
    '    Next
    'End If
    
    DrawDatawave picDatawave, CLng(osub.fRecID)
    
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

Private Sub tsHeader_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "tsHeader_Click"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    
    For X = picHeader.LBound To picHeader.UBound
        If tsHeader.SelectedItem.Index - 1 <> X Then picHeader(X).Visible = False
    Next
    
    If tsHeader.SelectedItem.Index - 1 <= picHeader.UBound Then
        picHeader(tsHeader.SelectedItem.Index - 1).Move tsHeader.ClientLeft, tsHeader.ClientTop, tsHeader.ClientWidth, tsHeader.ClientHeight
        picHeader(tsHeader.SelectedItem.Index - 1).Visible = True
        picHeader(tsHeader.SelectedItem.Index - 1).ZOrder 0
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


Private Sub txtAccountName_Change()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtAccountName_Change"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    If osub.fRecID = 0 And Len(Trim(txtAccountName.Text)) > 0 And Len(Trim(txtgUID.Text)) > 0 And Len(Trim(txtgPWD.Text)) > 0 Then
        cmdAddAddress.Enabled = True
        cmdAddEmail.Enabled = True
        cmdAddPhone.Enabled = True
        cmdAddService.Enabled = True
        cmdChangePlan.Enabled = True
        cmdBill.Enabled = True
        tsDetails.Enabled = True
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

Private Sub txtAccountName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtAccountName_MouseMove"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    
    If bPlay = False And picHeader(0).Visible = True Then
            
        L = GetCursorPos(oXY)
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "This is where you enter the name of the account for the services and details to reside in."
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

Private Sub txtBillingCycle_KeyPress(Index As Integer, KeyAscii As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtBillingCycle_KeyPress"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    Case Asc("0") To Asc("9")
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

Private Sub txtgPWD_Change()

    If osub.fRecID = 0 And Len(Trim(txtAccountName.Text)) > 0 And Len(Trim(txtgUID.Text)) > 0 And Len(Trim(txtgPWD.Text)) > 0 Then
        cmdAddAddress.Enabled = True
        cmdAddEmail.Enabled = True
        cmdAddPhone.Enabled = True
        cmdAddService.Enabled = True
        cmdChangePlan.Enabled = True
        cmdBill.Enabled = True
        tsDetails.Enabled = True
        osub.fAccountName = Trim(txtAccountName.Text)
        osub.fgPassword = Trim(txtgPWD.Text)
        osub.fgUsername = Trim(txtgUID.Text)
    End If
    
End Sub

Private Sub txtgPWD_KeyPress(KeyAscii As Integer)
    
    If osub.fRecID <> 0 Then
        KeyAscii = 0
        Dim fPass As String
        fPass = pass(osub.fgPassword, "Change Clients Online Password")
        If fPass <> osub.fgPassword Then
            osub.fgPassword = fPass
            txtgPWD.Text = osub.fgPassword
        End If
    End If
    
End Sub

Private Sub txtgPWD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtgPWD_MouseMove"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    
    If bPlay = False And picHeader(0).Visible = True Then
            
        L = GetCursorPos(oXY)
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y

        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "When the user of the account set, want to change his detail on the web and access other services this is the password that is used, it is stored as a MD5 Checksum."
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

Private Sub txtgUID_Change()

    If osub.fRecID = 0 And Len(Trim(txtAccountName.Text)) > 0 And Len(Trim(txtgUID.Text)) > 0 And Len(Trim(txtgPWD.Text)) > 0 Then
        cmdAddAddress.Enabled = True
        cmdAddEmail.Enabled = True
        cmdAddPhone.Enabled = True
        cmdAddService.Enabled = True
        cmdChangePlan.Enabled = True
        cmdBill.Enabled = True
        tsDetails.Enabled = True
    End If
    
End Sub

Private Sub txtgUID_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtgUID_MouseMove"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    
    If bPlay = False And picHeader(0).Visible = True Then
            
        L = GetCursorPos(oXY)
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "When the user of the account set, want to change his detail on the web and access other services this is the username that is used."
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

Private Sub txtSearchName_KeyPress(KeyAscii As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtSearchName_KeyPress"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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

    Select Case KeyAscii
    Case 13
        KeyAscii = 0
        pb1.Value = 0
        If Trim(txtSearchName) = "" Then Exit Sub
        If lvPlans.ListItems.Count = 0 Then
            MsgBox "You must create some services before you can select which ones where refered too"
            Exit Sub
        End If
        
        Dim rsload As adodb.Recordset
        
        If chkSearchAccountNames.Value <> 0 Then
            bResult = MySQL.OpenTable(directConn, rsload, , MySQL.virtualisp("select * from accountinfo Where AccountName Like '%" & txtSearchName.Text & "%'", "accountinfo", True, Login.bMaster))
            
            If rsload.RecordCount > 0 Then
                rsload.MoveFirst
                While Not rsload.EOF And Err.Number = 0
                    Set itmX = lvReferedBy.ListItems.Add(, "a" & rsload!RecID, rsload!AccountName)
                    rsload.MoveNext
                Wend
            End If
        End If
        
        pb1.Value = pb1.Value + 1
        
        If chkSearchContacts.Value <> 0 Then
        
            bResult = MySQL.OpenTable(directConn, rsload, , MySQL.virtualisp("select accountinfo.AccountName, acci_addresses.Contactname, acci_addresses.AccI_RecID, acci_addresses.RecID from accountinfo, acci_addresses Where acci_addresses.AccI_RecID = accountinfo.RecID AND acci_addresses.ContactName Like '%" & txtSearchName.Text & "%'", "accountinfo", True, Login.bMaster))
            
            If rsload.RecordCount > 0 Then
                rsload.MoveFirst
                While Not rsload.EOF And Err.Number = 0
                    Set itmX = lvReferedBy.ListItems.Add(, "d" & rsload!RecID & "_" & rsload!acci_RecID, rsload!AccountName)
                    itmX.SubItems(1) = rsload!ContactName
                    rsload.MoveNext
                Wend
            End If
        
        End If
        pb1.Value = pb1.Value + 1
        
        If chkSearchContacts.Value <> 0 Then
        
            bResult = MySQL.OpenTable(directConn, rsload, , MySQL.virtualisp("select accountinfo.AccountName, acci_emailaddresses.Contactname, acci_emailaddresses.AccI_RecID, acci_emailaddresses.RecID from accountinfo, acci_emailaddresses Where acci_emailaddresses.AccI_RecID = accountinfo.RecID AND acci_emailaddresses.ContactName Like '%" & txtSearchName.Text & "%'", "accountinfo", True, Login.bMaster))
            
            If rsload.RecordCount > 0 Then
                rsload.MoveFirst
                While Not rsload.EOF And Err.Number = 0
                    Set itmX = lvReferedBy.ListItems.Add(, "d" & rsload!RecID & "_" & rsload!acci_RecID, rsload!AccountName)
                    itmX.SubItems(1) = rsload!ContactName
                    rsload.MoveNext
                Wend
            End If
        
        End If
        pb1.Value = pb1.Value + 1
        
        If chkSearchContacts.Value <> 0 Then
        
            bResult = MySQL.OpenTable(directConn, rsload, , MySQL.virtualisp("select accountinfo.AccountName, acci_services.Contactname, acci_services.AccI_RecID, acci_services.RecID from accountinfo, acci_services Where acci_services.AccI_RecID = accountinfo.RecID AND acci_services.ContactName Like '%" & txtSearchName.Text & "%'", "accountinfo", True))
            
            If rsload.RecordCount > 0 Then
                rsload.MoveFirst
                While Not rsload.EOF And Err.Number = 0
                    Set itmX = lvReferedBy.ListItems.Add(, "d" & rsload!RecID & "_" & rsload!acci_RecID, rsload!AccountName)
                    itmX.SubItems(1) = rsload!ContactName
                    rsload.MoveNext
                Wend
            End If
        
        End If
        pb1.Value = pb1.Value + 1
        
        If chkSearchUsername.Value <> 0 Then
        
            If MySQL.OpenTable(directConn, rsload, , MySQL.virtualisp("select accountinfo.AccountName, acci_services.ContactName, acci_services.AccI_RecID, acci_services.RecID from accountinfo, acci_services Where acci_services.AccI_RecID = accountinfo.RecID AND acci_services.Username Like '%" & txtSearchName.Text & "%'", "accountinfo", True)) = True Then
            
                If rsload.RecordCount > 0 Then
                    rsload.MoveFirst
                    While Not rsload.EOF And Err.Number = 0
                        Set itmX = lvReferedBy.ListItems.Add(, "d" & rsload!RecID & "_" & rsload!acci_RecID, rsload!AccountName)
                        itmX.SubItems(1) = rsload!ContactName
                        rsload.MoveNext
                    Wend
                End If
            
            End If
            
        End If
        
        pb1.Value = pb1.Value + 1
        
        If chkSearchContacts.Value <> 0 Then
        
            bResult = MySQL.OpenTable(directConn, rsload, , MySQL.virtualisp("select accountinfo.AccountName, acci_phonenumbers.Contactname, acci_phonenumbers.AccI_RecID, acci_phonenumbers.RecID from accountinfo, acci_phonenumbers Where acci_phonenumbers.AccI_RecID = accountinfo.RecID AND acci_phonenumbers.ContactName Like '%" & txtSearchName.Text & "%'", "accountinfo", True))
            
            If rsload.RecordCount > 0 Then
                rsload.MoveFirst
                While Not rsload.EOF And Err.Number = 0
                    Set itmX = lvReferedBy.ListItems.Add(, "d" & rsload!RecID & "_" & rsload!acci_RecID, rsload!AccountName)
                    itmX.SubItems(1) = rsload!ContactName
                    rsload.MoveNext
                Wend
            End If
        End If
        pb1.Value = pb1.Value + 1
       
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


Public Function SaveColumnWidths()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "SaveColumnWidths"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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

   
    Call GUI.SaveColWidths(lvAddresses, Me)
    Call GUI.SaveColWidths(lvEmail, Me)
    Call GUI.SaveColWidths(lvPhone, Me)
    Call GUI.SaveColWidths(lvPlans, Me)
    Call GUI.SaveColWidths(lvTransactions, Me)
    'Call gui.SaveColWidths(lvFTPAccounts, Me)
        
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
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    Call GUI.LoadColWidths(lvAddresses, Me)
    Call GUI.LoadColWidths(lvEmail, Me)
    Call GUI.LoadColWidths(lvPhone, Me)
    Call GUI.LoadColWidths(lvPlans, Me)
    Call GUI.LoadColWidths(lvTransactions, Me)
    
    'Call gui.LoadColWidths(lvPOP3, Me)
    'Call gui.LoadColWidths(lvFTPAccounts, Me)
        
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

Public Function LoadAccountClass()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "LoadAccountClass"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    Dim bResult As Boolean
    
    bResult = MySQL.OpenTable(directConn, rsload, , "select * from accountclass")
    
    If rsload.RecordCount > 0 Then
        Dim itmX As ListItem
        While Not rsload.EOF And Err.Number = 0
            Set itmX = lvClassification.ListItems.Add(, "k" & rsload!RecID, rsload!Description)
            rsload.MoveNext
        Wend
    End If
    
    'lvClassification.ListItems.Add , "k1", "Commerical"
    'lvClassification.ListItems.Add , "k2", "Residential"
    'lvClassification.ListItems.Add , "k3", "SOHO"
    'lvClassification.ListItems.Add , "k4", ".ORG Enterprise"
    'lvClassification.ListItems.Add , "k5", "Technology Group"
    'lvClassification.ListItems.Add , "k6", "Industrial"
    
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

Public Function LoadFlags()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "LoadFlags"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    Dim bResult As Boolean
    
    bResult = MySQL.OpenTable(directConn, rsload, , "select * from flags Where FlagType = 'a'")
    
    If rsload.RecordCount > 0 Then
        
        While Not rsload.EOF And Err.Number = 0
            cmbFlagA.AddItem rsload!FlagDesc
            cmbFlagA.ItemData(cmbFlagA.ListCount - 1) = rsload!RecID
            rsload.MoveNext
        Wend
        cmbFlagA.ListIndex = 0
    End If
        
    bResult = MySQL.OpenTable(directConn, rsload, , "select * from flags Where FlagType = 'b'")
    
    If rsload.RecordCount > 0 Then
        rsload.MoveFirst
        While Not rsload.EOF And Err.Number = 0
            cmbFlagB.AddItem rsload!FlagDesc
            cmbFlagB.ItemData(cmbFlagB.ListCount - 1) = rsload!RecID
            rsload.MoveNext
        Wend
        cmbFlagB.ListIndex = 0
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

'Public Function SaveServices(RecID As Variant)
'
'
'    '*[ Error Checking Variables ]**********************************************************************************
'
'
'    Const RoutineName = "SaveServices"
'    Const ContainerName = "frmCustomerRec"
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
'    Dim itmX As ListItem
'    Dim sa As Integer
'    Dim rsSave As ADODB.Recordset
'    Dim rsRadius As ADODB.Recordset
'
'    If lvPlans.ListItems.Count > 0 Then
'
'        For sa = 1 To lvPlans.ListItems.Count
'            Set itmX = lvPlans.ListItems(sa)
'            If Left(itmX.Key, 1) <> "r" Then
'                If Left(itmX.Key, 1) <> "k" Then
'                    If Left(itmX.Key, 1) <> "x" Then
'
'                        Call MySQL.OpenTable(directConn, rsSave, , "select * from acci_services Limit 0")
'                        'rsSave.AddNew
'                        MySQL.Execute directConn, "INSERT INTO acci_services (acci_RecID, ServiceID, Username, Password, BaseURL, DynamicField1, DynamicField2, DynamicField3, DynamicField4, DynamicField5, Checked) VALUES ('" & osub.fRecID& "','" & itmX.Tag & "','" & MySQL.ESC(itmX.SubItems(2)) & "','" & MySQL.ESC(itmX.SubItems(3)) & "','" & MySQL.ESC(itmX.SubItems(4)) & "','" & MySQL.ESC(itmX.SubItems(5)) & "','" & MySQL.ESC(itmX.SubItems(6)) & "','" & MySQL.ESC(itmX.SubItems(7)) & "','" & MySQL.ESC(itmX.SubItems(8)) & "','" & MySQL.ESC(itmX.SubItems(9)) & "','" & IIf(itmX.Checked = True, "-1", "0") & "')"
'
'                        MySQL.Execute directConn, "UPDATE acci_services SET Password=AES_ENCRYPT('" & txtgPWD.Text & "','" & odb.colSalts.ReturnSalt("md5Password") & "'), Where `User-Name` = '" & itmX.SubItems(1) & "' and ServiceID = " & itmX.Tag
'
'                        If rsSave!RadiusID <> 0 Then
'                            bResult = MySQL.OpenTable(directConn, rsRadius, , "select * from radiusaccounts Where RecID = " & rsSave!RadiusID & " Limit 1")
'                            rsRadius!Checked = itmX.Checked
'                            rsRadius.Update
'                        End If
'
'                    ElseIf Left(itmX.Key, 1) = "x" Then
'
'                        Call MySQL.OpenTable(directConn, rsSave, , "select * from acci_services where RecID = " & Mid(itmX.Key, 2) & " Limit 1")
'
'                        If rsSave!RadiusID <> 0 Then
'                            MySQL.Execute directConn, "UPDATE radiusaccount Set Checked = " & IIf(itmX.Checked = True, "-1", "0") & " where RecID = " & rsSave!RadiusID
'                        End If
'
'                        MySQL.Execute directConn, "UPDATE acci_services SET ServiceID = '" & itmX.Tag & "', Username = '" & MySQL.ESC(itmX.SubItems(2)) & "', BaseURL='" & MySQL.ESC(itmX.SubItems(4)) & "',DynamicField1='" & MySQL.ESC(itmX.SubItems(5)) & "',DynamicField2='" & MySQL.ESC(itmX.SubItems(6)) & "',DynamicField3='" & MySQL.ESC(itmX.SubItems(7)) & "',DynamicField4='" & MySQL.ESC(itmX.SubItems(8)) & "',DynamicField5='" & MySQL.ESC(itmX.SubItems(9)) & "',Checked=" & IIf(itmX.Checked = True, "'-1'", "'0'") & " where RecID = rssave!RecID"
'
'                    End If
'                End If
'            End If
'        Next
'
'    End If
'
'
'
'
'Exit Function
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
'End Function

Public Sub SetFormAccess()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "SetFormAccess"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    cmbFlagA.Enabled = Login.bMaster
    cmbFlagB.Enabled = Login.bMaster
    'frameAppetite.Visible = Login.bMaster
    
    cmbSysopID.Locked = Not Login.bOwnership
    cmbVirtualID.Locked = Not Login.bOwnership
    
    Select Case Login.lLevel
    Case 0
        cmdAddAddress.Enabled = Not True
        cmdAddEmail.Enabled = Not True
        cmdAddPhone.Enabled = Not True
        cmdAddService.Enabled = Not True
        cmdChangePlan.Enabled = Not True
        cmdBill.Enabled = Not True
        dtpDOB.Enabled = Not True
        txtSearchName.Locked = Not False
        lvEmail.Enabled = Not True
        lvAddresses.Enabled = Not True
        lvPlans.Enabled = Not True
        lvTransactions.Enabled = Not True
        lvClassification.Enabled = Not True
        txtgPWD.Locked = Not False
        txtgUID.Locked = Not False
        txtAccountName.Locked = Not False
        cmbFlagA.Enabled = Not True
        cmbFlagB.Enabled = Not True
        chkCancelled.Enabled = Not True
        optBillingCycle(0).Enabled = Not True
        optBillingCycle(1).Enabled = Not True
        optBillingCycle(2).Enabled = Not True
        txtBillingCycle(0).Locked = Not False
        txtBillingCycle(1).Locked = Not False
        txtBillingCycle(2).Locked = Not False
        'txtRealm.Locked = Not False
        cmdSave.Enabled = False
    Case 1 To 20
        cmdAddAddress.Enabled = Not True
        cmdAddEmail.Enabled = Not True
        cmdAddPhone.Enabled = Not True
        cmdAddService.Enabled = Not True
        cmdChangePlan.Enabled = Not True
        cmdBill.Enabled = Not True
        dtpDOB.Enabled = True
        txtSearchName.Locked = False
        lvEmail.Enabled = Not True
        lvAddresses.Enabled = Not True
        lvPlans.Enabled = Not True
        lvTransactions.Enabled = Not True
        lvClassification.Enabled = Not True
        txtgPWD.Locked = Not False
        txtgUID.Locked = Not False
        txtAccountName.Locked = Not False
        cmbFlagA.Enabled = Not True
        cmbFlagB.Enabled = Not True
        chkCancelled.Enabled = Not True
        optBillingCycle(0).Enabled = Not True
        optBillingCycle(1).Enabled = Not True
        optBillingCycle(2).Enabled = Not True
        txtBillingCycle(0).Locked = Not False
        txtBillingCycle(1).Locked = Not False
        txtBillingCycle(2).Locked = Not False
        'txtRealm.Locked = Not False
    Case 21 To 40
        cmdAddAddress.Enabled = Not True
        cmdAddEmail.Enabled = Not True
        cmdAddPhone.Enabled = Not True
        cmdAddService.Enabled = Not True
        cmdChangePlan.Enabled = Not True
        cmdBill.Enabled = True
        dtpDOB.Enabled = True
        txtSearchName.Locked = False
        lvEmail.Enabled = True
        lvAddresses.Enabled = True
        lvPlans.Enabled = Not True
        lvTransactions.Enabled = Not True
        lvClassification.Enabled = True
        txtgPWD.Locked = False
        txtgUID.Locked = False
        txtAccountName.Locked = False
        cmbFlagA.Enabled = True
        cmbFlagB.Enabled = True
        chkCancelled.Enabled = True
        optBillingCycle(0).Enabled = Not True
        optBillingCycle(1).Enabled = Not True
        optBillingCycle(2).Enabled = Not True
        txtBillingCycle(0).Locked = Not False
        txtBillingCycle(1).Locked = Not False
        txtBillingCycle(2).Locked = Not False
        'txtRealm.Locked = Not False
    Case 41 To 60
        cmdAddAddress.Enabled = True
        cmdAddEmail.Enabled = True
        cmdAddPhone.Enabled = True
        cmdAddService.Enabled = True
        cmdChangePlan.Enabled = True
        cmdBill.Enabled = True
        dtpDOB.Enabled = True
        txtSearchName.Locked = False
        lvEmail.Enabled = True
        lvAddresses.Enabled = True
        lvPlans.Enabled = True
        lvTransactions.Enabled = True
        lvClassification.Enabled = True
        txtgPWD.Locked = False
        txtgUID.Locked = False
        txtAccountName.Locked = False
        cmbFlagA.Enabled = True
        cmbFlagB.Enabled = True
        chkCancelled.Enabled = True
        optBillingCycle(0).Enabled = True
        optBillingCycle(1).Enabled = True
        optBillingCycle(2).Enabled = True
        txtBillingCycle(0).Locked = False
        txtBillingCycle(1).Locked = False
        txtBillingCycle(2).Locked = False
        'txtRealm.Locked = False
    Case 61 To 80
        cmdAddAddress.Enabled = True
        cmdAddEmail.Enabled = True
        cmdAddPhone.Enabled = True
        cmdAddService.Enabled = True
        cmdChangePlan.Enabled = True
        cmdBill.Enabled = True
        dtpDOB.Enabled = True
        txtSearchName.Locked = False
        lvEmail.Enabled = True
        lvAddresses.Enabled = True
        lvPlans.Enabled = True
        lvTransactions.Enabled = True
        lvClassification.Enabled = True
        txtgPWD.Locked = False
        txtgUID.Locked = False
        txtAccountName.Locked = False
        cmbFlagA.Enabled = True
        cmbFlagB.Enabled = True
        chkCancelled.Enabled = True
        optBillingCycle(0).Enabled = True
        optBillingCycle(1).Enabled = True
        optBillingCycle(2).Enabled = True
        txtBillingCycle(0).Locked = False
        txtBillingCycle(1).Locked = False
        txtBillingCycle(2).Locked = False
        'txtRealm.Locked = False
    Case 81 To 99
        cmdAddAddress.Enabled = True
        cmdAddEmail.Enabled = True
        cmdAddPhone.Enabled = True
        cmdAddService.Enabled = True
        cmdChangePlan.Enabled = True
        cmdBill.Enabled = True
        dtpDOB.Enabled = True
        txtSearchName.Locked = False
        lvEmail.Enabled = True
        lvAddresses.Enabled = True
        lvPlans.Enabled = True
        lvTransactions.Enabled = True
        lvClassification.Enabled = True
        txtgPWD.Locked = False
        txtgUID.Locked = False
        txtAccountName.Locked = False
        cmbFlagA.Enabled = True
        cmbFlagB.Enabled = True
        chkCancelled.Enabled = True
        optBillingCycle(0).Enabled = True
        optBillingCycle(1).Enabled = True
        optBillingCycle(2).Enabled = True
        txtBillingCycle(0).Locked = False
        txtBillingCycle(1).Locked = False
        txtBillingCycle(2).Locked = False
        'txtRealm.Locked = False
        mvBilling.Enabled = True
    Case 100
        cmdAddAddress.Enabled = True
        cmdAddEmail.Enabled = True
        cmdAddPhone.Enabled = True
        cmdAddService.Enabled = True
        cmdChangePlan.Enabled = True
        cmdBill.Enabled = True
        dtpDOB.Enabled = True
        txtSearchName.Locked = False
        lvEmail.Enabled = True
        lvAddresses.Enabled = True
        lvPlans.Enabled = True
        lvTransactions.Enabled = True
        lvClassification.Enabled = True
        txtgPWD.Locked = False
        txtgUID.Locked = False
        txtAccountName.Locked = False
        cmbFlagA.Enabled = True
        cmbFlagB.Enabled = True
        chkCancelled.Enabled = True
        optBillingCycle(0).Enabled = True
        optBillingCycle(1).Enabled = True
        optBillingCycle(2).Enabled = True
        txtBillingCycle(0).Locked = False
        txtBillingCycle(1).Locked = False
        txtBillingCycle(2).Locked = False
        'txtRealm.Locked = False
        mvBilling.Enabled = True
    End Select

    If osub.fRecID = 0 Then
        cmdAddAddress.Enabled = False
        cmdAddEmail.Enabled = False
        cmdAddPhone.Enabled = False
        cmdAddService.Enabled = False
        cmdChangePlan.Enabled = False
        cmdBill.Enabled = False
        tsDetails.Enabled = False
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

Public Function LoadHardware()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "LoadHardware"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    Dim rsHardware As adodb.Recordset
    
    Dim ib As Byte
    
    For ib = cmbField.LBound To cmbField.UBound
        Call MySQL.OpenTable(directConn, rsHardware, , "select distinct " & cmbField(ib).Tag & " from acci_hardware")
        If rsHardware.RecordCount > 0 Then
            While Not rsHardware.EOF And Err.Number = 0
                If Not IsNull(rsHardware(cmbField(ib).Tag)) Then
                    If Trim(rsHardware(cmbField(ib).Tag)) <> "" Then
                        cmbField(ib).AddItem rsHardware(cmbField(ib).Tag)
                    End If
                End If
                rsHardware.MoveNext
            Wend
        End If
    Next

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

'Public Sub SaveHardware(oRecID As Variant)
'
'
'    '*[ Error Checking Variables ]**********************************************************************************
'
'
'    Const RoutineName = "SaveHardware"
'    Const ContainerName = "frmCustomerRec"
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
'    On Error Resume Next
'    Dim rsHardware As ADODB.Recordset
'
'    Call MySQL.OpenTable(directConn, rsHardware, , "select * from acci_hardware where acci_RecID = " & oRecID)
'
'    If rsHardware.RecordCount = 0 Then
'        MySQL.Execute directConn, "INSERT INTO acci_hardware (Modem, Processor, VideoCard, Monitor, PCType, Networkcard, OS, Printer, Mainboard, acci_RecID) " + _
'                               " VALUES('" + MySQL.ESC(cmbField(0).Text) + "','" + MySQL.ESC(cmbField(1).Text) + "','" + MySQL.ESC(cmbField(2).Text) + "','" + MySQL.ESC(cmbField(3).Text) + "','" + MySQL.ESC(cmbField(4).Text) + "','" + MySQL.ESC(cmbField(5).Text) + "','" + MySQL.ESC(cmbField(6).Text) + "','" + MySQL.ESC(cmbField(7).Text) + "','" + MySQL.ESC(cmbField(8).Text) + "'," & oRecID & ")"
'    Else
'        MySQL.Execute directConn, "UPDATE acci_hardware SET Modem = '" + MySQL.ESC(cmbField(0).Text) + "', Processor = '" + MySQL.ESC(cmbField(1).Text) + "', VideoCard = '" + MySQL.ESC(cmbField(2).Text) + "', Monitor = '" + MySQL.ESC(cmbField(3).Text) + "', PCType = '" + MySQL.ESC(cmbField(4).Text) + "', Networkcard = '" + MySQL.ESC(cmbField(5).Text) + "', OS = '" + MySQL.ESC(cmbField(6).Text) + "', Printer = '" + MySQL.ESC(cmbField(7).Text) + "', Mainboard = '" + MySQL.ESC(cmbField(8).Text) + "' where acci_RecId = " & oRecID
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

Private Sub txtSearchName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtSearchName_MouseMove"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    
    If bPlay = False And picTSContainer(2).Visible = True Then
            
        L = GetCursorPos(oXY)
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.GestureAt oXY.X, oXY.Y
        If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "Enter a name to search for in the database to select as a referee. Just enter the name and press enter, if there there check the box and fill in the options."
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

Public Function LoadVISPs()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "LoadVISPs"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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



    Dim SQL As Variant
    
    If ViSPMAP.Count > 0 Then
        Dim iCnt As Long
        For iCnt = 1 To ViSPMAP.Count
        
            SQL = SQL + "'" & ViSPMAP(iCnt).RecIDb & "', "
            
        Next
        SQL = Left(SQL, Len(SQL) - 2)
        
        Dim rsload As adodb.Recordset
        
        Call MySQL.OpenTable(directConn, rsload, , "select RecID In (" & SQL & ") as bFound, RecID, Description from virtualisp")
        
        If rsload.State = adStateOpen Then
            If rsload.RecordCount > 0 Then
                Do
                    cmbVirtualID.AddItem IIf(IsNull(rsload!Description), "(null)", rsload!Description)
                    cmbVirtualID.ItemData(cmbVirtualID.ListCount - 1) = rsload!RecID
                    If rsload!RecID = Login.lVirtualID Then cmbVirtualID.ListIndex = cmbVirtualID.ListCount - 1
                    
                    rsload.MoveNext
                    
                Loop Until rsload.EOF Or Err.Number <> 0
            End If
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

Public Sub LoadSysops()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "LoadSysops"
    Const ContainerName = "frmCustomerRec"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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



    Dim SQL As Variant
    
    If ViSPMAP.Count > 0 Then
        Dim iCnt As Long
        For iCnt = 1 To ViSPMAP.Count
        
            SQL = SQL + "'" & ViSPMAP(iCnt).RecIDb & "', "
            
        Next
        SQL = Left(SQL, Len(SQL) - 2)
        
        Dim rsload As adodb.Recordset
        
        Call MySQL.OpenTable(directConn, rsload, , "select VirtualID in (" & SQL & ") as bFound, RecID, Username, Firstname, Surname from sysops")
        
        If rsload.State = adStateOpen Then
            If rsload.RecordCount > 0 Then
                Do
                    cmbSysopID.AddItem IIf(IsNull(rsload!Username), "[(null)] - ", "[" & rsload!Username & "] - ") & IIf(IsNull(rsload!Firstname), "", rsload!Firstname) & " " & IIf(IsNull(rsload!Surname), "", rsload!Surname)
                    cmbSysopID.ItemData(cmbSysopID.ListCount - 1) = rsload!RecID
                    If rsload!RecID = Login.lSysopID Then cmbSysopID.ListIndex = cmbSysopID.ListCount - 1
                    rsload.MoveNext
                Loop Until rsload.EOF Or Err.Number <> 0
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

Private Sub wwwFiiledb_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)

    
End Sub

Private Sub wwwFiiledb_StatusTextChange(ByVal Text As String)

End Sub

Private Sub wwwFileDB_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)

    If InStr(URL, Me.UNCPath) = 0 Then
        Cancel = True
    End If
End Sub

