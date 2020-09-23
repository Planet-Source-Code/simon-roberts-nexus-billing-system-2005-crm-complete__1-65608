VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSysop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sysop"
   ClientHeight    =   8070
   ClientLeft      =   8835
   ClientTop       =   4605
   ClientWidth     =   8760
   Icon            =   "frmSysop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmSysop.frx":0442
   ScaleHeight     =   538
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   584
   Begin VB.PictureBox picTS 
      BorderStyle     =   0  'None
      Height          =   6825
      Index           =   2
      Left            =   630
      ScaleHeight     =   6825
      ScaleWidth      =   7725
      TabIndex        =   41
      Top             =   360
      Visible         =   0   'False
      Width           =   7725
      Begin VB.Frame Frame4 
         Caption         =   "Security Level"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   60
         TabIndex        =   44
         Top             =   2160
         Width           =   7605
         Begin MSComctlLib.Slider sld 
            Height          =   630
            Left            =   720
            TabIndex        =   45
            Top             =   720
            Width           =   6765
            _ExtentX        =   11933
            _ExtentY        =   1111
            _Version        =   393216
            Max             =   100
            TickFrequency   =   5
         End
         Begin MSComctlLib.ImageList ilSysops 
            Left            =   3510
            Top             =   180
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
                  Picture         =   "frmSysop.frx":13FFC
                  Key             =   "k100"
                  Object.Tag             =   "100% Security Access"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSysop.frx":14316
                  Key             =   "k080_099"
                  Object.Tag             =   "High Level Access"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSysop.frx":14630
                  Key             =   "k060_079"
                  Object.Tag             =   "Server Administrator"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSysop.frx":14F0A
                  Key             =   "k040_059"
                  Object.Tag             =   "Network Admin"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSysop.frx":1535C
                  Key             =   "k001_020"
                  Object.Tag             =   "Service Plans Editor"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSysop.frx":157AE
                  Key             =   "k020_039"
                  Object.Tag             =   "Rainbow Warrior"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSysop.frx":15C00
                  Key             =   "k000"
                  Object.Tag             =   "Low Level Access"
               EndProperty
            EndProperty
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Access Level Description"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   150
            TabIndex        =   46
            Top             =   360
            Width           =   3030
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   240
            Picture         =   "frmSysop.frx":16052
            Top             =   750
            Width           =   480
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Account Description"
         Height          =   1995
         Left            =   60
         TabIndex        =   42
         Top             =   120
         Width           =   7605
         Begin VB.TextBox txtField 
            Height          =   1605
            Index           =   17
            Left            =   150
            MultiLine       =   -1  'True
            TabIndex        =   43
            Tag             =   "Description"
            Top             =   240
            Width           =   7365
         End
      End
   End
   Begin VB.PictureBox picTS 
      BorderStyle     =   0  'None
      Height          =   6765
      Index           =   3
      Left            =   750
      ScaleHeight     =   6765
      ScaleWidth      =   7815
      TabIndex        =   47
      Top             =   360
      Width           =   7815
      Begin VB.Frame Frame9 
         Caption         =   "Finalisation Check List"
         Height          =   3795
         Left            =   3900
         TabIndex        =   53
         Top             =   2880
         Width           =   3735
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   255
            Left            =   3330
            TabIndex        =   65
            ToolTipText     =   "This will load form 1.1 in the browser so you can print it and get it filled out."
            Top             =   360
            Width           =   285
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Document Recieved at Exitstencil Press Australia"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   5
            Left            =   210
            TabIndex        =   64
            Tag             =   "cCompleted"
            Top             =   3180
            Width           =   3345
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Briefed New Sysop on commission scales and rates of pay by your ViSP"
            Height          =   435
            Index           =   4
            Left            =   210
            TabIndex        =   63
            Tag             =   "cBreif2"
            Top             =   2580
            Width           =   3345
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Briefed New Sysop on contact systems and price adjustments. "
            Height          =   435
            Index           =   3
            Left            =   210
            TabIndex        =   62
            Tag             =   "cBreif1"
            Top             =   2010
            Width           =   3345
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Mailed Taxation form and Form 1.1 to Exitstencil Press Australia"
            Height          =   435
            Index           =   2
            Left            =   210
            TabIndex        =   61
            Tag             =   "cMailedForms"
            Top             =   1440
            Width           =   3345
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Got Sysop to fill out australian taxation employment declaration"
            Height          =   435
            Index           =   1
            Left            =   210
            TabIndex        =   60
            Tag             =   "cFilledTaxation"
            Top             =   870
            Width           =   3345
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Got Sysop to fill out application form for sysop account on project alpha"
            Height          =   435
            Index           =   0
            Left            =   210
            TabIndex        =   59
            Tag             =   "cFilledApplication"
            Top             =   300
            Width           =   3345
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Tax File Number"
         Height          =   885
         Left            =   3900
         TabIndex        =   52
         Top             =   1920
         Width           =   3735
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
            Index           =   20
            Left            =   120
            TabIndex        =   56
            Tag             =   "TFN"
            Top             =   270
            Width           =   3465
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Rate of Superannuation (%)"
         Height          =   855
         Left            =   3900
         TabIndex        =   51
         Top             =   990
         Width           =   3735
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
            Index           =   19
            Left            =   270
            TabIndex        =   55
            Tag             =   "SuperRate"
            Text            =   "10"
            Top             =   300
            Width           =   3030
         End
         Begin MSComCtl2.UpDown UpDown2 
            Height          =   420
            Left            =   3300
            TabIndex        =   58
            Top             =   300
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   741
            _Version        =   393216
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtField(19)"
            BuddyDispid     =   196616
            BuddyIndex      =   19
            OrigLeft        =   3390
            OrigTop         =   300
            OrigRight       =   3645
            OrigBottom      =   735
            Max             =   50
            Min             =   1
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Income Tax Rate (%)"
         Height          =   855
         Left            =   3900
         TabIndex        =   50
         Top             =   60
         Width           =   3735
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   420
            Left            =   3270
            TabIndex        =   57
            Top             =   270
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   741
            _Version        =   393216
            Value           =   10
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtField(18)"
            BuddyDispid     =   196616
            BuddyIndex      =   18
            OrigLeft        =   3390
            OrigTop         =   270
            OrigRight       =   3645
            OrigBottom      =   705
            Max             =   50
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
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
            Index           =   18
            Left            =   270
            TabIndex        =   54
            Tag             =   "IncomeRate"
            Text            =   "10"
            Top             =   270
            Width           =   3000
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Link Sysop Account with VISP"
         Height          =   6645
         Left            =   180
         TabIndex        =   48
         Top             =   60
         Width           =   3585
         Begin VB.ListBox lstVISP 
            Height          =   6105
            Left            =   150
            TabIndex        =   49
            Top             =   330
            Width           =   3285
         End
      End
   End
   Begin VB.PictureBox picTS 
      BorderStyle     =   0  'None
      Height          =   3105
      Index           =   1
      Left            =   780
      ScaleHeight     =   3105
      ScaleWidth      =   2955
      TabIndex        =   28
      Top             =   450
      Visible         =   0   'False
      Width           =   2955
      Begin VB.Frame Frame1 
         Caption         =   "Country"
         Height          =   825
         Index           =   16
         Left            =   4020
         TabIndex        =   39
         Top             =   2790
         Width           =   3645
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
            Height          =   375
            Index           =   16
            Left            =   90
            TabIndex        =   40
            Tag             =   "Country"
            Top             =   270
            Width           =   3435
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "State"
         Height          =   825
         Index           =   15
         Left            =   180
         TabIndex        =   37
         Top             =   2790
         Width           =   3645
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
            Height          =   375
            Index           =   15
            Left            =   90
            TabIndex        =   38
            Tag             =   "State"
            Top             =   270
            Width           =   3435
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Postcode"
         Height          =   825
         Index           =   14
         Left            =   4020
         TabIndex        =   35
         Top             =   1890
         Width           =   3645
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
            Height          =   375
            Index           =   14
            Left            =   90
            TabIndex        =   36
            Tag             =   "Postcode"
            Top             =   270
            Width           =   3435
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Suburb"
         Height          =   825
         Index           =   13
         Left            =   180
         TabIndex        =   33
         Top             =   1890
         Width           =   3645
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
            Height          =   375
            Index           =   13
            Left            =   90
            TabIndex        =   34
            Tag             =   "Suburb"
            Top             =   270
            Width           =   3435
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Street Two"
         Height          =   825
         Index           =   12
         Left            =   180
         TabIndex        =   31
         Top             =   1020
         Width           =   7485
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
            Height          =   375
            Index           =   12
            Left            =   90
            TabIndex        =   32
            Tag             =   "Street2"
            Top             =   270
            Width           =   7275
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Street One"
         Height          =   825
         Index           =   11
         Left            =   180
         TabIndex        =   29
         Top             =   150
         Width           =   7485
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
            Height          =   375
            Index           =   11
            Left            =   90
            TabIndex        =   30
            Tag             =   "Street1"
            Top             =   270
            Width           =   7275
         End
      End
   End
   Begin VB.PictureBox picTS 
      BorderStyle     =   0  'None
      Height          =   6795
      Index           =   0
      Left            =   690
      ScaleHeight     =   6795
      ScaleWidth      =   7785
      TabIndex        =   2
      Top             =   420
      Width           =   7785
      Begin VB.OptionButton Option1 
         Caption         =   "Pay by BPAY"
         Height          =   435
         Index           =   2
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   27
         Tag             =   "bPayMethod"
         Top             =   6060
         Width           =   2385
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Pay by Direct Deposit"
         Height          =   435
         Index           =   1
         Left            =   2820
         Style           =   1  'Graphical
         TabIndex        =   26
         Tag             =   "bPayMethod"
         Top             =   6060
         Width           =   2325
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Pay by cheque"
         Height          =   435
         Index           =   0
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   25
         Tag             =   "bPayMethod"
         Top             =   6060
         Width           =   2355
      End
      Begin VB.Frame Frame1 
         Caption         =   "BPay Number"
         Height          =   765
         Index           =   10
         Left            =   3990
         TabIndex        =   23
         Top             =   5130
         Width           =   3675
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
            Height          =   375
            Index           =   10
            Left            =   90
            TabIndex        =   24
            Tag             =   "bPayNo"
            Top             =   270
            Width           =   3495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Branch BSB"
         Height          =   765
         Index           =   9
         Left            =   3990
         TabIndex        =   21
         Top             =   4290
         Width           =   3675
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
            Height          =   375
            Index           =   9
            Left            =   90
            TabIndex        =   22
            Tag             =   "BSB"
            Top             =   270
            Width           =   3495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Bank Account Number"
         Height          =   765
         Index           =   8
         Left            =   3990
         TabIndex        =   19
         Top             =   3480
         Width           =   3675
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
            Height          =   375
            Index           =   8
            Left            =   90
            TabIndex        =   20
            Tag             =   "AccountNo"
            Top             =   270
            Width           =   3495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Mobile Phone Number"
         Height          =   765
         Index           =   7
         Left            =   210
         TabIndex        =   17
         Top             =   5130
         Width           =   3675
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
            Height          =   375
            Index           =   7
            Left            =   90
            TabIndex        =   18
            Tag             =   "Mobile"
            Top             =   270
            Width           =   3495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Work Phone Number"
         Height          =   765
         Index           =   6
         Left            =   210
         TabIndex        =   15
         Top             =   4290
         Width           =   3675
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
            Height          =   375
            Index           =   6
            Left            =   90
            TabIndex        =   16
            Tag             =   "Work"
            Top             =   270
            Width           =   3495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Home Phone Number"
         Height          =   765
         Index           =   5
         Left            =   210
         TabIndex        =   13
         Top             =   3480
         Width           =   3675
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
            Height          =   375
            Index           =   5
            Left            =   90
            TabIndex        =   14
            Tag             =   "Home"
            Top             =   270
            Width           =   3495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Email Address Two"
         Height          =   765
         Index           =   4
         Left            =   210
         TabIndex        =   11
         Top             =   2640
         Width           =   7485
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
            Height          =   375
            Index           =   4
            Left            =   90
            TabIndex        =   12
            Tag             =   "Email2"
            Top             =   270
            Width           =   7275
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Email Address One"
         Height          =   765
         Index           =   3
         Left            =   210
         TabIndex        =   9
         Top             =   1800
         Width           =   7485
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
            Height          =   375
            Index           =   3
            Left            =   90
            TabIndex        =   10
            Tag             =   "Email1"
            Top             =   270
            Width           =   7275
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Surname"
         Height          =   765
         Index           =   2
         Left            =   3990
         TabIndex        =   7
         Top             =   990
         Width           =   3675
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
            Height          =   375
            Index           =   2
            Left            =   90
            TabIndex        =   8
            Tag             =   "Surname"
            Top             =   270
            Width           =   3495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "First name"
         Height          =   765
         Index           =   1
         Left            =   210
         TabIndex        =   5
         Top             =   990
         Width           =   3675
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
            Height          =   375
            Index           =   1
            Left            =   90
            TabIndex        =   6
            Tag             =   "Firstname"
            Top             =   270
            Width           =   3495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Username"
         Height          =   765
         Index           =   0
         Left            =   210
         TabIndex        =   3
         Top             =   150
         Width           =   7485
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
            Height          =   375
            Index           =   0
            Left            =   90
            TabIndex        =   4
            Tag             =   "Username"
            Top             =   270
            Width           =   7275
         End
      End
   End
   Begin MSComctlLib.TabStrip ts 
      Height          =   6885
      Left            =   330
      TabIndex        =   0
      Top             =   270
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   12144
      MultiRow        =   -1  'True
      Placement       =   2
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Personal Details"
            Object.ToolTipText     =   "Here is where you adjust the personal information for the Sysop"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Postal Address"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Constraints"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Links"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0086D28D&
      Height          =   315
      Left            =   420
      TabIndex        =   1
      Top             =   7470
      Width           =   1245
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0086D28D&
      BorderWidth     =   3
      Height          =   405
      Left            =   330
      Shape           =   4  'Rounded Rectangle
      Top             =   7410
      Width           =   1395
   End
End
Attribute VB_Name = "frmSysop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RecID As Double

Private Sub Check1_Click(Index As Integer)

    Static lastIndex As Integer
    
    If lastIndex = Index + 1 Then Exit Sub
    lastIndex = Index + 1
    
    If Check1(Index).Value = 2 Then Check1(Index).Value = 0
    If Check1(Index).Value = 1 Then Check1(Index).Value = 2
    If Check1(Index).Value = 0 Then Check1(Index).Value = 1
    
    
End Sub

Private Sub Command1_Click()

    Shell "http://www.ep.net.au/docs/projectalpha/sysop_application_1.1.doc"
    
End Sub

Private Sub Form_Load()

    On Error Resume Next

    Dim rsLoad As ADODB.Recordset
    Dim rsVISP As ADODB.Recordset
    Dim lVirtualID As Long
    
    Dim cx As Byte
    
    'For cx = optConst.LBound To optConst.UBound
    '    If Login.bMaster = False And Login.bPrimary = False Then optConst(cx).Enabled = False
    '    gSleep
    'Next cx

    Check1(5).Enabled = Login.bMaster
    
    If RecID <> 0 Then
        
        Call oMySQL.OpenTable(oConn, rsLoad, , "Select * from sysops where RecID = " & RecID)
        
        
        sld.Value = Val(rsLoad!SecurityLevel)
        Call sld_Click
        Dim bx As Byte
        
        Me.Caption = "Sysops - " & IIf(IsNull(rsLoad!Username), "", rsLoad!Username)
        
        For bx = txtField.LBound To txtField.UBound
            If Not IsNull(rsLoad(txtField(bx).Tag)) Then
                txtField(bx) = rsLoad(txtField(bx).Tag)
            Else
                txtField(bx) = ""
            End If
        Next bx
    
        For bx = Check1.LBound To Check1.UBound
            If Not IsNull(rsLoad(Check1(bx).Tag)) Then
                Check1(bx).Value = rsLoad(Check1(bx).Tag)
            Else
                Check1(bx).Value = 0
            End If
        Next bx
    
    
        If Not IsNull(rsLoad!bPayMethod) Then Option1(Val(rsLoad!bPayMethod)).Value = True
        
        lVirtualID = rsLoad!VirtualID
        
    End If
    
    Select Case Login.bMaster
    Case True
        Call oMySQL.OpenTable(oConn, rsVISP, , "Select RecID, Description from virtualisp")
    Case Else
    
    
        Call oMySQL.OpenTable(oConn, rsVISP, , "Select RecID, Description from virtualisp where RecID = " & Login.lVirtualID & " or VirtualID = " & Login.lVirtualID)
    End Select
    
    If rsVISP.RecordCount > 0 Then
        While Not rsVISP.EOF
            lstVISP.AddItem IIf(IsNull(rsVISP!Description), "", rsVISP!Description)
            If rsVISP!RecID = lVirtualID Then lstVISP.ListIndex = lstVISP.ListCount - 1
            lstVISP.ItemData(lstVISP.ListCount - 1) = rsVISP!RecID
            rsVISP.MoveNext
        Wend
    
    End If
    
    Call ts_Click
    
End Sub

Private Sub Label1_Click()

    Dim bIndex As Boolean
    
    If Len(Trim(txtField(0).Text)) < 4 Then
        MsgBox " The username must be four characters or numbers for the account."
        Exit Sub
    End If
    
    If RecID = 0 Then
        
        
        Dim Password As String
        Dim oPWD As New frmPWD
        
        oPWD.Show 1
        
        On Error Resume Next
        Do
            Err.Clear
            RecID = oMySQL.GetTMPRecID("sysops", oConn)
            Call oMySQL.Execute(oConn, "Insert into sysops (RecID,Password) VALUES ('" & RecID & "',encode('" & oPWD.outPassword & "','" & PasswordSalt & "'))")
        Loop Until Err.Number = 0
        bIndex = True
        
        Dim oPerm As New frmPerm
        oPerm.RecID = RecID
        oPerm.Show
        
    End If
    
    Me.Hide
        
        Call oMySQL.Execute(oConn, "update sysops set SecurityLevel = '" & sld.Value & "' where RecID = " & RecID)
        
        Dim bx As Byte
        
        For bx = txtField.LBound To txtField.UBound
            Call oMySQL.Execute(oConn, "update sysops set `" & txtField(bx).Tag & "` = '" & oMySQL.ESC(txtField(bx).Text) & "' where RecID = " & RecID)
        Next bx
    
        For bx = Check1.LBound To Check1.UBound
            Call oMySQL.Execute(oConn, "update sysops set `" & Check1(bx).Tag & "` = '" & Check1(bx).Value & "' where RecID = " & RecID)
        Next bx
        
        
        For bx = txtField.LBound To txtField.UBound
            Call oMySQL.Execute(oConn, "update sysops set `" & txtField(bx).Tag & "` = '" & oMySQL.ESC(txtField(bx).Text) & "' where RecID = " & RecID)
        Next bx

        Call oMySQL.Execute(oConn, "update sysops set `VirtualID` = '" & lstVISP.ItemData(lstVISP.ListIndex) & "' where RecID = " & RecID)

        For bx = Option1.LBound To Option1.UBound
            If Option1(bx).Value = True Then
                Call oMySQL.Execute(oConn, "update sysops set `" + Option1(bx).Tag + "` = '" & bx & "' where RecID = " & RecID)
                Exit For
            End If
        Next bx

    Dim itmX As ListItem
    
    Select Case bIndex
    Case True
        Set itmX = frmMain.lvSysops.ListItems.Add(, "r" & RecID, String(3 - Len("" & sld.Value), "0") + "" & sld.Value)
        itmX.Tag = oPWD.outPassword
    Case False
        Set itmX = frmMain.lvSysops.SelectedItem
    End Select
    
    itmX.SubItems(1) = txtField(0)
    itmX.SubItems(2) = txtField(1) + " " + txtField(2)
    itmX.SubItems(3) = "[--]"
    itmX.SubItems(4) = "[--]"
    itmX.SubItems(5) = "[--]"
    itmX.SubItems(6) = "[--]"
    itmX.SubItems(7) = "[--]"
    itmX.SubItems(8) = lstVISP.List(lstVISP.ListIndex)
    
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
        If CByte(sld.Value) >= imgMin And CByte(sld.Value) <= imgMax Then
            itmX.SmallIcon = imgX
        End If
    Next imgX

    
    
    
    Unload Me
    

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Shape1.BorderColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
    
End Sub

Private Sub sld_Click()

    Dim imgX As Byte
    Dim imgMin As Byte
    Dim imgMax As Byte
        
    If sld.Enabled = True Then
        If sld.Value > Login.lLevel Then sld.Value = Login.lLevel
    End If
    
    For imgX = 1 To ilSysops.ListImages.Count
        If InStr(ilSysops.ListImages(imgX).Key, "_") > 0 Then
            imgMin = CByte(Mid(ilSysops.ListImages(imgX).Key, 2, 3))
            imgMax = CByte(Mid(ilSysops.ListImages(imgX).Key, 6, 3))
        Else
            imgMin = CByte(Mid(ilSysops.ListImages(imgX).Key, 2, 3))
            imgMax = imgMin
        End If
        If sld.Value >= imgMin And sld.Value <= imgMax Then
            Image1.Picture = ilSysops.ListImages(imgX).Picture
            Label2.Caption = "" & sld.Value & "% - " & ilSysops.ListImages(imgX).Tag
            Exit For
        End If
    Next imgX
    
    
End Sub

Private Sub ts_Click()

    Dim bx As Byte
    
    For bx = picTS.LBound To picTS.UBound
        
        Select Case ts.SelectedItem.Index
        Case bx + 1
            picTS(bx).Move ts.ClientLeft, ts.ClientTop, ts.ClientWidth, ts.ClientHeight
            picTS(bx).ZOrder 0
            picTS(bx).Visible = True
        Case Else
            picTS(bx).Visible = False
        End Select
    Next
    
End Sub

Private Sub txtField_GotFocus(Index As Integer)

    txtField(Index).SelStart = 0
    txtField(Index).SelLength = Len(txtField(Index).Text)

End Sub

Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)

    Select Case Index
    Case 18, 19
    
        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or (InStr(txtField(Index).Text, ".") = 0 And KeyAscii = Asc(".")) Or KeyAscii = 8 Then
        
        Else
            If KeyAscii = 13 Then
            
            Else
                KeyAscii = 0
            End If
        
        End If
    Case Else
    
    
    End Select
    
    Select Case KeyAscii
    Case 13
    
        KeyAscii = 0
        SendKeys "{TAB}"
    
    End Select
    
End Sub
