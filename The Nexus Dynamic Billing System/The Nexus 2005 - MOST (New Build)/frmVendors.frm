VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmVendors 
   BackColor       =   &H009191F4&
   Caption         =   "Vendors"
   ClientHeight    =   10470
   ClientLeft      =   2295
   ClientTop       =   3150
   ClientWidth     =   14055
   Icon            =   "frmVendors.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10470
   ScaleWidth      =   14055
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picts 
      BackColor       =   &H00EC7A71&
      BorderStyle     =   0  'None
      Height          =   1155
      Index           =   3
      Left            =   300
      ScaleHeight     =   1155
      ScaleWidth      =   1335
      TabIndex        =   34
      Top             =   450
      Visible         =   0   'False
      Width           =   1335
      Begin VB.CommandButton cmdAddAddress 
         Caption         =   "&Add Address"
         Height          =   285
         Left            =   150
         TabIndex        =   35
         Top             =   2220
         Width           =   1065
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
         Picture         =   "frmVendors.frx":030A
      End
   End
   Begin VB.PictureBox picts 
      BackColor       =   &H00EC7A71&
      BorderStyle     =   0  'None
      Height          =   1305
      Index           =   1
      Left            =   3420
      ScaleHeight     =   1305
      ScaleWidth      =   1395
      TabIndex        =   31
      Top             =   360
      Visible         =   0   'False
      Width           =   1395
      Begin VB.CommandButton cmdAddPhone 
         Caption         =   "&Add Phone"
         Height          =   285
         Left            =   90
         TabIndex        =   32
         Top             =   6720
         Width           =   1305
      End
      Begin MSComctlLib.ListView lvPhone 
         Height          =   6585
         Left            =   120
         TabIndex        =   33
         Top             =   150
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   11615
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
         Picture         =   "frmVendors.frx":0F8D
      End
   End
   Begin VB.PictureBox picts 
      BackColor       =   &H00EC7A71&
      BorderStyle     =   0  'None
      Height          =   1215
      Index           =   2
      Left            =   1830
      ScaleHeight     =   1215
      ScaleWidth      =   1455
      TabIndex        =   28
      Top             =   390
      Visible         =   0   'False
      Width           =   1455
      Begin VB.CommandButton cmdAddEmail 
         Caption         =   "&Add e-Mail"
         Height          =   285
         Left            =   60
         TabIndex        =   29
         Top             =   1950
         Width           =   1305
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
         Picture         =   "frmVendors.frx":166C
      End
   End
   Begin VB.PictureBox picts 
      BackColor       =   &H009191F4&
      BorderStyle     =   0  'None
      Height          =   6405
      Index           =   0
      Left            =   480
      ScaleHeight     =   6405
      ScaleWidth      =   12915
      TabIndex        =   2
      Top             =   3690
      Width           =   12915
      Begin VB.Frame Frame1 
         BackColor       =   &H006977A7&
         Caption         =   "DSL Domain name"
         ForeColor       =   &H00FFFFFF&
         Height          =   705
         Index           =   14
         Left            =   5010
         TabIndex        =   51
         Top             =   4710
         Width           =   5355
         Begin VB.TextBox txtField 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   14
            Left            =   150
            MaxLength       =   128
            TabIndex        =   52
            Tag             =   "DSLDomain"
            Top             =   270
            Width           =   5055
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H005EB76B&
         Caption         =   "Purchase Orders Are Emailed To"
         Height          =   705
         Index           =   13
         Left            =   5010
         TabIndex        =   49
         Top             =   3930
         Width           =   5355
         Begin VB.TextBox txtField 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   13
            Left            =   150
            MaxLength       =   128
            TabIndex        =   50
            Tag             =   "poemailaddy"
            Top             =   270
            Width           =   5055
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H001394F2&
         Caption         =   "Functions"
         Height          =   1275
         Left            =   10470
         TabIndex        =   46
         Top             =   120
         Width           =   2115
         Begin VB.CommandButton cmFnc 
            Appearance      =   0  'Flat
            BackColor       =   &H001394F2&
            Caption         =   "Save Current Changes"
            Enabled         =   0   'False
            Height          =   345
            Index           =   1
            Left            =   180
            MaskColor       =   &H001394F2&
            TabIndex        =   48
            Top             =   750
            Width           =   1785
         End
         Begin VB.CommandButton cmFnc 
            Appearance      =   0  'Flat
            BackColor       =   &H001394F2&
            Caption         =   "New Vendor"
            Height          =   345
            Index           =   0
            Left            =   180
            MaskColor       =   &H001394F2&
            TabIndex        =   47
            Top             =   300
            Width           =   1785
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H006858B6&
         ForeColor       =   &H00FFFFFF&
         Height          =   705
         Index           =   1
         Left            =   7890
         TabIndex        =   44
         Top             =   5490
         Width           =   2475
         Begin VB.CheckBox Check1 
            BackColor       =   &H006858B6&
            Caption         =   "This Vendor is Tax Free"
            ForeColor       =   &H00FFFFFF&
            Height          =   405
            Index           =   1
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   45
            Tag             =   "TaxFree"
            Top             =   210
            Width           =   2235
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H0047ADE4&
         Caption         =   "Directors"
         ForeColor       =   &H00404040&
         Height          =   2595
         Index           =   1
         Left            =   60
         TabIndex        =   37
         Top             =   3930
         Width           =   4845
         Begin VB.Frame Frame1 
            BackColor       =   &H0047ADE4&
            Caption         =   "Director One"
            ForeColor       =   &H00404040&
            Height          =   705
            Index           =   12
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   4635
            Begin VB.TextBox txtField 
               Alignment       =   2  'Center
               Height          =   315
               Index           =   12
               Left            =   90
               MaxLength       =   128
               TabIndex        =   43
               Tag             =   "cDirector1"
               Top             =   270
               Width           =   4455
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H0047ADE4&
            Caption         =   "Director Two"
            ForeColor       =   &H00404040&
            Height          =   705
            Index           =   11
            Left            =   120
            TabIndex        =   40
            Top             =   990
            Width           =   4635
            Begin VB.TextBox txtField 
               Alignment       =   2  'Center
               Height          =   315
               Index           =   11
               Left            =   90
               MaxLength       =   128
               TabIndex        =   41
               Tag             =   "cDirector2"
               Top             =   270
               Width           =   4455
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H0047ADE4&
            Caption         =   "Director Three"
            ForeColor       =   &H00404040&
            Height          =   705
            Index           =   10
            Left            =   120
            TabIndex        =   38
            Top             =   1770
            Width           =   4635
            Begin VB.TextBox txtField 
               Alignment       =   2  'Center
               Height          =   315
               Index           =   10
               Left            =   90
               MaxLength       =   128
               TabIndex        =   39
               Tag             =   "cDirector3"
               Top             =   270
               Width           =   4455
            End
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C86A8D&
         Caption         =   "Customer Support Information"
         ForeColor       =   &H004FD2F9&
         Height          =   1065
         Left            =   60
         TabIndex        =   23
         Top             =   2790
         Width           =   10335
         Begin VB.Frame Frame1 
            BackColor       =   &H00C86A8D&
            Caption         =   "Support Phone Email Address"
            ForeColor       =   &H004FD2F9&
            Height          =   735
            Index           =   9
            Left            =   4890
            TabIndex        =   26
            Top             =   210
            Width           =   5295
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
               Height          =   405
               Index           =   9
               Left            =   90
               MaxLength       =   128
               TabIndex        =   27
               Tag             =   "cSupportPhone"
               Top             =   240
               Width           =   5085
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C86A8D&
            Caption         =   "Support Phone Number"
            ForeColor       =   &H004FD2F9&
            Height          =   735
            Index           =   8
            Left            =   120
            TabIndex        =   24
            Top             =   210
            Width           =   4635
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
               Height          =   405
               Index           =   8
               Left            =   120
               MaxLength       =   128
               TabIndex        =   25
               Tag             =   "SupportEmail"
               Top             =   240
               Width           =   4425
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H006858B6&
         ForeColor       =   &H00FFFFFF&
         Height          =   705
         Index           =   0
         Left            =   5010
         TabIndex        =   21
         Top             =   5490
         Width           =   2805
         Begin VB.CheckBox Check1 
            BackColor       =   &H006858B6&
            Caption         =   "This Vendor is Currently Active"
            ForeColor       =   &H00FFFFFF&
            Height          =   405
            Index           =   0
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   22
            Tag             =   "Active"
            Top             =   210
            Width           =   2625
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EC7A71&
         Caption         =   "Vendor's Details"
         ForeColor       =   &H00FFFFFF&
         Height          =   2595
         Index           =   0
         Left            =   60
         TabIndex        =   14
         Top             =   120
         Width           =   4845
         Begin VB.Frame Frame1 
            BackColor       =   &H00EC7A71&
            Caption         =   "ACN"
            ForeColor       =   &H00FFFFFF&
            Height          =   705
            Index           =   2
            Left            =   120
            TabIndex        =   19
            Top             =   1770
            Width           =   4635
            Begin VB.TextBox txtField 
               Alignment       =   2  'Center
               Height          =   315
               Index           =   2
               Left            =   90
               MaxLength       =   128
               TabIndex        =   20
               Tag             =   "ACN"
               Top             =   270
               Width           =   4455
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00EC7A71&
            Caption         =   "ABN"
            ForeColor       =   &H00FFFFFF&
            Height          =   705
            Index           =   1
            Left            =   120
            TabIndex        =   17
            Top             =   990
            Width           =   4635
            Begin VB.TextBox txtField 
               Alignment       =   2  'Center
               Height          =   315
               Index           =   1
               Left            =   90
               MaxLength       =   128
               TabIndex        =   18
               Tag             =   "ABN"
               Top             =   270
               Width           =   4455
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00EC7A71&
            Caption         =   "Vendors Name"
            ForeColor       =   &H00FFFFFF&
            Height          =   705
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   4635
            Begin VB.TextBox txtField 
               Alignment       =   2  'Center
               Height          =   315
               Index           =   0
               Left            =   90
               MaxLength       =   128
               TabIndex        =   16
               Tag             =   "vName"
               Top             =   270
               Width           =   4455
            End
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C4AF84&
         Caption         =   "Banking Details"
         Height          =   1845
         Left            =   5010
         TabIndex        =   5
         Top             =   120
         Width           =   5355
         Begin VB.Frame Frame1 
            BackColor       =   &H00C4AF84&
            Caption         =   "Account Name"
            Height          =   705
            Index           =   7
            Left            =   2730
            TabIndex        =   12
            Top             =   1020
            Width           =   2505
            Begin VB.TextBox txtField 
               Alignment       =   2  'Center
               Height          =   315
               Index           =   7
               Left            =   90
               MaxLength       =   128
               TabIndex        =   13
               Tag             =   "AccountName"
               Top             =   270
               Width           =   2295
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C4AF84&
            Caption         =   "Account Number"
            Height          =   705
            Index           =   6
            Left            =   120
            TabIndex        =   10
            Top             =   1020
            Width           =   2505
            Begin VB.TextBox txtField 
               Alignment       =   2  'Center
               Height          =   315
               Index           =   6
               Left            =   90
               MaxLength       =   128
               TabIndex        =   11
               Tag             =   "cAccount"
               Top             =   270
               Width           =   2295
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C4AF84&
            Caption         =   "Bank BSB"
            Height          =   705
            Index           =   5
            Left            =   2730
            TabIndex        =   8
            Top             =   240
            Width           =   2505
            Begin VB.TextBox txtField 
               Alignment       =   2  'Center
               Height          =   315
               Index           =   5
               Left            =   90
               MaxLength       =   128
               TabIndex        =   9
               Tag             =   "cBSB"
               Top             =   270
               Width           =   2295
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C4AF84&
            Caption         =   "BPay Number"
            Height          =   705
            Index           =   4
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   2505
            Begin VB.TextBox txtField 
               Alignment       =   2  'Center
               Height          =   315
               Index           =   4
               Left            =   90
               MaxLength       =   128
               TabIndex        =   7
               Tag             =   "BPay"
               Top             =   270
               Width           =   2295
            End
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C4AF84&
         Caption         =   "TFN"
         Height          =   705
         Index           =   3
         Left            =   5010
         TabIndex        =   3
         Top             =   2010
         Width           =   5355
         Begin VB.TextBox txtField 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   3
            Left            =   150
            MaxLength       =   128
            TabIndex        =   4
            Tag             =   "cTFN"
            Top             =   270
            Width           =   5055
         End
      End
   End
   Begin MSComctlLib.TabStrip ts 
      Height          =   4245
      Left            =   150
      TabIndex        =   1
      Top             =   3450
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   7488
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Main Details"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Phone Numbers"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "eMail Addresses"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Addresses"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lvVendors 
      Height          =   3225
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   13785
      _ExtentX        =   24315
      _ExtentY        =   5689
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "!RecID!"
         Text            =   "Vendor ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "!vName!"
         Text            =   "Vendor Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "!ABN!"
         Text            =   "ABN"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "!Director1!"
         Text            =   "Director 1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "!Active!^YesNo^"
         Text            =   "Active"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "!RecID!^select count(*) as nResult from plantemplates where VendorID = ^"
         Text            =   "Products"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnuHidden 
      Caption         =   "mnuHidden"
      Visible         =   0   'False
      Begin VB.Menu mnuProducts 
         Caption         =   "View Product"
      End
   End
End
Attribute VB_Name = "frmVendors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Loading As Boolean

Dim rsvendors As ADODB.Recordset
Dim oFields() As odbFields
Sub Setfield(sName As String, Indx As Integer, Value As Variant)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Setfield"
    Const ContainerName = "frmVendors"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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



    Select Case sName
    Case txtField(0).Name
        txtField(Indx) = Value
    Case Check1(0).Name
        Check1(Indx) = Val(Value)
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


Function Getfield(sName As String, Indx As Integer) As Variant

    '*[ Error Checking Variables ]**********************************************************************************
    Const RoutineName = "Main"
    Const ContainerName = "Globals"
    '***************************************************************************************************************


    '
    '***********************************************************************************************
    '**  Project Alpha ® 2003, 2004 +                                                             **
    '***********************************************************************************************
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

    Select Case sName
    Case txtField(0).Name
        Getfield = txtField(Indx)
    Case Check1(0).Name
        Getfield = IIf(Check1(Indx) = 0, 0, 1)
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

Sub Loadvendors()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Loadvendors"
    Const ContainerName = "frmVendors"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


       
    Dim sql1 As String
    Dim sql2 As String
    
    Dim ix As Byte
    sql1 = ""
    For ix = LBound(oFields) To UBound(oFields)
        Select Case oFields(ix).Encrypted
        Case True
            sql1 = sql1 + "AES_DECRYPT(`vendors`.`" & oFields(ix).FieldName & "`,'" & oFields(ix).Salt & "') as `" + oFields(ix).FieldName + "`, "
        Case False
            sql1 = sql1 + "`vendors`.`" & oFields(ix).FieldName & "`, "
        End Select
    Next ix

    sql1 = Left(sql1, Len(sql1) - 2)
'    MsgBox "select `RecID`, " & sql1 & " from vendors"
    
    Dim rsResult As ADODB.Recordset
    
    If MySQL.OpenTable(ADOConn, rsvendors, , MySQL.virtualisp("select distinct `vendors`.`RecID`, " & sql1 & " from vendors", "vendors", False, Login.bMaster)) = True Then
        
        
        
        Call MySQL.fillLV(ADOConn, rsvendors, lvVendors, False, , 1)
        
        
    End If
    
    cmFnc(1).Enabled = False

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

Private Sub Check1_Click(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Check1_Click"
    Const ContainerName = "frmVendors"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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

    cmFnc(1).Enabled = True
    
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


Public Sub SaveAddresses(lRecID As Variant)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "SaveAddresses"
    Const ContainerName = "frmVendors"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    
    If lvAddresses.ListItems.Count > 0 Then
    
        For sa = 1 To lvAddresses.ListItems.Count
            Set itmX = lvAddresses.ListItems(sa)
            If itmX.Key = "" Then

                On Error Resume Next
                Do
                    Err.Clear
                    RecID = MySQL.GetTMPRecID("vendors_addresses", ADOConn)
                    Call MySQL.Execute(ADOConn, "INSERT INTO vendors_addresses (RecID, VendorID, ContactName, Street1, Street2, Suburb, State, PostCode, Country, Checked) " + _
                            "VALUES ('" & RecID & "','" & lRecID & "','" & MySQL.ESC(itmX.Text) & "','" & MySQL.ESC(itmX.SubItems(1)) & "','" & MySQL.ESC(itmX.SubItems(2)) & "','" & MySQL.ESC(itmX.SubItems(3)) & "','" & MySQL.ESC(itmX.SubItems(4)) & "','" & MySQL.ESC(itmX.SubItems(5)) & "','" & MySQL.ESC(itmX.SubItems(6)) & "','" & IIf(itmX.Checked = True, -1, 0) & "')")

                    If Err.Number > 0 Then cDebug Err.Description
                Loop Until Err.Number = 0
                itmX.Key = "x" & RecID
                                
            ElseIf Left(itmX.Key, 1) = "e" Then
            
                
                
                Call MySQL.OpenTable(ADOConn, rsSave, , "select * from visp_addresses where RecID = " & Mid(itmX.Key, 2) & " Limit 1")
                
                Call MySQL.Execute(ADOConn, "update vendors_addresses set ContactName = '" & MySQL.ESC(itmX.Text) & " where RecID = " & Mid(itmX.Key, 2))
                Call MySQL.Execute(ADOConn, "update vendors_addresses set Street1 = '" & MySQL.ESC(itmX.SubItems(1)) & " where RecID = " & Mid(itmX.Key, 2))
                Call MySQL.Execute(ADOConn, "update vendors_addresses set Street2 = '" & MySQL.ESC(itmX.SubItems(2)) & " where RecID = " & Mid(itmX.Key, 2))
                Call MySQL.Execute(ADOConn, "update vendors_addresses set Suburb = '" & MySQL.ESC(itmX.SubItems(3)) & " where RecID = " & Mid(itmX.Key, 2))
                Call MySQL.Execute(ADOConn, "update vendors_addresses set State = '" & MySQL.ESC(itmX.SubItems(4)) & " where RecID = " & Mid(itmX.Key, 2))
                Call MySQL.Execute(ADOConn, "update vendors_addresses set PostCode = '" & MySQL.ESC(itmX.SubItems(5)) & " where RecID = " & Mid(itmX.Key, 2))
                Call MySQL.Execute(ADOConn, "update vendors_addresses set Country = '" & MySQL.ESC(itmX.SubItems(6)) & " where RecID = " & Mid(itmX.Key, 2))
                Call MySQL.Execute(ADOConn, "update vendors_addresses set Checked = '" & IIf(itmX.Checked = True, "-1", "0") & " where RecID = " & Mid(itmX.Key, 2))
                
                itmX.Key = "x" & Mid(itmX.Key, 2)
                Rem rsSave.Update
                
            End If
        Next
    
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

Public Sub SavePhone(lRecID As Variant)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "SavePhone"
    Const ContainerName = "frmVendors"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    
    If lvPhone.ListItems.Count > 0 Then
    
            For sa = 1 To lvPhone.ListItems.Count
            Set itmX = lvPhone.ListItems(sa)
            If itmX.Key = "" Then

                On Error Resume Next
                Do
                    Err.Clear
                    RecID = MySQL.GetTMPRecID("vendors_phone", ADOConn)
                    Call MySQL.Execute(ADOConn, "INSERT INTO vendors_phone (RecID, VendorID, ContactName, PhoneNumber, Extension, ShortNote, Checked) " + _
                            "VALUES ('" & RecID & "','" & lRecID & "','" & MySQL.ESC(itmX.Text) & "','" & MySQL.ESC(itmX.SubItems(1)) & "','" & MySQL.ESC(itmX.SubItems(2)) & "','" & MySQL.ESC(itmX.SubItems(3)) & "','" & IIf(itmX.Checked = True, -1, 0) & "')")

                    If Err.Number > 0 Then cDebug Err.Description
                Loop Until Err.Number = 0
                itmX.Key = "x" & RecID
                                
            ElseIf Left(itmX.Key, 1) = "e" Then
            
                
                
                Call MySQL.OpenTable(ADOConn, rsSave, , "select * from visp_addresses where RecID = " & Mid(itmX.Key, 2) & " Limit 1")
                
                Call MySQL.Execute(ADOConn, "update vendors_phone set ContactName = '" & MySQL.ESC(itmX.Text) & " where RecID = " & Mid(itmX.Key, 2))
                Call MySQL.Execute(ADOConn, "update vendors_phone set PhoneNumber = '" & MySQL.ESC(itmX.SubItems(1)) & " where RecID = " & Mid(itmX.Key, 2))
                Call MySQL.Execute(ADOConn, "update vendors_phone set Extension = '" & MySQL.ESC(itmX.SubItems(2)) & " where RecID = " & Mid(itmX.Key, 2))
                Call MySQL.Execute(ADOConn, "update vendors_phone set ShortNote = '" & MySQL.ESC(itmX.SubItems(3)) & " where RecID = " & Mid(itmX.Key, 2))
                Call MySQL.Execute(ADOConn, "update vendors_phone set Checked = '" & IIf(itmX.Checked = True, "-1", "0") & " where RecID = " & Mid(itmX.Key, 2))
                
                itmX.Key = "x" & Mid(itmX.Key, 2)
                Rem rsSave.Update
                
            End If
        Next
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

Public Sub SaveEmail(lRecID As Variant)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "SaveEmail"
    Const ContainerName = "frmVendors"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    
    If lveMail.ListItems.Count > 0 Then
    
        
            For sa = 1 To lveMail.ListItems.Count
            Set itmX = lveMail.ListItems(sa)
            If itmX.Key = "" Then

                On Error Resume Next
                Do
                    Err.Clear
                    RecID = MySQL.GetTMPRecID("vendors_email", ADOConn)
                    Call MySQL.Execute(ADOConn, "INSERT INTO vendors_email (RecID, VendorID, ContactName, Emailaddress, Checked) " + _
                            "VALUES ('" & RecID & "','" & lRecID & "','" & MySQL.ESC(itmX.Text) & "','" & MySQL.ESC(itmX.SubItems(1)) & "','" & IIf(itmX.Checked = True, -1, 0) & "')")

                    If Err.Number > 0 Then cDebug Err.Description
                Loop Until Err.Number = 0
                itmX.Key = "x" & RecID
                                
            ElseIf Left(itmX.Key, 1) = "e" Then
            
                
                
                Call MySQL.OpenTable(ADOConn, rsSave, , "select * from visp_addresses where RecID = " & Mid(itmX.Key, 2) & " Limit 1")
                
                Call MySQL.Execute(ADOConn, "update vendors_email set ContactName = '" & MySQL.ESC(itmX.Text) & " where RecID = " & Mid(itmX.Key, 2))
                Call MySQL.Execute(ADOConn, "update vendors_email set Emailaddress = '" & MySQL.ESC(itmX.SubItems(1)) & " where RecID = " & Mid(itmX.Key, 2))
                Call MySQL.Execute(ADOConn, "update vendors_email set Checked = '" & IIf(itmX.Checked = True, "-1", "0") & " where RecID = " & Mid(itmX.Key, 2))
                
                itmX.Key = "x" & Mid(itmX.Key, 2)
                Rem rsSave.Update
                
            End If
        Next
    
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


Private Sub cmFnc_Click(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmFnc_Click"
    Const ContainerName = "frmVendors"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    Dim rsSave As ADODB.Recordset
    Dim SQL As String
    Dim SQLa As String
    Dim SQLb As String
    
    Select Case Index
    Case 0 ' New Record
        
        
        For bx = txtField.LBound To txtField.UBound
            
            txtField(bx).Locked = False
        
        Next
        
        cmdAddAddress.Enabled = txtField(0).Locked
        cmdAddEmail.Enabled = txtField(0).Locked
        cmdAddPhone.Enabled = txtField(0).Locked
        cmFnc(0).Enabled = txtField(0).Locked
        cmFnc(1).Enabled = txtField(0).Locked
    
        If cmFnc(1).Enabled = True Then
            Select Case MsgBox("Do you wish to save the changes you have made?", vbYesNo + vbQuestion, "Save Changes")
            Case vbCancel
                Cancel = True
            Case vbYes
                Call cmFnc_Click(1)
            End Select
        End If
        
        lvVendors.Tag = ""
        For X = 0 To UBound(oFields)
            Setfield oFields(X).ControlName, oFields(X).ControlIndx, ""
        Next
        
        lveMail.ListItems.Clear
        lvPhone.ListItems.Clear
        lvAddresses.ListItems.Clear
        
        cmFnc(1).Enabled = False
    Case 1 ' Save all
        If lvVendors.Tag = "" Then
            
            If txtField(0) = "" Then
                frmAgent.oChar.GestureAt 400, 400
                frmAgent.oChar.Play "Decline"
                frmAgent.oChar.Speak "You must enter the vendors name first."
                Exit Sub
            End If
            
            If txtField(1) = "" Or txtField(2) = "" Then
                frmAgent.oChar.GestureAt 500, 500
                frmAgent.oChar.Play "Decline"
                frmAgent.oChar.Speak "You must enter the vendors ABN or ACN, you can enter both if you wish."
                Exit Sub
            End If
            
            
            SQL = ""
            SQLa = ""
            SQLb = ""
            
            For X = 0 To UBound(oFields)
                Select Case oFields(X).Encrypted
                Case True
                    SQL = SQL + "`" & oFields(X).FieldName & "` = AES_ENCRYPT('" & Getfield(oFields(X).ControlName, oFields(X).ControlIndx) & "','" & oFields(X).Salt & "'), "
                Case False
                    SQL = SQL + "`" & oFields(X).FieldName & "` = '" & Getfield(oFields(X).ControlName, oFields(X).ControlIndx) & "', "
                End Select
            Next X
            
            
            For X = 0 To UBound(oFields)
                Select Case oFields(X).Encrypted
                Case True
                    SQLa = SQLa + "`" & oFields(X).FieldName & "`, "
                    If IsNumeric(Getfield(oFields(X).ControlName, oFields(X).ControlIndx)) Then
                        SQLb = SQLb = "AES_ENCRYPT('" & Getfield(oFields(X).ControlName, oFields(X).ControlIndx) & "','" & oFields(X).Salt & "'), "
                    Else
                        If Getfield(oFields(X).ControlName, oFields(X).ControlIndx) = "" Then
                            SQLb = SQLb + "'', "
                        Else
                            SQLb = SQLb + "AES_ENCRYPT('" & Getfield(oFields(X).ControlName, oFields(X).ControlIndx) & "','" & oFields(X).Salt & "'), "
                        End If
                    End If
                Case False
                    SQLa = SQLa + "`" & oFields(X).FieldName & "`, "
                    If IsNumeric(Getfield(oFields(X).ControlName, oFields(X).ControlIndx)) Then
                        SQLb = SQLb + "'" & IIf(Getfield(oFields(X).ControlName, oFields(X).ControlIndx) = -1, 1, IIf(Getfield(oFields(X).ControlName, oFields(X).ControlIndx) = False, 0, Getfield(oFields(X).ControlName, oFields(X).ControlIndx))) & "', "
                    Else
                        If Getfield(oFields(X).ControlName, oFields(X).ControlIndx) = "" Then
                            SQLb = SQLb + "'', "
                        Else
                            SQLb = SQLb + "'" & Getfield(oFields(X).ControlName, oFields(X).ControlIndx) & "', "
                        End If
                    End If
                    
                End Select
            Next X
            
            'Debug.Print SQL
            'Debug.Print SQLa
            'Debug.Print SQLb
            'Stop
            
            
            Dim VendorID As Long
            Save = True
            If Save = True Then
            
                frmAgent.oChar.Play "Write"
                frmAgent.oChar.Speak "Saving and encrypting new vendor '" & txtField(0) & "' to the server."
                
                On Error Resume Next
                Do
                    Err.Clear
                    VendorID = MySQL.GetTMPRecID("vendors", ADOConn)
                    lvVendors.Tag = VendorID
                    ADOConn.Execute "insert into vendors (`RecID`," & Left(SQLa, Len(SQLa) - 2) & ",VirtualID, SysopID) VALuES('" & VendorID & "'," & Left(SQLb, Len(SQLb) - 2) & ",'" & Login.lVirtualID & "','" & Login.lSysopID & "')"
                    
                Loop Until Err.Number = 0
                    
                Call Form_Load
                
            End If
            
            SavePhone VendorID
            SaveAddresses VendorID
            SaveEmail VendorID
            
        Else
            
            frmAgent.oChar.Play "Write"
            frmAgent.oChar.Speak "Saving and encrypting data to the server."
            
            For X = 0 To UBound(oFields)
                Select Case oFields(X).Encrypted
                Case True
                    SQL = SQL + "`" & oFields(X).FieldName & "` = AES_ENCRYPT('" & Getfield(oFields(X).ControlName, oFields(X).ControlIndx) & "','" & oFields(X).Salt & "'), "
                Case False
                    SQL = SQL + "`" & oFields(X).FieldName & "` = '" & Getfield(oFields(X).ControlName, oFields(X).ControlIndx) & "', "
                End Select
            Next X
                    
            MySQL.Execute ADOConn, "update vendors set " & Left(SQL, Len(SQL) - 2) & " where RecID = " & lvVendors.Tag
                    
            SavePhone lvVendors.Tag
            SaveAddresses lvVendors.Tag
            SaveEmail lvVendors.Tag
            
        
        End If
        
        Dim itmX As ListItem
        
        If Save = True Then
            Set itmX = lvVendors.ListItems.Add(, "v" & VendorID, VendorID)
        Else
            Set itmX = lvVendors.SelectedItem
        End If
        
        
        For X = 2 To lvVendors.ColumnHeaders.Count
            For ix = 0 To UBound(oFields)
                If lvVendors.ColumnHeaders(X).Tag = oFields(ix).FieldName Then
                    itmX.SubItems(X - 1) = Getfield(oFields(ix).ControlName, oFields(ix).ControlIndx)
                    Exit For
                End If
            Next
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

Private Sub lvAddresses_ItemCheck(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvAddresses_ItemCheck"
    Const ContainerName = "frmVendors"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    cmFnc(1).Enabled = True
    
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

Private Sub lvAddresses_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvAddresses_ItemClick"
    Const ContainerName = "frmVendors"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    lvAddresses.Tag = True
    
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

Private Sub lvEmail_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvEmail_ItemClick"
    Const ContainerName = "frmVendors"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    lveMail.Tag = True
    
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

Private Sub lvPhone_DblClick()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvPhone_DblClick"
    Const ContainerName = "frmVendors"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
            itmX.Text = ffrmPhoneNo.sContactName
            itmX.SubItems(1) = ffrmPhoneNo.sPhonenumber
            itmX.SubItems(2) = ffrmPhoneNo.sExtension
            itmX.SubItems(3) = ffrmPhoneNo.sNote
            If itmX.Key <> "" Then itmX.Key = "e" & Mid(itmX.Key, 2)
            cmFnc(1).Enabled = True
            
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
Private Sub lvAddresses_DblClick()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvAddresses_DblClick"
    Const ContainerName = "frmVendors"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
            
            itmX.Text = ffrmSnailMail.sContactName
            itmX.SubItems(1) = ffrmSnailMail.sStreetLine1
            itmX.SubItems(2) = ffrmSnailMail.sStreetLine2
            itmX.SubItems(3) = ffrmSnailMail.sSuburb
            itmX.SubItems(4) = ffrmSnailMail.sState
            itmX.SubItems(5) = ffrmSnailMail.sPostcode
            itmX.SubItems(6) = ffrmSnailMail.sCountry
            If itmX.Key <> "" Then itmX.Key = "e" & Mid(itmX.Key, 2)
            cmFnc(1).Enabled = True
            
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
Private Sub lvEmail_DblClick()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvEmail_DblClick"
    Const ContainerName = "frmVendors"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    If lveMail.Tag <> "" Then
    
        Dim ffrmEmail As frmEmail
        Set ffrmEmail = New frmEmail
        ffrmEmail.sContactName = lveMail.SelectedItem.Text
        ffrmEmail.sEmailAddress = lveMail.SelectedItem.SubItems(1)
        ffrmEmail.Show 1
        
        If ffrmEmail.iCloseState = frmCloseSave Then
            Dim itmX As ListItem
            Set itmX = lveMail.SelectedItem
            itmX.Text = ffrmEmail.sContactName
            itmX.SubItems(1) = ffrmEmail.sEmailAddress
            If itmX.Key <> "" Then itmX.Key = "e" & Mid(itmX.Key, 2)
            cmFnc(1).Enabled = True
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

Private Sub lvEmail_ItemCheck(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvEmail_ItemCheck"
    Const ContainerName = "frmVendors"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    cmFnc(1).Enabled = True
    
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
    Const ContainerName = "frmVendors"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    Call GUI.LoadColWidths(lvVendors, Me)
    Call GUI.LoadColWidths(lvAddresses, Me)
    Call GUI.LoadColWidths(lveMail, Me)
    Call GUI.LoadColWidths(lvPhone, Me)

    Dim ix As Byte
    Dim Count As Byte
    ReDim oFields(txtField.UBound + 1 + Check1.UBound) As odbFields
    For ix = txtField.LBound To txtField.UBound
        
        
        Select Case Left(txtField(ix).Tag, 1)
        Case "c"
        
            oFields(Count).Encrypted = True
            oFields(Count).Salt = odb.colSalts.ReturnSalt("VendorSalt")
            oFields(Count).FieldName = Mid(txtField(ix).Tag, 2)
            oFields(Count).ControlName = txtField(ix).Name
            oFields(Count).ControlIndx = ix
        
        Case Else
        
            oFields(Count).Encrypted = False
            oFields(Count).Salt = odb.colSalts.ReturnSalt("VendorSalt")
            oFields(Count).FieldName = txtField(ix).Tag
            oFields(Count).ControlName = txtField(ix).Name
            oFields(Count).ControlIndx = ix
        
        
        End Select
        
        Count = Count + 1
    Next
    
    For ix = Check1.LBound To Check1.UBound
        Select Case Left(Check1(ix).Tag, 1)
        Case "c"
        
            oFields(Count).Encrypted = True
            oFields(Count).Salt = odb.colSalts.ReturnSalt("VendorSalt")
            oFields(Count).FieldName = Mid(txtField(ix).Tag, 2)
            oFields(Count).ControlName = Check1(ix).Name
            oFields(Count).ControlIndx = ix
        
        Case Else
        
            oFields(Count).Encrypted = False
            oFields(Count).Salt = odb.colSalts.ReturnSalt("VendorSalt")
            oFields(Count).FieldName = Check1(ix).Tag
            oFields(Count).ControlName = Check1(ix).Name
            oFields(Count).ControlIndx = ix
        
        End Select
        
        Count = Count + 1
    Next
    
    Dim bx As Byte
    
    For bx = txtField.LBound To txtField.UBound
        
        txtField(bx).Locked = IIf(Login.lLevel >= 90, True, False)
    
    Next
    
    cmdAddAddress.Enabled = IIf(Login.lLevel >= 90, True, False)
    cmdAddEmail.Enabled = IIf(Login.lLevel >= 90, True, False)
    cmdAddPhone.Enabled = IIf(Login.lLevel >= 90, True, False)
    cmFnc(0).Enabled = IIf(Login.lLevel >= 90, True, False)
    cmFnc(1).Enabled = IIf(Login.lLevel >= 90, True, False)
    
    Loadvendors
    
    'cmFnc(1).Enabled = False
    
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_QueryUnload"
    Const ContainerName = "frmVendors"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    If cmFnc(1).Enabled = True Then
        Select Case MsgBox("Do you wish to save the changes you have made?", vbYesNoCancel + vbQuestion, "Save Changes")
        Case vbCancel
            Cancel = True
        Case vbYes
            Call cmFnc_Click(1)
        End Select
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

Private Sub Form_Resize()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Resize"
    Const ContainerName = "frmVendors"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    Dim plane As Single
    
    plane = Me.ScaleHeight - 480
    If plane > 500 And Me.ScaleWidth > 1700 Then
    
        lvVendors.Move 60, 60, Me.ScaleWidth - 120, plane / 10 * 4
        ts.Move 60, lvVendors.Top + lvVendors.Height + 120, Me.ScaleWidth - 120, Me.ScaleHeight - lvVendors.Height - 240
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

Private Sub Form_Unload(Cancel As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Unload"
    Const ContainerName = "frmVendors"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    Call GUI.SaveColWidths(lvVendors, Me)
    Call GUI.SaveColWidths(lvAddresses, Me)
    Call GUI.SaveColWidths(lveMail, Me)
    Call GUI.SaveColWidths(lvPhone, Me)
    
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




Private Sub lvPhone_ItemCheck(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvPhone_ItemCheck"
    Const ContainerName = "frmVendors"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    cmFnc(1).Enabled = True
    
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

Private Sub lvPhone_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvPhone_ItemClick"
    Const ContainerName = "frmVendors"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    lvPhone.Tag = True
    
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

Private Sub lvvendors_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvvendors_ColumnClick"
    Const ContainerName = "frmVendors"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    GUI.ColumnSort ColumnHeader, lvVendors
    
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

Private Sub lvvendors_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvvendors_ItemClick"
    Const ContainerName = "frmVendors"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    Loading = True
    
    Dim rsload As ADODB.Recordset
    
    If cmFnc(1).Enabled = True Then
        Select Case MsgBox("Do you wish to save the changes you have made?", vbYesNo + vbQuestion, "Save Changes")
        Case vbCancel
            Cancel = True
        Case vbYes
            Call cmFnc_Click(1)
        End Select
    End If
    
    Dim ix As Byte
    sql1 = ""
    For ix = LBound(oFields) To UBound(oFields)
        Select Case oFields(ix).Encrypted
        Case True
            sql1 = sql1 + "AES_DECRYPT(`" & oFields(ix).FieldName & "`,'" & oFields(ix).Salt & "') as `" + oFields(ix).FieldName + "`, "
        Case False
            sql1 = sql1 + "`" & oFields(ix).FieldName & "`, "
        End Select
    Next ix

    sql1 = Left(sql1, Len(sql1) - 2)
    
    Dim rsResult As ADODB.Recordset
    
    If MySQL.OpenTable(ADOConn, rsload, , "select RecID, " & sql1 & " from vendors where RecID = " & Mid(Item.Key, 2)) = True Then
   
    
        For X = LBound(oFields) To UBound(oFields)
                        
            
            Select Case MySQL.fldType(rsload(oFields(X).FieldName).Type)
            Case "Tiny Integer"
                If IsNull(rsload(oFields(X).FieldName)) Then
                    Setfield oFields(X).ControlName, oFields(X).ControlIndx, ""
                Else
                    Setfield oFields(X).ControlName, oFields(X).ControlIndx, IIf(Val(rsload(oFields(X).FieldName)) = -1, 1, Val(rsload(oFields(X).FieldName)))
                End If
            Case Else
                Setfield oFields(X).ControlName, oFields(X).ControlIndx, IIf(IsNull(rsload(oFields(X).FieldName)), "", rsload(oFields(X).FieldName))
            End Select
            
        Next
    
    End If
    
    lvVendors.Tag = Mid(Item.Key, 2)
    
    
    lvAddresses.ListItems.Clear
    lvPhone.ListItems.Clear
    lveMail.ListItems.Clear
    
    lvAddresses.Tag = ""
    lvPhone.Tag = ""
    lveMail.Tag = ""
    
    On Error Resume Next
    
    
    bResult = MySQL.OpenTable(ADOConn, rsload, , "select * from vendors_addresses Where VendorID = " & Mid(Item.Key, 2))
    
    If rsload.RecordCount > 0 Then
        rsload.MoveFirst
        While Not rsload.EOF And Err.Number = 0
            Set itmX = lvAddresses.ListItems.Add(, "r" & rsload!RecID, rsload!ContactName)
            itmX.SubItems(1) = rsload!Street1
            itmX.SubItems(2) = rsload!Street2
            itmX.SubItems(3) = rsload!Suburb
            itmX.SubItems(4) = rsload!State
            itmX.SubItems(5) = rsload!Postcode
            itmX.SubItems(6) = rsload!Country
            itmX.Checked = IIf(rsload!Checked <> 0, True, False)
            rsload.MoveNext
        Wend
    End If

    bResult = MySQL.OpenTable(ADOConn, rsload, , "select * from vendors_email Where VendorID = " & Mid(Item.Key, 2))
    
    If rsload.RecordCount > 0 Then
        rsload.MoveFirst
        While Not rsload.EOF And Err.Number = 0
            Set itmX = lveMail.ListItems.Add(, "r" & rsload!RecID, rsload!ContactName)
            itmX.SubItems(1) = rsload!EmailAddress
            itmX.Checked = IIf(rsload!Checked <> 0, True, False)
            rsload.MoveNext
        Wend
    End If

    bResult = MySQL.OpenTable(ADOConn, rsload, , "select * from vendors_phone Where VendorID = " & Mid(Item.Key, 2))
    
    If rsload.RecordCount > 0 Then
        rsload.MoveFirst
        While Not rsload.EOF And Err.Number = 0
            Set itmX = lvPhone.ListItems.Add(, "r" & rsload!RecID, rsload!ContactName)
            itmX.SubItems(1) = rsload!PhoneNumber
            itmX.SubItems(2) = rsload!Extension
            itmX.SubItems(3) = rsload!ShortNote
            itmX.Checked = IIf(rsload!Checked <> 0, True, False)
            rsload.MoveNext
        Wend
    End If

    cmFnc(1).Enabled = False
    rsload.Filter = ""
    
    Loading = False
    
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

Private Sub lvVendors_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvVendors_MouseDown"
    Const ContainerName = "frmVendors"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    If lvVendors.SelectedItem Is Nothing Then Exit Sub
    
    If Login.bMaster = False Then Exit Sub
    
    If Button = 2 Then PopupMenu mnuHidden
    
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

Private Sub mnuProducts_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "mnuProducts_Click"
    Const ContainerName = "frmVendors"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    Dim fmProd As New frmProducts
    
    fmProd.VendorID = Val(Mid(lvVendors.SelectedItem.Key, 2))
    fmProd.Show 1
    
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

Private Sub picTS_Resize(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "picTS_Resize"
    Const ContainerName = "frmVendors"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
        lvPhone.Move 60, 60, picTS(Index).ScaleWidth - 120, picTS(Index).ScaleHeight - 180 - cmdAddPhone.Height
        lvPhone.Refresh
        cmdAddPhone.Move 60, lvPhone.Top + lvPhone.Height + 60
    Case 2
        lveMail.Move 60, 60, picTS(Index).ScaleWidth - 120, picTS(Index).ScaleHeight - 180 - cmdAddPhone.Height
        lveMail.Refresh
        cmdAddEmail.Move 60, lveMail.Top + lveMail.Height + 60
    Case 3
        lvAddresses.Move 60, 60, picTS(Index).ScaleWidth - 120, picTS(Index).ScaleHeight - 180 - cmdAddPhone.Height
        lvAddresses.Refresh
        cmdAddAddress.Move 60, lvAddresses.Top + lvAddresses.Height + 60

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
    Const ContainerName = "frmVendors"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    
    For ix = 0 To picTS.UBound
        picTS(ix).Visible = False
    Next ix
    
    picTS(ts.SelectedItem.Index - 1).Move ts.clientLeft, ts.clientTop, ts.clientWidth, ts.clientHeight
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

Private Sub cmdAddAddress_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdAddAddress_Click"
    Const ContainerName = "frmVendors"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    Dim ffrmSnailMail As frmSnailMail
    Set ffrmSnailMail = New frmSnailMail
    ffrmSnailMail.Show 1
    
    If ffrmSnailMail.iCloseState = frmCloseSave Then
        Dim itmX As ListItem
        Set itmX = lvAddresses.ListItems.Add(, , ffrmSnailMail.sContactName)
        itmX.SubItems(1) = ffrmSnailMail.sStreetLine1
        itmX.SubItems(2) = ffrmSnailMail.sStreetLine2
        itmX.SubItems(3) = ffrmSnailMail.sSuburb
        itmX.SubItems(4) = ffrmSnailMail.sState
        itmX.SubItems(5) = ffrmSnailMail.sPostcode
        itmX.SubItems(6) = ffrmSnailMail.sCountry
        'If lRecID = 0 Then SaveInformation Else Call SaveAddresses(lRecID)
        cmFnc(1).Enabled = True
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

Private Sub cmdAddEmail_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdAddEmail_Click"
    Const ContainerName = "frmVendors"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    Dim ffrmEmail As frmEmail
    Set ffrmEmail = New frmEmail
    If lveMail.ListItems.Count = 0 Then
        ffrmEmail.sContactName = txtField(0).Text
        ffrmEmail.sEmailAddress = "@" & txtField(1).Text
    End If
    ffrmEmail.Show 1
    
    If ffrmEmail.iCloseState = frmCloseSave Then
        Dim itmX As ListItem
        Set itmX = lveMail.ListItems.Add(, , ffrmEmail.sContactName)
        itmX.SubItems(1) = ffrmEmail.sEmailAddress
        'If lRecID = 0 Then SaveInformation Else Call SaveEmail(lRecID)
        cmFnc(1).Enabled = True
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



Private Sub cmdAddPhone_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdAddPhone_Click"
    Const ContainerName = "frmVendors"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    Dim ffrmPhoneNo As frmPhoneNumber
    Set ffrmPhoneNo = New frmPhoneNumber
    If lvPhone.ListItems.Count = 0 Then
        ffrmPhoneNo.sContactName = txtField(0).Text
    End If
    
    ffrmPhoneNo.Show 1
    
    If ffrmPhoneNo.iCloseState = frmCloseSave Then
        Dim itmX As ListItem
        Set itmX = lvPhone.ListItems.Add(, , ffrmPhoneNo.sContactName)
        itmX.SubItems(1) = ffrmPhoneNo.sPhonenumber
        itmX.SubItems(2) = ffrmPhoneNo.sExtension
        itmX.SubItems(3) = ffrmPhoneNo.sNote
        cmFnc(1).Enabled = True
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


Private Sub txtfield_Change(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtfield_Change"
    Const ContainerName = "frmVendors"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    If Loading = False Then cmFnc(1).Enabled = True Else cmFnc(1).Enabled = False
    
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
