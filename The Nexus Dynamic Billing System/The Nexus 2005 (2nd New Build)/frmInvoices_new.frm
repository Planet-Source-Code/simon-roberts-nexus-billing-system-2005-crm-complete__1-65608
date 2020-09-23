VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmStatements 
   BackColor       =   &H00D2D3A7&
   Caption         =   "Invoices and Payment History"
   ClientHeight    =   8745
   ClientLeft      =   2145
   ClientTop       =   4305
   ClientWidth     =   10425
   Icon            =   "frmInvoices_new.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8745
   ScaleWidth      =   10425
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picResize 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   7245
      Left            =   3270
      MousePointer    =   9  'Size W E
      ScaleHeight     =   7245
      ScaleWidth      =   75
      TabIndex        =   4
      Top             =   870
      Width           =   75
   End
   Begin VB.PictureBox picTreeView 
      BackColor       =   &H00D2D3A7&
      BorderStyle     =   0  'None
      Height          =   7245
      Left            =   0
      ScaleHeight     =   7245
      ScaleWidth      =   3255
      TabIndex        =   2
      Top             =   870
      Width           =   3255
      Begin VB.ComboBox cmbDisplay 
         BackColor       =   &H00BA3F3F&
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
         Height          =   420
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   60
         Width           =   3135
      End
      Begin MSComctlLib.TreeView tvInvoices 
         Height          =   6375
         Left            =   60
         TabIndex        =   3
         Top             =   510
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   11245
         _Version        =   393217
         Indentation     =   441
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         HotTracking     =   -1  'True
         ImageList       =   "ilTreeview"
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
      Begin MSComctlLib.ProgressBar pb2 
         Height          =   195
         Left            =   60
         TabIndex        =   6
         Top             =   6930
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin VB.PictureBox picListview 
      BackColor       =   &H00D2D3A7&
      BorderStyle     =   0  'None
      Height          =   7245
      Index           =   0
      Left            =   3360
      ScaleHeight     =   7245
      ScaleWidth      =   6945
      TabIndex        =   0
      Top             =   840
      Width           =   6945
      Begin VB.CommandButton Command1 
         BackColor       =   &H00D2D3A7&
         Caption         =   "Group Payment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4650
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   90
         Width           =   2295
      End
      Begin VB.CommandButton cmdPayment 
         BackColor       =   &H00D2D3A7&
         Caption         =   "Make Payment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4620
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   690
         Width           =   2295
      End
      Begin VB.PictureBox picTS 
         BorderStyle     =   0  'None
         Height          =   2985
         Index           =   1
         Left            =   4950
         ScaleHeight     =   2985
         ScaleWidth      =   2715
         TabIndex        =   18
         Top             =   2160
         Visible         =   0   'False
         Width           =   2715
         Begin MSComctlLib.ListView lvPaymentMade 
            Height          =   1605
            Left            =   30
            TabIndex        =   19
            Top             =   90
            Width           =   7125
            _ExtentX        =   12568
            _ExtentY        =   2831
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
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "When"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Amount"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "GST"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Subcharge"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.PictureBox picTS 
         BorderStyle     =   0  'None
         Height          =   2865
         Index           =   0
         Left            =   270
         ScaleHeight     =   2865
         ScaleWidth      =   6525
         TabIndex        =   16
         Top             =   2130
         Width           =   6525
         Begin MSComctlLib.ListView lvItems 
            Height          =   2445
            Left            =   120
            TabIndex        =   17
            Top             =   90
            Width           =   6405
            _ExtentX        =   11298
            _ExtentY        =   4313
            View            =   3
            Arrange         =   1
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
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Description"
               Object.Width           =   4057
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Debited (Ex Tax)"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "GST Debited"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Total Debit"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Due When"
               Object.Width           =   3704
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Amount Paid"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Credited"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin MSComctlLib.TabStrip ts 
         Height          =   3765
         Left            =   120
         TabIndex        =   15
         Top             =   1590
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   6641
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Statement"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Payments"
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
      Begin VB.TextBox txtComment 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   150
         Locked          =   -1  'True
         MaxLength       =   65534
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         ScrollBars      =   3  'Both
         TabIndex        =   14
         Top             =   5370
         Width           =   6735
      End
      Begin VB.CheckBox chkFinalised 
         BackColor       =   &H00D2D3A7&
         Caption         =   "Finalised"
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
         Height          =   375
         Left            =   4620
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1080
         Width           =   2295
      End
      Begin MSComctlLib.ProgressBar PB1 
         Height          =   195
         Left            =   30
         TabIndex        =   1
         Top             =   6990
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Line Line3 
         X1              =   3090
         X2              =   3090
         Y1              =   390
         Y2              =   1200
      End
      Begin VB.Shape Shape1 
         Height          =   1155
         Left            =   30
         Top             =   390
         Width           =   4425
      End
      Begin VB.Line Line2 
         Index           =   2
         X1              =   60
         X2              =   4440
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   60
         X2              =   4440
         Y1              =   930
         Y2              =   930
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   60
         X2              =   4440
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Credited:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   1635
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$ 0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   4
         Left            =   3780
         TabIndex        =   24
         Top             =   930
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   22
         Top             =   60
         Width           =   735
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   960
         TabIndex        =   21
         Top             =   30
         Width           =   4395
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$ 0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   270
         Index           =   0
         Left            =   3780
         TabIndex        =   13
         Top             =   390
         Width           =   615
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$ 0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   270
         Index           =   1
         Left            =   3780
         TabIndex        =   12
         Top             =   660
         Width           =   615
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000 00:00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   2070
         TabIndex        =   11
         Top             =   1200
         Width           =   2325
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Due:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1545
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Paid:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   8
         Top             =   690
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount to Debit (future):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   7
         Top             =   390
         Width           =   2355
      End
   End
   Begin VB.Timer mTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   7020
      Top             =   240
   End
   Begin MSComctlLib.ImageList ilTreeview 
      Left            =   4950
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   45
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":0BE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":14BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":1D98
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":2672
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":2F4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":3826
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":4100
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":49DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":52B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":5B8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":6468
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":6D42
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":761C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":7EF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":87D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":90AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":9984
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":A25E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":AB38
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":B412
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":BCEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":C5C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":CEA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":D77A
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":E054
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":E92E
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":F208
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":FAE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":103BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":10C96
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":11570
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":11E4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":12724
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":12FFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":138D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":141B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":14A8C
            Key             =   "book"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":14DA6
            Key             =   "news"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":150C0
            Key             =   "filter"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":153DA
            Key             =   "world"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":156F4
            Key             =   "Finalised"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":15B46
            Key             =   "Unfinalised"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":15F98
            Key             =   "Overdue"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoices_new.frx":163EA
            Key             =   "Partially"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Register"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F3B4E3&
      Height          =   510
      Left            =   420
      TabIndex        =   26
      Top             =   90
      Width           =   3690
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   -510
      Picture         =   "frmInvoices_new.frx":16B3C
      Stretch         =   -1  'True
      Top             =   -480
      Width           =   2295
   End
End
Attribute VB_Name = "frmStatements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mButtonDown As Boolean
Dim iLastPoint As POINTAPI
Dim LastMovement As POINTAPI
Dim nodeP As Node
Dim bChanged As Boolean


Private Sub cmbDisplay_Change()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmbDisplay_Change"
    Const ContainerName = "frmStatements"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    Dim NodX As Node
    Dim NodeX As Node
    Dim rsload As adodb.Recordset
    Dim bResult As Boolean
    
    Select Case cmbDisplay.ListIndex
    Case 0
        bResult = MySQL.OpenTable(directConn, rsload, , MySQL.virtualisp("select distinct accountinfo.AccountName ,invoicetraxr.* from accountinfo, invoicetraxr Where invoicetraxr.acci_RecID = accountinfo.RecID and invoicetraxr.Finalised = 0 Order By accountinfo.AccountName DESC", "accountinfo", True) + " ")
    Case 1
        bResult = MySQL.OpenTable(directConn, rsload, , MySQL.virtualisp("select distinct accountinfo.AccountName ,invoicetraxr.* from accountinfo, invoicetraxr Where invoicetraxr.acci_RecID = accountinfo.RecID and invoicetraxr.Finalised = 0 Order By accountinfo.AccountName DESC", "accountinfo", True) + " ")
    Case 2
        bResult = MySQL.OpenTable(directConn, rsload, , MySQL.virtualisp("select distinct accountinfo.AccountName ,invoicetraxr.* from accountinfo, invoicetraxr Where invoicetraxr.acci_RecID = accountinfo.RecID and invoicetraxr.Finalised <> 0 Order By accountinfo.AccountName DESC", "accountinfo", True) + " ")
    Case 3
        bResult = MySQL.OpenTable(directConn, rsload, , MySQL.virtualisp("select distinct accountinfo.AccountName ,invoicetraxr.* from accountinfo, invoicetraxr Where invoicetraxr.acci_RecID = accountinfo.RecID Order By accountinfo.AccountName DESC", "accountinfo", True) + " ")
    End Select
    
    tvInvoices.NodeS.Clear
    
    Dim acci_RecID As Variant
    
    If rsload.State <> adStateOpen Then Exit Sub
    
    
    If rsload.RecordCount > 0 Then
        pb2.Value = 0
        pb2.Max = rsload.RecordCount
        While Not rsload.EOF And Err.Number = 0
            Select Case cmbDisplay.ListIndex
            Case 0, 2, 3
                
                If acci_RecID <> rsload!acci_RecID Then
                    Set NodeX = tvInvoices.NodeS.Add(, , "k" & rsload!acci_RecID, IIf(IsNull(rsload!AccountName), "(null)", rsload!AccountName), 3, 7)
                    acci_RecID = rsload!acci_RecID
                End If
                Set NodX = tvInvoices.NodeS.Add(NodeX.Key, tvwChild, "i" & rsload!RecID, rsload!RecID & " / " & rsload!InvoiceSerial + " - [" & Format(rsload!TotalDue - (rsload!AmountPaid + rsload!AmountCredited), "Currency") & "]")
                NodX.Tag = rsload!AccountName
                If rsload!TotalDue - (rsload!AmountPaid + rsload!AmountCredited) = 0 Then
                    MySQL.Execute directConn, "UPDATE invoicetraxr SET Finalised = -1 where RecID = " & rsload!RecID
                    'sLoad!Finalised = True
                    NodX.Image = 42
                Else
                    NodX.Image = IIf(Val(rsload!Finalised) = 0, "Unfinalised", 42)
                End If
                
                NodX.Image = IIf(Val(rsload!Finalised) = 0, "Unfinalised", 42)
                If rsload!AmountPaid <> 0 And rsload!AmountPaid < rsload!TotalDue Then
                    NodX.Image = "Partially"
                End If
                If rsload!AmountPaid < rsload!TotalDue And DateDiff("s", rsload!PaymentDue, sysnow) > 0 Then NodX.Image = "Overdue"

            Case 1
                
                Set NodX = tvInvoices.NodeS.Add(, , "i" & rsload!RecID, rsload!RecID & " / " & rsload!InvoiceSerial + " - [" & Format(rsload!TotalDue - (rsload!AmountPaid + rsload!AmountCredited), "Currency") & "]")
                If rsload!TotalDue - (rsload!AmountPaid + rsload!AmountCredited) = 0 Then
                    MySQL.Execute directConn, "UPDATE invoicetraxr SET Finalised = -1 where RecID = " & rsload!RecID
                    'sLoad!Finalised = True
                    NodX.Image = 42
                Else
                    NodX.Image = IIf(Val(rsload!Finalised) = 0, "Unfinalised", 42)
                End If
                
                NodX.Tag = rsload!AccountName
                If rsload!AmountPaid <> 0 And rsload!AmountPaid < rsload!TotalDue Then
                    NodX.Image = "Partially"
                End If
                If rsload!AmountPaid < rsload!TotalDue And DateDiff("s", rsload!PaymentDue, sysnow) > 0 Then NodX.Image = "Overdue"
                
            End Select
            
            
            rsload.MoveNext
            pb2.Value = pb2.Value + 1
        Wend
    End If
    
    pb2.Value = pb2.Max
    
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

Private Sub cmbDisplay_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmbDisplay_Click"
    Const ContainerName = "frmStatements"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    Call cmbDisplay_Change
    
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

Private Sub cmdPayment_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdPayment_Click"
    Const ContainerName = "frmStatements"
    '***************************************************************************************************************


    If chkFinalised.Value = False Then
    
    Else
        Exit Sub
    End If
'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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



    If Not nodeP Is Nothing Then
        Dim frmPayment As frmInvPayment
        Set frmPayment = New frmInvPayment
        Dim rsload As adodb.Recordset
        
        frmPayment.l_RecID = Val(Mid(nodeP.Key, 2))
        Call MySQL.OpenTable(directConn, rsload, , "select acci_RecID from invoicetraxr where RecID = " & frmPayment.l_RecID)
        frmPayment.acci_RecID = Val(rsload!acci_RecID)
        frmPayment.s_AccountName = lblField(3).Caption
        frmPayment.c_TotalDue = CCur(Mid(lblField(0).Caption, 2))
        frmPayment.c_TotalPaid = CCur(Mid(lblField(1).Caption, 2))
        frmPayment.Show 1
                
        Call tvInvoices_NodeClick(nodeP)
        'lvReceivables.selectedItem.SubItems(5) = Format(frmPayment.c_TotalPaid, "Currency")
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

Private Sub Command1_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Command1_Click"
    Const ContainerName = "frmStatements"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    Dim fGroup As New frmGroup
    
    fGroup.Show 1
    
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
    Const ContainerName = "frmStatements"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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

    If bBigFont = True Then
    
        lvItems.Font.Size = 16.5
        lvPaymentMade.Font.Size = 16.5
        txtComment.FontSize = 18
        tvInvoices.Font.Size = 17
        
    End If

    LoadColumns
    
    cmdPayment.Visible = Login.bMaster
    Command1.Visible = Login.bVISP
    
    cmbDisplay.AddItem "Sort By Customer (Pending)"
    cmbDisplay.AddItem "Sort By Invoice Number (Pending)"
    cmbDisplay.AddItem "Sort by Finalised"
    cmbDisplay.AddItem "View All"
    
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
    Const ContainerName = "frmStatements"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    SaveColumns
    Dim rsload As adodb.Recordset
    Dim bResult As Boolean
    
    If bChanged = True Then
        bResult = MySQL.OpenTable(directConn, rsload, , "select * from invoicetraxr Where RecID = " & Mid(nodeP.Key, 2))
        rsload!Comment = txtComment
        rsload.Update
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
    Const ContainerName = "frmStatements"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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



    If Me.ScaleHeight > 900 And Me.ScaleWidth > 3000 Then
        'tsAcc.Move 60, tsAcc.Top, Me.ScaleWidth - 180, Me.ScaleHeight - 180 - tsAcc.Top
        
        'picListview(0).Move
        picTreeView.Height = Me.ScaleHeight - picTreeView.Top - 60
        picListview(0).Width = IIf(Me.Width - picResize.Left - picResize.Width - 120 < 0, 10, Me.Width - picResize.Left - picResize.Width - 120)
        picListview(0).Height = Me.ScaleHeight - picListview(0).Top - 60
        picResize.Height = Me.ScaleHeight - picResize.Top
        'Line1.X2 = Me.ScaleWidth
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

Private Sub lblField_DblClick(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lblField_DblClick"
    Const ContainerName = "frmStatements"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    Case 2
        If IsDate(lblField(2)) = True Then
            Dim fDate As New frmDate
            fDate.dDate = CDate(lblField(2))
            fDate.Show 1
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

Private Sub lvItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvItems_ColumnClick"
    Const ContainerName = "frmStatements"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    Call GUI.ColumnSort(ColumnHeader, lvItems)
    
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

Private Sub lvPaymentMade_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvPaymentMade_ColumnClick"
    Const ContainerName = "frmStatements"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    Call GUI.ColumnSort(ColumnHeader, lvPaymentMade)
    
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

Private Sub lvStatementItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvStatementItems_ColumnClick"
    Const ContainerName = "frmStatements"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    'Call GUI.ColumnSort(ColumnHeader, lvStatementItems)
    
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
    Const ContainerName = "frmStatements"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    
    picTreeView.Width = picResize.Left
    picListview(0).Width = IIf(Me.Width - picResize.Left - picResize.Width - 120 < 0, 10, Me.Width - picResize.Left - picResize.Width - 120)
    picListview(0).Left = picResize.Left + picResize.Width
    
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

Private Sub picListview_Resize(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "picListview_Resize"
    Const ContainerName = "frmStatements"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    If picListview(0).Height > 3200 And picListview(0).Width > 500 Then
        ts.Move 60, ts.Top, picListview(0).Width - 120, picListview(0).Height - pb1.Height - 120 - ts.Top - txtComment.Height - 60
        txtComment.Move 60, ts.Top + ts.Height + 60, picListview(0).Width - 120
        
        pb1.Move 60, ts.Top + ts.Height + 60, picListview(0).Width - 120
        chkFinalised.Left = picListview(0).Width - chkFinalised.Width - 60
        cmdPayment.Left = picListview(0).Width - chkFinalised.Width - 60
        Command1.Left = cmdPayment.Left
        
        If ts.SelectedItem.Index - 1 <= picTS.UBound Then
            picTS(ts.SelectedItem.Index - 1).Move ts.ClientLeft, ts.ClientTop, ts.ClientWidth, ts.ClientHeight
            picTS(ts.SelectedItem.Index - 1).Visible = True
            picTS(ts.SelectedItem.Index - 1).ZOrder 0
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

Private Sub picResize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "picResize_MouseDown"
    Const ContainerName = "frmStatements"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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



Public Function SaveColumns()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "SaveColumns"
    Const ContainerName = "frmStatements"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    Call GUI.SaveColWidths(lvItems, Me)
    'Call GUI.SaveColWidths(lvStatementItems, Me)
    Call GUI.SaveColWidths(lvPaymentMade, Me)
    
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

Public Function LoadColumns()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "LoadColumns"
    Const ContainerName = "frmStatements"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    Call GUI.LoadColWidths(lvItems, Me)
    'Call GUI.LoadColWidths(lvStatementItems, Me)
    Call GUI.LoadColWidths(lvPaymentMade, Me)

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
    Const ContainerName = "frmStatements"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    Const ContainerName = "frmStatements"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    If picTreeView.Width > 180 And picTreeView.Height > 800 Then
        cmbDisplay.Move 60, 60, picTreeView.Width - 120
        tvInvoices.Move 60, cmbDisplay.Top + cmbDisplay.Height + 60, picTreeView.Width - 120, picTreeView.Height - 240 - pb2.Height - cmbDisplay.Height
        pb2.Move 60, tvInvoices.Height + tvInvoices.Top + 60, picTreeView.Width - 120
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


Private Sub picTS_Resize(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "picTS_Resize"
    Const ContainerName = "frmStatements"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    If picTS(Index).ScaleWidth > 500 And picTS(Index).ScaleHeight > 500 Then
        Select Case Index
        Case 0
            'lvStatementItems.Move 60, 60, picTS(Index).ScaleWidth - 120, (picTS(Index).ScaleHeight - 150) / 2
            lvItems.Move 60, 60, picTS(Index).ScaleWidth - 120, picTS(Index).ScaleHeight - 120
        Case 1
            lvPaymentMade.Move 60, 60, picTS(Index).ScaleWidth - 120, picTS(Index).ScaleHeight - 120
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

Private Sub ts_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "ts_Click"
    Const ContainerName = "frmStatements"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    
    For X = picTS.LBound To picTS.UBound
        If ts.SelectedItem.Index - 1 <> X Then picTS(X).Visible = False
    Next
    
    If ts.SelectedItem.Index - 1 <= picTS.UBound Then
        picTS(ts.SelectedItem.Index - 1).Move ts.ClientLeft, ts.ClientTop, ts.ClientWidth, ts.ClientHeight
        picTS(ts.SelectedItem.Index - 1).Visible = True
        picTS(ts.SelectedItem.Index - 1).ZOrder 0
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

Private Sub tvInvoices_NodeClick(ByVal Node As MSComctlLib.Node)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "tvInvoices_NodeClick"
    Const ContainerName = "frmStatements"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    Dim rsload As adodb.Recordset
    Dim iRecCount As Variant
    Dim ix As Variant
    Dim itmX As ListItem
    
    Select Case Left(Node.Key, 1)
    Case "i"
    
        lvItems.ListItems.Clear
        
        If bChanged = True Then
            bResult = MySQL.OpenTable(directConn, rsload, , "select * from invoicetraxr Where RecID = " & Mid(nodeP.Key, 2))
            rsload!Comment = txtComment
            rsload.Update
        End If
        
        bResult = MySQL.OpenTable(directConn, rsload, , "select * from invoicetraxr Where RecID = " & Mid(Node.Key, 2))
        
        lblField(0).Caption = Format(rsload!TotalDue - rsload!AmountCredited - rsload!AmountPaid, "Currency")
        lblField(1).Caption = Format(rsload!AmountPaid, "Currency")
        lblField(2).Caption = Format(rsload!PaymentDue, "dd-mm-yyyy Hh:Nn:Ss")
        lblField(3).Caption = IIf(IsNull(Node.Tag), "{NULL}", Node.Tag)
        lblField(4).Caption = Format(rsload!AmountCredited, "Currency")
        
        txtComment = IIf(IsNull(rsload!Comment), "", rsload!Comment)
        txtComment.Locked = False
        bChanged = False
        
        chkFinalised.Value = -Val(rsload!Finalised)
                
        bResult = MySQL.OpenTable(directConn, rsload, , "select Count(*) as RecordCount from invoiceout, accountinfo Where accountinfo.RecID = invoiceout.AccI_RecID AND invoiceout.TraxrID = " & Mid(Node.Key, 2) & " Limit 1")
    
        rsload.MoveLast
    
        If bResult = True Then iRecCount = IIf(IsNull(rsload!RecordCount), 0, rsload!RecordCount)
        
        Dim rsStatement As New adodb.Recordset
        Dim lInvRecID As Long
'        lvStatementItems.ListItems.Clear
        pb1.Value = 0
        If iRecCount <> 0 Then
            pb1.Max = iRecCount
            For ix = 0 To iRecCount Step 30
                bResult = MySQL.OpenTable(directConn, rsload, , "select accountinfo.AccountName, invoiceout.* from invoiceout, accountinfo Where accountinfo.RecID = invoiceout.AccI_RecID AND invoiceout.TraxrID = " & Mid(Node.Key, 2) & " Limit " & ix & ",30")
                Select Case rsload.RecordCount
                Case 0, -1
                Case Else
                    rsload.MoveFirst
                    While Not rsload.EOF And Err.Number = 0
                        gSleep
                        'If lInvRecID <> rsLoad!SubRecID Then
                        '    lInvRecID = rsLoad!SubRecID
                        '    Call MySQL.OpenTable(directConn, rsStatement, , "select * from statementitems where InvRecID = " & lInvRecID)
                        '    If rsStatement.RecordCount > 0 Then
                        '        Set itmX = lvStatementItems.ListItems.Add(, "r" & rsStatement!RecID, rsStatement!Description)
                        '        itmX.SubItems(1) = "" & rsStatement!RecID
                        '        itmX.SubItems(2) = "" & rsStatement!Items
                      '          itmX.SubItems(3) = Format(rsStatement!TotalDue, "Currency")
                        '        itmX.SubItems(4) = Format(rsStatement!Created, "dd-mm-yyyy Hh:Nn:Ss")
                        '    End If
                        'End If
                        Set itmX = lvItems.ListItems.Add(, "r" & rsload!RecID, rsload!Description)
                        
                        'If rsStatement.State <> adStateClosed Then
                        '    If rsStatement.RecordCount > 0 Then
                        '        itmX.SubItems(1) = Format(IIf(IsNull(rsStatement!RecID), 0, rsStatement!RecID), "###,###,###,###,###")
                        '    End If
                        'End If
                        
                        itmX.SubItems(1) = Format(IIf(IsNull(rsload!AmountDue), 0, rsload!AmountDue), "Currency")
                        itmX.SubItems(2) = Format(IIf(IsNull(rsload!GSTCharged), 0, rsload!GSTCharged), "Currency")
                        itmX.SubItems(3) = Format(IIf(IsNull(rsload!TotalDue), 0, rsload!TotalDue), "Currency")
                        itmX.SubItems(4) = Format(IIf(IsNull(rsload!PaymentDue), #9/19/1950#, rsload!PaymentDue), "dd-mm-yyyy Hh:Nn:Ss")
                        itmX.SubItems(5) = Format(IIf(IsNull(rsload!AmountPaid), 0, rsload!AmountPaid), "Currency")
                        itmX.SubItems(6) = Format(IIf(IsNull(rsload!AmountRefunded + rsload!GSTRefunded), 0, rsload!AmountRefunded + rsload!GSTRefunded), "Currency")
                        If pb1.Max < pb1.Value + 1 Then pb1.Value = pb1.Value + 0 Else pb1.Value = pb1.Value + 1
                        rsload.MoveNext
                    Wend
                End Select
            Next ix
        End If
        
        LoadPayments Mid(Node.Key, 2)
        
        Set nodeP = Node
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
Public Function LoadPayments(Optional ltmpRecID As Variant)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "LoadPayments"
    Const ContainerName = "frmStatements"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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
    Dim itmX As ListItem
    Dim bResult As Boolean
    
    bResult = MySQL.OpenTable(directConn, rsload, , "select * from invtrx_payment where InvTrxRecID = " & ltmpRecID)
    lvPaymentMade.ListItems.Clear
    If rsload.RecordCount > 0 Then
        rsload.MoveFirst
        While Not rsload.EOF And Err.Number = 0
            Set itmX = lvPaymentMade.ListItems.Add(, "r" & rsload!RecID, Format(rsload!WhenPaid, "dd-mm-yyyy Hh:Nn:Ss"))
            itmX.SubItems(1) = IIf(rsload!Amount = 0, "", Format(rsload!Amount, "Currency"))
            itmX.SubItems(2) = IIf(rsload!GST = 0, "", Format(rsload!GST, "Currency"))
            itmX.SubItems(3) = IIf(rsload!sub = 0, "", Format(rsload!sub, "Currency"))
            rsload.MoveNext
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

Private Sub txtComment_Change()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtComment_Change"
    Const ContainerName = "frmStatements"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
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


    bChanged = True
    
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
