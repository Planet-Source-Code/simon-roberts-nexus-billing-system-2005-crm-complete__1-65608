VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEXP 
   BackColor       =   &H00A9C9AE&
   Caption         =   "Expences managment"
   ClientHeight    =   10875
   ClientLeft      =   1635
   ClientTop       =   2505
   ClientWidth     =   16980
   Icon            =   "frmEXP_new.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10875
   ScaleWidth      =   16980
   Begin VB.Frame Frame2 
      BackColor       =   &H00A9C9AE&
      Height          =   10605
      Left            =   5280
      TabIndex        =   2
      Top             =   90
      Width           =   11595
      Begin VB.Frame Frame3 
         BackColor       =   &H0094BD7B&
         Caption         =   "Days and Date to Display"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2355
         Left            =   150
         TabIndex        =   4
         Top             =   240
         Width           =   11355
         Begin VB.CommandButton cmdQuery 
            BackColor       =   &H0094BD7B&
            Caption         =   "Requery"
            Enabled         =   0   'False
            Height          =   345
            Left            =   9690
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   1500
            Width           =   1515
         End
         Begin VB.CommandButton cmdExport 
            BackColor       =   &H0094BD7B&
            Caption         =   "Export to CSV"
            Height          =   345
            Left            =   9690
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   1890
            Width           =   1515
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H0094BD7B&
            Caption         =   "Totals"
            Height          =   1035
            Left            =   150
            TabIndex        =   10
            Top             =   1200
            Width           =   5385
            Begin VB.Label lblTotals 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "$ 0.00 Unpaid Expenditure"
               Height          =   225
               Index           =   5
               Left            =   2790
               TabIndex        =   16
               Top             =   690
               Width           =   2445
            End
            Begin VB.Label lblTotals 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "0 Total Unpaid Transactions"
               Height          =   225
               Index           =   4
               Left            =   180
               TabIndex        =   15
               Top             =   690
               Width           =   2445
            End
            Begin VB.Label lblTotals 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "$ 0.00 Paided Expenditure"
               Height          =   225
               Index           =   3
               Left            =   2790
               TabIndex        =   14
               Top             =   450
               Width           =   2445
            End
            Begin VB.Label lblTotals 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "0 Total Paid Transactions"
               Height          =   225
               Index           =   2
               Left            =   180
               TabIndex        =   13
               Top             =   450
               Width           =   2445
            End
            Begin VB.Label lblTotals 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "$ 0.00 Total Expenditure"
               Height          =   225
               Index           =   1
               Left            =   2790
               TabIndex        =   12
               Top             =   210
               Width           =   2445
            End
            Begin VB.Label lblTotals 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "0 Total Transactions"
               Height          =   225
               Index           =   0
               Left            =   180
               TabIndex        =   11
               Top             =   210
               Width           =   2445
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H0094BD7B&
            Caption         =   "Date query is reporting from"
            Height          =   885
            Left            =   150
            TabIndex        =   5
            Top             =   270
            Width           =   5385
            Begin VB.Frame Frame6 
               BackColor       =   &H0094BD7B&
               Caption         =   "To"
               Height          =   585
               Index           =   1
               Left            =   2730
               TabIndex        =   7
               Top             =   210
               Width           =   2565
               Begin VB.TextBox txtDate 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H0094BD7B&
                  BorderStyle     =   0  'None
                  Height          =   285
                  Index           =   1
                  Left            =   120
                  TabIndex        =   8
                  Top             =   210
                  Width           =   2325
               End
            End
            Begin VB.Frame Frame6 
               BackColor       =   &H0094BD7B&
               Caption         =   "From"
               Height          =   585
               Index           =   0
               Left            =   90
               TabIndex        =   6
               Top             =   210
               Width           =   2565
               Begin VB.TextBox txtDate 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H0094BD7B&
                  BorderStyle     =   0  'None
                  Height          =   285
                  Index           =   0
                  Left            =   120
                  TabIndex        =   9
                  Top             =   210
                  Width           =   2325
               End
            End
         End
      End
      Begin MSComctlLib.ListView lvIn 
         Height          =   7785
         Left            =   150
         TabIndex        =   3
         Top             =   2700
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   13732
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
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
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00A9C9AE&
      Caption         =   "In Book Category Tree"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10635
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   4995
      Begin VB.CommandButton cmdMakeItem 
         Caption         =   "Create &Item Entry"
         Default         =   -1  'True
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2760
         TabIndex        =   18
         Top             =   10080
         Width           =   2085
      End
      Begin VB.CommandButton cmdMakeCategory 
         Caption         =   "Create &Category"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   150
         TabIndex        =   17
         Top             =   10080
         Width           =   2085
      End
      Begin MSComctlLib.TreeView tvCat 
         Height          =   9675
         Left            =   150
         TabIndex        =   1
         Top             =   330
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   17066
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
         Top             =   0
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
               Picture         =   "frmEXP_new.frx":0442
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":0894
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":0CE6
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1138
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":158A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":19DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1E2E
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":2280
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":26D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":2B24
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":2F76
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":33C8
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":381A
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":3B34
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":3F86
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":43D8
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":482A
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":4C7C
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":50CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":5520
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":5972
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":5DC4
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":6216
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":6668
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":6ABA
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":6F0C
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":735E
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":77B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":7C02
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":8054
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":84A6
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":88F8
               Key             =   ""
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":8D4A
               Key             =   ""
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":919C
               Key             =   ""
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":95EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":9A40
               Key             =   ""
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":9E92
               Key             =   ""
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":A2E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":A736
               Key             =   ""
            EndProperty
            BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":AB88
               Key             =   ""
            EndProperty
            BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":AFDA
               Key             =   ""
            EndProperty
            BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":B42C
               Key             =   ""
            EndProperty
            BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":B87E
               Key             =   ""
            EndProperty
            BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":BCD0
               Key             =   ""
            EndProperty
            BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":C122
               Key             =   ""
            EndProperty
            BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":C574
               Key             =   ""
            EndProperty
            BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":C9C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":CE18
               Key             =   ""
            EndProperty
            BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":D26A
               Key             =   ""
            EndProperty
            BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":D6BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":DB0E
               Key             =   ""
            EndProperty
            BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":DF60
               Key             =   ""
            EndProperty
            BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":E3B2
               Key             =   ""
            EndProperty
            BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":E804
               Key             =   ""
            EndProperty
            BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":EC56
               Key             =   ""
            EndProperty
            BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":F0A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":F4FA
               Key             =   ""
            EndProperty
            BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":F94C
               Key             =   ""
            EndProperty
            BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":FD9E
               Key             =   ""
            EndProperty
            BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":101F0
               Key             =   ""
            EndProperty
            BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":10642
               Key             =   ""
            EndProperty
            BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":10A94
               Key             =   ""
            EndProperty
            BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":10EE6
               Key             =   ""
            EndProperty
            BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":11338
               Key             =   ""
            EndProperty
            BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1178A
               Key             =   ""
            EndProperty
            BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":11BDC
               Key             =   ""
            EndProperty
            BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1202E
               Key             =   ""
            EndProperty
            BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":12480
               Key             =   ""
            EndProperty
            BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":128D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":12D24
               Key             =   ""
            EndProperty
            BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":13176
               Key             =   ""
            EndProperty
            BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":135C8
               Key             =   ""
            EndProperty
            BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":13A1A
               Key             =   ""
            EndProperty
            BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1416C
               Key             =   ""
            EndProperty
            BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":145BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":14A10
               Key             =   ""
            EndProperty
            BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":14E62
               Key             =   ""
            EndProperty
            BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":152B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":15706
               Key             =   ""
            EndProperty
            BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":15A20
               Key             =   ""
            EndProperty
            BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":15D3A
               Key             =   ""
            EndProperty
            BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":16054
               Key             =   ""
            EndProperty
            BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1636E
               Key             =   ""
            EndProperty
            BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":16688
               Key             =   ""
            EndProperty
            BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":169A2
               Key             =   ""
            EndProperty
            BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":16CBC
               Key             =   ""
            EndProperty
            BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":16FD6
               Key             =   ""
            EndProperty
            BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":172F0
               Key             =   ""
            EndProperty
            BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1760A
               Key             =   ""
            EndProperty
            BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":17924
               Key             =   ""
            EndProperty
            BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":17C3E
               Key             =   ""
            EndProperty
            BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":17F58
               Key             =   ""
            EndProperty
            BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":18272
               Key             =   ""
            EndProperty
            BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1858C
               Key             =   ""
            EndProperty
            BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":188A6
               Key             =   ""
            EndProperty
            BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":18BC0
               Key             =   ""
            EndProperty
            BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":18EDA
               Key             =   ""
            EndProperty
            BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":191F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1950E
               Key             =   ""
            EndProperty
            BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":19828
               Key             =   ""
            EndProperty
            BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":19B42
               Key             =   ""
            EndProperty
            BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":19E5C
               Key             =   ""
            EndProperty
            BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1A176
               Key             =   ""
            EndProperty
            BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1A490
               Key             =   ""
            EndProperty
            BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1A8E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1AD34
               Key             =   ""
            EndProperty
            BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1B186
               Key             =   ""
            EndProperty
            BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1B5D8
               Key             =   ""
            EndProperty
            BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1BA2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1BE7C
               Key             =   ""
            EndProperty
            BeginProperty ListImage111 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1C2CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage112 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1C720
               Key             =   ""
            EndProperty
            BeginProperty ListImage113 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1CB72
               Key             =   ""
            EndProperty
            BeginProperty ListImage114 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1CFC4
               Key             =   ""
            EndProperty
            BeginProperty ListImage115 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1D416
               Key             =   ""
            EndProperty
            BeginProperty ListImage116 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1D868
               Key             =   ""
            EndProperty
            BeginProperty ListImage117 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1DCBA
               Key             =   ""
            EndProperty
            BeginProperty ListImage118 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1E10C
               Key             =   ""
            EndProperty
            BeginProperty ListImage119 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1E55E
               Key             =   ""
            EndProperty
            BeginProperty ListImage120 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1E9B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage121 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1EE02
               Key             =   ""
            EndProperty
            BeginProperty ListImage122 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1F254
               Key             =   ""
            EndProperty
            BeginProperty ListImage123 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1F6A6
               Key             =   ""
            EndProperty
            BeginProperty ListImage124 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1FAF8
               Key             =   ""
            EndProperty
            BeginProperty ListImage125 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":1FF4A
               Key             =   ""
            EndProperty
            BeginProperty ListImage126 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":2039C
               Key             =   ""
            EndProperty
            BeginProperty ListImage127 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":207EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage128 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":20C40
               Key             =   ""
            EndProperty
            BeginProperty ListImage129 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":21092
               Key             =   ""
            EndProperty
            BeginProperty ListImage130 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":214E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage131 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":21936
               Key             =   ""
            EndProperty
            BeginProperty ListImage132 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":21D88
               Key             =   ""
            EndProperty
            BeginProperty ListImage133 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":221DA
               Key             =   ""
            EndProperty
            BeginProperty ListImage134 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":2262C
               Key             =   ""
            EndProperty
            BeginProperty ListImage135 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":22A7E
               Key             =   ""
            EndProperty
            BeginProperty ListImage136 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":22ED0
               Key             =   ""
            EndProperty
            BeginProperty ListImage137 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":23322
               Key             =   ""
            EndProperty
            BeginProperty ListImage138 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":23774
               Key             =   ""
            EndProperty
            BeginProperty ListImage139 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":23BC6
               Key             =   ""
            EndProperty
            BeginProperty ListImage140 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":24018
               Key             =   ""
            EndProperty
            BeginProperty ListImage141 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":2446A
               Key             =   ""
            EndProperty
            BeginProperty ListImage142 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":248BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage143 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":24D0E
               Key             =   ""
            EndProperty
            BeginProperty ListImage144 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":25160
               Key             =   ""
            EndProperty
            BeginProperty ListImage145 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":255B2
               Key             =   ""
            EndProperty
            BeginProperty ListImage146 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":25A04
               Key             =   ""
            EndProperty
            BeginProperty ListImage147 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEXP_new.frx":25E56
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmEXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdMakeCategory_Click()

    Dim fCat As New frmCreateCat
    
    Set fCat.tv = tvCat
    fCat.Show 1
    
    If fCat.RecID <> 0 Then
        Me.LoadCategories tvCat
    End If
    
End Sub

Private Sub cmdMakeItem_Click()

    Dim fExp As New frmEXPa
    
    fExp.CategoryID = Val(Mid(tvCat.SelectedItem.Key, 2))
    fExp.Show 1
    
    If fExp.RecID <> 0 Then
        
        cmdQuery_Click
    
    End If
    
End Sub

Private Sub cmdQuery_Click()

    On Error Resume Next
    Dim rsload As adodb.Recordset
    
    'Debug.Print "Copy: " & MySQL.virtualisp("select count(distinct RecID) as RecCount, sum(AmountDue) as sumDue, sum(AmountPaid) as sumPaid from exp_inbook where CategoryID = '" & Mid(tvCat.selectedItem.Key, 2) & "' and (Created >= '" & Format(CDate(txtDate(0).Text), "yyyymmddhhnnss") & "' and Created <= '" & Format(CDate(txtDate(1).Text), "yyyymmddhhnnss") & "')", "exp_inbook")
    
    Call MySQL.OpenTable(directConn, rsload, , "select count(distinct exp_inbook.RecID) as RecCount, sum(exp_inbook.AmountDue) as sumDue, sum(exp_inbook.AmountPaid) as sumPaid from exp_inbook where CategoryID = '" & Mid(tvCat.SelectedItem.Key, 2) & "' and (Created >= '" & Format(CDate(txtDate(0).Text), "yyyymmddhhnnss") & "' and Created <= '" & Format(CDate(txtDate(1).Text), "yyyymmddhhnnss") & "') and exp_inbook.VirtualID = '" & Login.lVirtualID & "'")
    
    If rsload.State = adStateOpen Then lblTotals(0).Caption = rsload!RecCount & " Total Transactions"
    If rsload.State = adStateOpen Then lblTotals(3).Caption = Format(IIf(IsNull(rsload!sumPaid), 0, rsload!sumPaid), "Currency") & " Paid Expenditure"
    If rsload.State = adStateOpen Then lblTotals(1).Caption = Format(IIf(IsNull(rsload!sumDue), 0, rsload!sumDue), "Currency") & " Total Expenditure"
    If rsload.State = adStateOpen Then lblTotals(5).Caption = Format(IIf(IsNull(rsload!sumDue), 0, rsload!sumDue) - IIf(IsNull(rsload!sumPaid), 0, rsload!sumPaid), "Currency") & " Unpaid Expenditure"
    
    Call MySQL.OpenTable(directConn, rsload, , MySQL.virtualisp("select count(distinct exp_inbook.RecID) as RecCount from exp_inbook where exp_inbook.AmountPaid >= exp_inbook.AmountDue and exp_inbook.CategoryID = '" & Mid(tvCat.SelectedItem.Key, 2) & "' and (exp_inbook.Created >= '" & Format(CDate(txtDate(0).Text), "yyyymmddhhnnss") & "' and exp_inbook.Created <= '" & Format(CDate(txtDate(1).Text), "yyyymmddhhnnss") & "') Group by exp_inbook.RecID", "exp_inbook"))
    If rsload.State = adStateOpen Then lblTotals(2).Caption = rsload!RecCount & " Total Paid Expenditures"
    
    Call MySQL.OpenTable(directConn, rsload, , MySQL.virtualisp("select count(distinct exp_inbook.RecID) as RecCount from exp_inbook where AmountPaid < AmountDue and CategoryID = '" & Mid(tvCat.SelectedItem.Key, 2) & "' and (Created >= '" & Format(CDate(txtDate(0).Text), "yyyymmddhhnnss") & "' and Created <= '" & Format(CDate(txtDate(1).Text), "yyyymmddhhnnss") & "')", "exp_inbook"))
    If rsload.State = adStateOpen Then lblTotals(4).Caption = rsload!RecCount & " Total Unpaid Transactions"
    
    Call MySQL.OpenTable(directConn, rsload, , MySQL.virtualisp("select distinct exp_inbook.* from exp_inbook where CategoryID = '" & Mid(tvCat.SelectedItem.Key, 2) & "' and (Created >= '" & Format(CDate(txtDate(0).Text), "yyyymmddhhnnss") & "' and Created <= '" & Format(CDate(txtDate(1).Text), "yyyymmddhhnnss") & "')", "exp_inbook"))
    
    If rsload.State = adStateOpen Then
        
        Call MySQL.fillLV(directConn, rsload, lvIn)
        
    End If

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()

    Me.LoadCategories tvCat
        
    GUI.LoadColWidths lvIn, Me
    
    txtDate(0).Text = GetSetting(App.ProductName, "main", "expfrom", Format(Now, "yyyy-mm-dd ttttt"))
    txtDate(1).Text = GetSetting(App.ProductName, "main", "expTo", Format(Now, "yyyy-mm-dd ttttt"))
    
End Sub


Sub LoadCategories(tv As TreeView)

    tv.NodeS.Clear
    
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

Private Sub optRep_Click(Index As Integer)

    Select Case Index
    Case 0 'today
        txtDate(1).Text = Format(Now, "yyyy-mm-dd 12:01:00am")
        txtDate(0).Text = Format(Now, "yyyy-mm-dd 24:00:00pm")
    Case 1 'yesturday
        txtDate(1).Text = Format(DateAdd(-1, "d", sysnow), "yyyy-mm-dd 12:01:00am")
        txtDate(0).Text = Format(DateAdd(-1, "d", sysnow), "yyyy-mm-dd 24:00:00pm")
    Case 2 'this week
        Select Case Format(Now, "dddd")
        Case "Sunday"
            txtDate(0).Text = Format(DateAdd(-6, "d", sysnow), "yyyy-mm-dd 12:01:00am")
            txtDate(1).Text = Format(Now, "yyyy-mm-dd 24:00:00pm")
        Case "Saturday"
            txtDate(0).Text = Format(DateAdd(-5, "d", sysnow), "yyyy-mm-dd 12:01:00am")
            txtDate(1).Text = Format(Now, "yyyy-mm-dd 24:00:00pm")
        Case "Friday"
            txtDate(0).Text = Format(DateAdd(-4, "d", sysnow), "yyyy-mm-dd 12:01:00am")
            txtDate(1).Text = Format(Now, "yyyy-mm-dd 24:00:00pm")
        Case "Thursday"
            txtDate(0).Text = Format(DateAdd(-3, "d", sysnow), "yyyy-mm-dd 12:01:00am")
            txtDate(1).Text = Format(Now, "yyyy-mm-dd 24:00:00pm")
        Case "Wednesday"
            txtDate(0).Text = Format(DateAdd(-2, "d", sysnow), "yyyy-mm-dd 12:01:00am")
            txtDate(1).Text = Format(Now, "yyyy-mm-dd 24:00:00pm")
        Case "Tuesday"
            txtDate(0).Text = Format(DateAdd(-1, "d", sysnow), "yyyy-mm-dd 12:01:00am")
            txtDate(1).Text = Format(Now, "yyyy-mm-dd 24:00:00pm")
        Case "Monday"
            txtDate(0).Text = Format(Now, "yyyy-mm-dd 12:01:00am")
            txtDate(1).Text = Format(Now, "yyyy-mm-dd 24:00:00pm")
        End Select
    Case 3 'last week
        Select Case Format(Now, "dddd")
        Case "Sunday"
            txtDate(0).Text = Format(DateAdd(-14, "d", sysnow), "yyyy-mm-dd 12:01:00am")
            txtDate(1).Text = Format(DateAdd(-7, "d", sysnow), "yyyy-mm-dd 24:00:00pm")
        Case "Saturday"
            txtDate(0).Text = Format(DateAdd(-13, "d", sysnow), "yyyy-mm-dd 12:01:00am")
            txtDate(1).Text = Format(DateAdd(-6, "d", sysnow), "yyyy-mm-dd 24:00:00pm")
        Case "Friday"
            txtDate(0).Text = Format(DateAdd(-12, "d", sysnow), "yyyy-mm-dd 12:01:00am")
            txtDate(1).Text = Format(DateAdd(-5, "d", sysnow), "yyyy-mm-dd 24:00:00pm")
        Case "Thursday"
            txtDate(0).Text = Format(DateAdd(-11, "d", sysnow), "yyyy-mm-dd 12:01:00am")
            txtDate(1).Text = Format(Now, "yyyy-mm-dd 24:00:00pm")
        Case "Wednesday"
            txtDate(0).Text = Format(DateAdd(-2, "d", sysnow), "yyyy-mm-dd 12:01:00am")
            txtDate(1).Text = Format(Now, "yyyy-mm-dd 24:00:00pm")
        Case "Tuesday"
            txtDate(0).Text = Format(DateAdd(-1, "d", sysnow), "yyyy-mm-dd 12:01:00am")
            txtDate(1).Text = Format(Now, "yyyy-mm-dd 24:00:00pm")
        Case "Monday"
            txtDate(0).Text = Format(Now, "yyyy-mm-dd 12:01:00am")
            txtDate(1).Text = Format(Now, "yyyy-mm-dd 24:00:00pm")
        End Select
    Case 4 'this month so far
    Case 5 'Previous month
    Case 6 'Previous quarter
    Case 7 'specified
    
    
    End Select
    
    
End Sub

Private Sub Form_Resize()

    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.ScaleHeight < 3400 Then Exit Sub
    
    Frame1.Move 60, 60, Me.ScaleWidth * 0.315, Me.ScaleHeight - 120
    Frame2.Move Frame1.Left + Frame1.Width + 120, 60, Me.ScaleWidth - (Frame1.Left + Frame1.Width + 180), Me.ScaleHeight - 120
    
    tvCat.Move tvCat.Left, tvCat.Top, Frame1.Width - (tvCat.Left * 2), Frame1.Height - cmdMakeCategory.Height - 480
    cmdMakeCategory.Top = tvCat.Top + tvCat.Height + 60
    cmdMakeItem.Move tvCat.Left + tvCat.Width - cmdMakeItem.Width, cmdMakeCategory.Top
    
    Frame3.Move Frame3.Left, Frame3.Top, Frame2.Width - Frame3.Left * 2
    cmdExport.Move Frame3.Width - cmdExport.Width - 120, Frame3.Height - cmdExport.Height - 120
    cmdQuery.Move cmdExport.Left, cmdExport.Top - cmdQuery.Height - 60
    
    lvIn.Move lvIn.Left, lvIn.Top, Frame3.Width, Frame2.Height - lvIn.Top - 240
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    GUI.SaveColWidths lvIn, Me

    SaveSetting App.ProductName, "main", "expfrom", txtDate(0).Text
    SaveSetting App.ProductName, "main", "expTo", txtDate(1).Text
    
End Sub

Private Sub lvIn_DblClick()

    If lvIn.SelectedItem Is Nothing Then
    
    Else
    
        Dim fExp As New frmEXPa
        
        fExp.RecID = Val(Mid(lvIn.SelectedItem.Key, 2))
        fExp.CategoryID = Val(Mid(tvCat.SelectedItem.Key, 2))
        fExp.Show 1
    
    End If
End Sub

Private Sub tvCat_NodeClick(ByVal Node As MSComctlLib.Node)

    cmdMakeItem.Enabled = True
    cmdQuery.Enabled = True
    
    If lvIn.Tag <> GUI.mapCategory(GUI.mapCategory.FindKey(Node.Key)).formcode Then
    
        Call MySQL.SetColumnHeaders(GUI.mapCategory(GUI.mapCategory.FindKey(Node.Key)).formcode, lvIn, "", directConn)
        lvIn.Tag = GUI.mapCategory(GUI.mapCategory.FindKey(Node.Key)).formcode
        
    End If
    
    Frame2.Caption = "" & Node.Text & " is the category you are currently working with"
    Call cmdQuery_Click
    
End Sub

Private Sub txtDate_DblClick(Index As Integer)

    Dim fDate As New frmDateTime
    
    fDate.dDate = CDate(txtDate(Index))
    
    fDate.Show 1
    
    txtDate(Index) = Format(fDate.dDate, "yyyy-mm-dd ttttt")
    
End Sub
