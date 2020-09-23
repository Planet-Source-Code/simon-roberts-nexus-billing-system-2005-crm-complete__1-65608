VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCreateCat 
   BackColor       =   &H00A9C9AE&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Create Category"
   ClientHeight    =   8145
   ClientLeft      =   4620
   ClientTop       =   5085
   ClientWidth     =   6585
   Icon            =   "frmCreateCat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00A9C9AE&
      Caption         =   "Do not create category"
      Height          =   405
      Index           =   1
      Left            =   3870
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7530
      Width           =   2565
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00A9C9AE&
      Caption         =   "Create Category"
      Height          =   405
      Index           =   0
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7530
      Width           =   2565
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00A9C9AE&
      Caption         =   "Create Category"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.Frame Frame5 
         BackColor       =   &H00A9C9AE&
         Caption         =   "Sub Node of Category"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   150
         TabIndex        =   11
         Top             =   4200
         Width           =   5985
         Begin VB.ListBox cmbNodes 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   2460
            Left            =   150
            TabIndex        =   12
            Top             =   300
            Width           =   5715
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00A9C9AE&
         Caption         =   "Minimum Security Level to Access Category"
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
         TabIndex        =   7
         Top             =   3390
         Width           =   5985
         Begin MSComctlLib.Slider sldSec 
            Height          =   315
            Left            =   150
            TabIndex        =   8
            Top             =   300
            Width           =   5745
            _ExtentX        =   10134
            _ExtentY        =   556
            _Version        =   393216
            Max             =   100
            TickFrequency   =   3
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00A9C9AE&
         Caption         =   "Icon"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1965
         Left            =   150
         TabIndex        =   3
         Top             =   1410
         Width           =   5985
         Begin VB.PictureBox picIcon 
            BackColor       =   &H00A9C9AE&
            BorderStyle     =   0  'None
            Height          =   1605
            Left            =   120
            ScaleHeight     =   1605
            ScaleWidth      =   5505
            TabIndex        =   6
            Top             =   240
            Width           =   5505
            Begin VB.Image imgIcon 
               BorderStyle     =   1  'Fixed Single
               Height          =   540
               Index           =   0
               Left            =   60
               Picture         =   "frmCreateCat.frx":030A
               Top             =   60
               Width           =   540
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00A9C9AE&
            BorderStyle     =   0  'None
            Height          =   1695
            Left            =   90
            ScaleHeight     =   1695
            ScaleWidth      =   5805
            TabIndex        =   4
            Top             =   210
            Width           =   5805
            Begin VB.VScrollBar vslIcon 
               Height          =   1665
               Left            =   5550
               Max             =   100
               TabIndex        =   5
               Top             =   0
               Width           =   255
            End
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00A9C9AE&
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   150
         TabIndex        =   1
         Top             =   330
         Width           =   5955
         Begin VB.TextBox txtField 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   510
            Left            =   120
            TabIndex        =   2
            Top             =   270
            Width           =   5745
         End
      End
   End
   Begin MSComctlLib.ImageList ilSML 
      Left            =   5940
      Top             =   120
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
            Picture         =   "frmCreateCat.frx":074C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":0B9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":0FF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1442
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1894
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":2138
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":258A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":29DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":2E2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":3280
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":36D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":3B24
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":3E3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":4290
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":46E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":4B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":4F86
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":53D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":582A
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":5C7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":60CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":6520
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":6972
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":6DC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":7216
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":7668
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":7ABA
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":7F0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":835E
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":87B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":8C02
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":9054
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":94A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":98F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":9D4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":A19C
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":A5EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":AA40
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":AE92
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":B2E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":B736
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":BB88
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":BFDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":C42C
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":C87E
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":CCD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":D122
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":D574
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":D9C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":DE18
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":E26A
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":E6BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":EB0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":EF60
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":F3B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":F804
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":FC56
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":100A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":104FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1094C
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":10D9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":111F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":11642
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":11A94
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":11EE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":12338
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1278A
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":12BDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1302E
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":13480
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":138D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":13D24
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":14476
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":148C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":14D1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1516C
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":155BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":15A10
            Key             =   ""
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":15D2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":16044
            Key             =   ""
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1635E
            Key             =   ""
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":16678
            Key             =   ""
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":16992
            Key             =   ""
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":16CAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":16FC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":172E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":175FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":17914
            Key             =   ""
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":17C2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":17F48
            Key             =   ""
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":18262
            Key             =   ""
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1857C
            Key             =   ""
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":18896
            Key             =   ""
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":18BB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":18ECA
            Key             =   ""
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":191E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":194FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":19818
            Key             =   ""
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":19B32
            Key             =   ""
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":19E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1A166
            Key             =   ""
         EndProperty
         BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1A480
            Key             =   ""
         EndProperty
         BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1A79A
            Key             =   ""
         EndProperty
         BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1ABEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1B03E
            Key             =   ""
         EndProperty
         BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1B490
            Key             =   ""
         EndProperty
         BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1B8E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1BD34
            Key             =   ""
         EndProperty
         BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1C186
            Key             =   ""
         EndProperty
         BeginProperty ListImage111 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1C5D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage112 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1CA2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage113 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1CE7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage114 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1D2CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage115 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1D720
            Key             =   ""
         EndProperty
         BeginProperty ListImage116 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1DB72
            Key             =   ""
         EndProperty
         BeginProperty ListImage117 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1DFC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage118 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1E416
            Key             =   ""
         EndProperty
         BeginProperty ListImage119 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1E868
            Key             =   ""
         EndProperty
         BeginProperty ListImage120 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1ECBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage121 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1F10C
            Key             =   ""
         EndProperty
         BeginProperty ListImage122 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1F55E
            Key             =   ""
         EndProperty
         BeginProperty ListImage123 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1F9B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage124 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":1FE02
            Key             =   ""
         EndProperty
         BeginProperty ListImage125 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":20254
            Key             =   ""
         EndProperty
         BeginProperty ListImage126 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":206A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage127 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":20AF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage128 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":20F4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage129 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":2139C
            Key             =   ""
         EndProperty
         BeginProperty ListImage130 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":217EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage131 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":21C40
            Key             =   ""
         EndProperty
         BeginProperty ListImage132 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":22092
            Key             =   ""
         EndProperty
         BeginProperty ListImage133 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":224E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage134 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":22936
            Key             =   ""
         EndProperty
         BeginProperty ListImage135 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":22D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage136 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":231DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage137 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":2362C
            Key             =   ""
         EndProperty
         BeginProperty ListImage138 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":23A7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage139 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":23ED0
            Key             =   ""
         EndProperty
         BeginProperty ListImage140 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":24322
            Key             =   ""
         EndProperty
         BeginProperty ListImage141 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":24774
            Key             =   ""
         EndProperty
         BeginProperty ListImage142 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":24BC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage143 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":25018
            Key             =   ""
         EndProperty
         BeginProperty ListImage144 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":2546A
            Key             =   ""
         EndProperty
         BeginProperty ListImage145 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":258BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage146 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":25D0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage147 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateCat.frx":26160
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCreateCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public tv As TreeView
Public SubNode As Double
Public Description As String
Public SecLevel As Byte
Public iIcon As Integer
Public RecID As Double


Private Sub Command1_Click(Index As Integer)

    Select Case Index
    Case 0
    
        If Trim(txtField.Text) <> "" Then
        
            Dim ix As Integer
            
            For ix = imgIcon.LBound To imgIcon.UBound
                If imgIcon(ix).BorderStyle = 1 Then
                    Me.iIcon = ix
                    Exit For
                End If
            Next ix
            
            If cmbNodes.ListIndex <> -1 Then
                SubNode = cmbNodes.ItemData(cmbNodes.ListIndex)
            Else
                SubNode = 0
            End If
            
            Description = txtField.Text
            SecLevel = sldSec.Value
            
            If Me.RecID = 0 Then
                On Error Resume Next
                Do
                       Err.Clear
                       
                    Me.RecID = MySQL.GetTMPRecID("exp_categories", ADOConn)
                    MySQL.Execute ADOConn, "insert into exp_categories (RecID,VirtualID,SysopID) Values('" & Me.RecID & "','" & Login.lVirtualID & "','" & Login.lSysopID & "')"
                    
                Loop Until Err.Number = 0
                
                GUI.mapCategory.Add "r" & Me.RecID, Me.RecID, SubNode, Val(Login.lVirtualID), Val(Login.lSysopID), iIcon, Description, "exp001", sldSec.Value, "r" & Me.RecID
                
                Select Case SubNode
                Case 0
                    tv.NodeS.Add , , "r" & Me.RecID, Description, iIcon, iIcon
                    
                Case Else
                
                    tv.NodeS.Add "r" & SubNode, tvwChild, "r" & Me.RecID, Description, iIcon, iIcon
                    
                End Select
            End If
            
            MySQL.Execute ADOConn, "update exp_categories set SubRecID = '" & SubNode & "', Description = '" & MySQL.ESC(Description) & "', SecLevel = '" & SecLevel & "' where RecID = " & Me.RecID
                
            GUI.mapCategory(GUI.mapCategory.FindKey("r" & Me.RecID)).Description = Description
            GUI.mapCategory(GUI.mapCategory.FindKey("r" & Me.RecID)).SecLevel = sldSec.Value
            GUI.mapCategory(GUI.mapCategory.FindKey("r" & Me.RecID)).SubRecID = SubNode
            GUI.mapCategory(GUI.mapCategory.FindKey("r" & Me.RecID)).Icon = iIcon
            
            Unload Me
            
        End If
        
        
    Case 1
    
        Unload Me
        
    End Select
End Sub

Private Sub Form_Load()

    Dim I As Integer
    Static Level As Integer
    
    For I = 1 To ilSML.ListImages.Count
        If I = 1 Then
            imgIcon(I - 1).Picture = ilSML.ListImages(I).ExtractIcon
            imgIcon(I - 1).Move 120, 0
        Else
            Load imgIcon(I - 1)
            Set imgIcon(I - 1).Container = picIcon
            imgIcon(I - 1).Picture = ilSML.ListImages(I).ExtractIcon
            imgIcon(I - 1).Move imgIcon(0).Left + imgIcon(I - 2).Left + imgIcon(0).Width, Level * (imgIcon(0).Width + 120)
            If imgIcon(I - 1).Left > picIcon.ScaleWidth - imgIcon(I - 1).Width Then
                Level = Level + 1
                imgIcon(I - 1).Move imgIcon(0).Left, Level * (imgIcon(0).Width + 120)
            End If
            imgIcon(I - 1).Visible = True
            imgIcon(I - 1).BorderStyle = 0
        End If
    
    Next I
    
    picIcon.Height = imgIcon(ilSML.ListImages.Count - 1).Top + imgIcon(ilSML.ListImages.Count - 1).Height
    Set picIcon.Container = Picture1
    picIcon.Move 0, 0
    Picture1.Move 90, 240, Frame3.Width - 180, Frame3.Height - 330
    
    cmbNodes.AddItem "[Primary Node]"
    cmbNodes.ItemData(cmbNodes.ListCount - 1) = 0
    
    
    If tv.NodeS.Count > 0 Then
        Dim lx As Long
       
        For lx = 1 To tv.NodeS.Count
            cmbNodes.AddItem tv.NodeS(lx).Text
            cmbNodes.ItemData(cmbNodes.ListCount - 1) = Val(Mid(tv.NodeS(lx).Key, 2))
            If tv.SelectedItem Is Nothing Then
            
            Else
            
                If tv.NodeS(lx).Index = tv.SelectedItem.Index Then
                    cmbNodes.ListIndex = cmbNodes.ListCount - 1
                End If
                
            End If
        Next lx
        
    End If
    
End Sub

Private Sub imgIcon_Click(Index As Integer)

    Dim ix As Integer
    
    For ix = imgIcon.LBound To imgIcon.UBound
        imgIcon(ix).BorderStyle = 0
    Next
    
    imgIcon(Index).BorderStyle = 1
    
End Sub

Private Sub Slider1_Change()

    If Slider1.Value > Login.lLevel Then Slider1.Value = Login.lLevel
    
End Sub

Private Sub vslIcon_Change()

    picIcon.Top = -((vslIcon.Value / 100) * (picIcon.Height - Picture1.Height))
    
    
End Sub
