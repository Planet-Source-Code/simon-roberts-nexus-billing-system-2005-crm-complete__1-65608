VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAgSites 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Your Agencies Sites"
   ClientHeight    =   9330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10920
   Icon            =   "frmAgSites.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9330
   ScaleWidth      =   10920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTS 
      BorderStyle     =   0  'None
      Height          =   8595
      Index           =   1
      Left            =   4860
      ScaleHeight     =   8595
      ScaleWidth      =   5805
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   5805
      Begin VB.Frame Frame3 
         Caption         =   "Sysop Manifest"
         Height          =   8385
         Index           =   1
         Left            =   30
         TabIndex        =   11
         Top             =   210
         Width           =   5625
         Begin VB.CommandButton cmdAddSysop 
            Caption         =   "Add Sysop"
            Enabled         =   0   'False
            Height          =   285
            Left            =   4200
            TabIndex        =   12
            Top             =   7980
            Width           =   1275
         End
         Begin MSComctlLib.ImageList ilSysops 
            Left            =   480
            Top             =   5610
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
                  Picture         =   "frmAgSites.frx":0442
                  Key             =   "k100"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAgSites.frx":075C
                  Key             =   "k080_099"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAgSites.frx":0A76
                  Key             =   "k060_079"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAgSites.frx":1350
                  Key             =   "k040_059"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAgSites.frx":17A2
                  Key             =   "k001_020"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAgSites.frx":1BF4
                  Key             =   "k020_039"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAgSites.frx":2046
                  Key             =   "k000"
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView lvSysops 
            Height          =   7575
            Left            =   150
            TabIndex        =   13
            Top             =   300
            Width           =   5325
            _ExtentX        =   9393
            _ExtentY        =   13361
            View            =   3
            Arrange         =   1
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393217
            Icons           =   "ilSysops"
            SmallIcons      =   "ilSysops"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Level"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Username"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Description"
               Object.Width           =   8819
            EndProperty
         End
      End
   End
   Begin MSComctlLib.TabStrip ts 
      Height          =   9135
      Left            =   4770
      TabIndex        =   2
      Top             =   90
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   16113
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Site Information"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sales Staff/Sysops"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sites"
      Height          =   9105
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   4545
      Begin MSComctlLib.ImageList ilTreeview 
         Left            =   3570
         Top             =   720
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   62
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":2498
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":2D72
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":364C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":3F26
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":4800
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":50DA
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":59B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":628E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":6B68
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":7442
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":7D1C
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":85F6
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":8ED0
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":97AA
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":A084
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":A95E
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":B238
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":BB12
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":C3EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":CCC6
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":D5A0
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":DE7A
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":E754
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":F02E
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":F908
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":101E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":10ABC
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":11396
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":11C70
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":1254A
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":12E24
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":136FE
               Key             =   ""
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":13FD8
               Key             =   ""
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":148B2
               Key             =   ""
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":1518C
               Key             =   ""
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":15A66
               Key             =   ""
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":16340
               Key             =   ""
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":16C1A
               Key             =   "book"
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":16F34
               Key             =   "news"
            EndProperty
            BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":1724E
               Key             =   "filter"
            EndProperty
            BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":17568
               Key             =   "world"
            EndProperty
            BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":17882
               Key             =   ""
            EndProperty
            BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":17CD4
               Key             =   ""
            EndProperty
            BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":17FEE
               Key             =   ""
            EndProperty
            BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":18440
               Key             =   ""
            EndProperty
            BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":18B92
               Key             =   ""
            EndProperty
            BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":18FE4
               Key             =   ""
            EndProperty
            BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":19436
               Key             =   ""
            EndProperty
            BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":19888
               Key             =   ""
            EndProperty
            BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":19CDA
               Key             =   ""
            EndProperty
            BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":1A12C
               Key             =   ""
            EndProperty
            BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":1A57E
               Key             =   ""
            EndProperty
            BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":1A9D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":1AE22
               Key             =   ""
            EndProperty
            BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":1B274
               Key             =   ""
            EndProperty
            BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":1B9C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":1BE18
               Key             =   ""
            EndProperty
            BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":1C26A
               Key             =   ""
            EndProperty
            BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":1C6BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":1CB0E
               Key             =   ""
            EndProperty
            BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":1CF60
               Key             =   ""
            EndProperty
            BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgSites.frx":1D3B2
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdCreate 
         Caption         =   "Create a Site for this Agency"
         Height          =   345
         Left            =   210
         TabIndex        =   10
         Top             =   8670
         Width           =   4125
      End
      Begin MSComctlLib.ListView lvSites 
         Height          =   8325
         Left            =   150
         TabIndex        =   1
         Top             =   270
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   14684
         View            =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         Icons           =   "ilTreeview"
         SmallIcons      =   "ilTreeview"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.PictureBox picTS 
      BorderStyle     =   0  'None
      Height          =   8715
      Index           =   0
      Left            =   4830
      ScaleHeight     =   8715
      ScaleWidth      =   5865
      TabIndex        =   3
      Top             =   450
      Width           =   5865
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   405
         Left            =   1830
         TabIndex        =   15
         Top             =   8220
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update Database"
         Enabled         =   0   'False
         Height          =   405
         Left            =   60
         TabIndex        =   14
         Top             =   8220
         Width           =   1695
      End
      Begin VB.Frame Frame4 
         Caption         =   "Quick Reports"
         Height          =   3825
         Left            =   60
         TabIndex        =   7
         Top             =   1920
         Width           =   5715
      End
      Begin VB.Frame Frame3 
         Caption         =   "Manager Contact Information"
         Height          =   1845
         Index           =   0
         Left            =   60
         TabIndex        =   5
         Top             =   30
         Width           =   5715
         Begin VB.TextBox txtField 
            Height          =   1395
            Index           =   0
            Left            =   120
            MaxLength       =   65535
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   270
            Width           =   5445
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Comments"
         Height          =   2235
         Left            =   60
         TabIndex        =   4
         Top             =   5820
         Width           =   5715
         Begin VB.TextBox txtField 
            Height          =   1815
            Index           =   1
            Left            =   120
            MaxLength       =   65535
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   270
            Width           =   5445
         End
      End
   End
End
Attribute VB_Name = "frmAgSites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lRecID As Long

Private Sub cmdAddSysop_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdAddSysop_Click"
    Const ContainerName = "frmAgSites"
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


    Dim ffrmSysop As frmSysopDetails
    Set ffrmSysop = New frmSysopDetails
    ffrmSysop.Show 1
    
    If ffrmSysop.iCloseState = frmCloseSave Then
        Dim itmX As ListItem
        Set itmX = lvSysops.ListItems.Add(, , String(3 - Len("" & ffrmSysop.byLevel), "0") + "" & ffrmSysop.byLevel)
        itmX.SubItems(1) = ffrmSysop.sUsername
        itmX.SubItems(2) = ffrmSysop.sDescription
        itmX.Tag = ffrmSysop.sPassword
        
        Dim lSysopID As Long
    
        On Error Resume Next
        Do
            Err.Clear
            lSysopID = MySQL.GetTMPRecID("sysops", ADOConn)
            MySQL.Execute ADOConn, "INSERT INTO sysops (Password, Username, Description, RecID, SecurityLevel, VirtualID, AgencyID) VALUES(encode('" & ffrmSysop.sPassword & "','" & odb.colSalts.ReturnSalt(PWSalt) & "'), '" & ffrmSysop.sUsername & "', '" + MySQL.ESC(ffrmSysop.sDescription) + "'," & lSysopID & "," & ffrmSysop.byLevel & "," & lRecID & "," & Login.lAgencyID & ")"
            If Err.Number > 0 Then cDebug Err.Description
        Loop Until Err.Number = 0
                
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
            If ffrmSysop.byLevel >= imgMin And ffrmSysop.byLevel <= imgMax Then
                itmX.SmallIcon = imgX
            End If
        Next imgX
        
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

Private Sub cmdCreate_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdCreate_Click"
    Const ContainerName = "frmAgSites"
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


    Dim fCAg As New frmCAgency
    Dim itmX As ListItem
    Dim rsSave As ADODB.Recordset
    
    fCAg.Caption = "Create Agency Site"
    fCAg.Show 1
    
    If fCAg.stName <> "" Then
        lRecID = MySQL.GetTMPRecID("virtualisp", ADOConn)
        MySQL.Execute ADOConn, "Insert into virtualisp (RecID, Description, Realm, ABN, ACN, Subscribed, JoiningFee, AgencyID, Icon, NextCycle, PreviousCycle, CreatedBy_SysopID) " & _
                                                        "VALUES ('" & lRecID & "','" & fCAg.stName & "','" & Login.sVISPDomain & "','" & "87096867775" & "','" & "096867775" & "','" & "100" & "','" & InputBox("Please enter the joining fee, without currency symbols (ie '0.00')", "Administration Joining Fee", "0.00") & _
                                                        "','" & Login.lAgencyID & "','" & fCAg.stIcon & "','" & Format(DateAdd("m", 1, sysNOW), "yyyy-mm-dd ttttt") & "'," & "NOW()" & ",'" & Login.lSysopID & "')"
        Dim lSysopID As Long
    
        On Error Resume Next
        Do
            Err.Clear
            lSysopID = MySQL.GetTMPRecID("sysops", ADOConn)
            MySQL.Execute ADOConn, "UPDATE virtualisp SET SysopID=" & lSysopID & " Where RecID = " & lRecID
            MySQL.Execute ADOConn, "INSERT INTO sysops (Password, Username, Description, RecID, SecurityLevel, VirtualID, AgencyID) VALUES(AES_ENCRYPT('" & fCAg.stPassword & "','" & odb.colSalts.ReturnSalt(PWSalt) & "'), '" & fCAg.stUsername & "', '" + MySQL.ESC(fCAg.stDesc) + "'," & lSysopID & ",75," & lRecID & "," & Login.lAgencyID & ")"
            If Err.Number > 0 Then cDebug Err.Description
        Loop Until Err.Number = 0
        
        lvSites.ListItems.Add , "s" & lRecID, fCAg.stName, fCAg.stIcon, fCAg.stIcon
        
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

Private Sub cmdUpdate_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdUpdate_Click"
    Const ContainerName = "frmAgSites"
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


    MySQL.Execute ADOConn, "Update VirutalISP Manager = '" & MySQL.ESC(txtField(0)) & "', Comment = '" & MySQL.ESC(txtField(1)) & "' where RecID = " & Mid(lvSites.SelectedItem.Key, 2)
        
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

Private Sub Command1_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Command1_Click"
    Const ContainerName = "frmAgSites"
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


    For ix = ilTreeview.ListImages.Count To 1 Step -1
        SavePicture ilTreeview.ListImages(ix).Picture, "z:\folder" & ix & ".ico"
    Next
    
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
    Const ContainerName = "frmAgSites"
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


    Dim rsOpen As ADODB.Recordset
    
    If MySQL.OpenTable(ADOConn, rsOpen, , "select RecID, Description, Icon from virtualisp where AgencyID = " & Login.lAgencyID) = True Then
        If rsOpen.BOF And rsOpen.EOF Then
        
        Else
            While Not rsOpen.EOF And Err.Number = 0
                lvSites.ListItems.Add , "s" & rsOpen!RecID, rsOpen!Description, Val(rsOpen!Icon), Val(rsOpen!Icon)
                rsOpen.MoveNext
            Wend
            
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

Private Sub lvSites_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvSites_ItemClick"
    Const ContainerName = "frmAgSites"
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


    lRecID = CLng(Mid(Item.Key, 2))
    
    lvSysops.ListItems.Clear
    
    cmdAddSysop.Enabled = True
    cmdUpdate.Enabled = True
    
    Dim rsOpen As ADODB.Recordset
    Dim itmX As ListItem
    
    If MySQL.OpenTable(ADOConn, rsOpen, , "select *,decode(Password,'" & odb.colSalts.ReturnSalt(PWSalt) & "') as DecPassword from sysops where VirtualID = " & lRecID) = True Then
        If rsOpen.BOF And rsOpen.EOF Then
                   
        Else
            While Not rsOpen.EOF And Err.Number = 0
            
                Set itmX = lvSysops.ListItems.Add(, "s" & rsOpen!RecID, String(3 - Len("" & rsOpen!SecurityLevel), "0") + "" & rsOpen!SecurityLevel)
            
                    itmX.SubItems(1) = rsOpen!Username
                    itmX.SubItems(2) = rsOpen!Description
                    itmX.Tag = rsOpen!DecPassword
                
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
                        If Val(rsOpen!SecurityLevel) >= imgMin And Val(rsOpen!SecurityLevel) <= imgMax Then
                            itmX.SmallIcon = imgX
                            Exit For
                        End If
                    Next imgX
                
                rsOpen.MoveNext
            Wend
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

Private Sub lvsysops_DblClick()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvsysops_DblClick"
    Const ContainerName = "frmAgSites"
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



    If lvSysops.SelectedItem Is Nothing Then
    
    Else
    
        Dim ffrmSysop As frmSysopDetails
        Set ffrmSysop = New frmSysopDetails
        Dim itmX As ListItem
        Set itmX = lvSysops.SelectedItem
        
        ffrmSysop.sUsername = itmX.SubItems(1)
        ffrmSysop.sPassword = itmX.Tag
        ffrmSysop.sDescription = itmX.SubItems(2)
        ffrmSysop.byLevel = itmX.Text
        ffrmSysop.Show 1
        
        If ffrmSysop.iCloseState = frmCloseSave Then
            itmX.Text = String(3 - Len("" & ffrmSysop.byLevel), "0") + "" & ffrmSysop.byLevel
            itmX.SubItems(1) = ffrmSysop.sUsername
            itmX.SubItems(2) = ffrmSysop.sDescription
            itmX.Tag = ffrmSysop.sPassword
            
            MySQL.Execute ADOConn, "UPDATE sysops SET Password=encode('" & ffrmSysop.sPassword & "','" & odb.colSalts.ReturnSalt(PWSalt) & "'), Username = '" & ffrmSysop.sUsername & "', Description = '" & MySQL.ESC(ffrmSysop.sDescription) & "', SecurityLevel = " & ffrmSysop.byLevel & " where RecID = " & Mid(lvSysops.SelectedItem.Key, 2)
            
        End If
    
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
            If ffrmSysop.byLevel >= imgMin And ffrmSysop.byLevel <= imgMax Then
                itmX.SmallIcon = imgX
            End If
        Next imgX
    
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
    Const ContainerName = "frmAgSites"
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
    
    For ix = 1 To ts.Tabs.Count
        If ts.SelectedItem.Index <> ix Then
            picTS(ts.TabIndex - 1).Visible = False
        Else
        
        End If
    Next ix
    
    picTS(ts.SelectedItem.Index - 1).Move ts.clientLeft, ts.clientTop, ts.clientWidth, ts.clientHeight
    picTS(ts.SelectedItem.Index - 1).ZOrder 0
    picTS(ts.SelectedItem.Index - 1).Visible = True
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
