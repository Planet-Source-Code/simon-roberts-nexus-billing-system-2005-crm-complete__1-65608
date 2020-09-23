VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmBot 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maintenance & Upkeep"
   ClientHeight    =   3765
   ClientLeft      =   4845
   ClientTop       =   5085
   ClientWidth     =   11580
   Icon            =   "frmBot2_re_new.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   11580
   Begin VB.PictureBox SMTP3 
      Height          =   480
      Left            =   750
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   38
      Top             =   2670
      Width           =   1200
   End
   Begin VB.PictureBox SMTP2 
      Height          =   480
      Left            =   180
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   39
      Top             =   3180
      Width           =   1200
   End
   Begin VB.Timer tmFile1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   180
      Top             =   1770
   End
   Begin VB.PictureBox SMTP1 
      Height          =   480
      Left            =   180
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   40
      Top             =   2670
      Width           =   1200
   End
   Begin VB.Timer tmProgress 
      Interval        =   5
      Left            =   180
      Top             =   2220
   End
   Begin VB.PictureBox picBots 
      BorderStyle     =   0  'None
      Height          =   3165
      Index           =   4
      Left            =   210
      ScaleHeight     =   3165
      ScaleWidth      =   8715
      TabIndex        =   22
      Top             =   450
      Visible         =   0   'False
      Width           =   8715
      Begin VB.Frame Frame4 
         Caption         =   "Statistics"
         Height          =   1695
         Index           =   1
         Left            =   1230
         TabIndex        =   26
         Top             =   870
         Width           =   7425
         Begin VB.Label Label1 
            Height          =   1245
            Left            =   210
            TabIndex        =   27
            Top             =   330
            Width           =   7095
         End
      End
      Begin VB.CommandButton cmdViSP 
         Caption         =   "Process ViSP's"
         Height          =   390
         Left            =   7110
         TabIndex        =   25
         Tag             =   "b0"
         Top             =   2700
         Width           =   1515
      End
      Begin VB.Frame Frame5 
         Caption         =   "Searching for ViSP Billing Date"
         Height          =   795
         Index           =   1
         Left            =   1230
         TabIndex        =   23
         Top             =   30
         Width           =   7425
         Begin MSComctlLib.ProgressBar pbViSP 
            Height          =   315
            Left            =   150
            TabIndex        =   24
            Top             =   330
            Width           =   7155
            _ExtentX        =   12621
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
         End
      End
      Begin VB.Image Image1 
         Height          =   960
         Index           =   4
         Left            =   90
         OLEDropMode     =   1  'Manual
         Picture         =   "frmBot2_re_new.frx":0442
         Stretch         =   -1  'True
         Top             =   90
         Width           =   975
      End
   End
   Begin VB.PictureBox picBots 
      BorderStyle     =   0  'None
      Height          =   3165
      Index           =   2
      Left            =   180
      ScaleHeight     =   3165
      ScaleWidth      =   8715
      TabIndex        =   10
      Top             =   420
      Visible         =   0   'False
      Width           =   8715
      Begin VB.Frame Frame4 
         Caption         =   "Statistics"
         Height          =   1695
         Index           =   0
         Left            =   1230
         TabIndex        =   14
         Top             =   870
         Width           =   7425
         Begin VB.Label lblBillingStats 
            Height          =   1245
            Left            =   210
            TabIndex        =   15
            Top             =   330
            Width           =   7095
         End
      End
      Begin VB.CommandButton cmdProcessBills 
         Caption         =   "Process Bills Now"
         Height          =   390
         Left            =   7110
         TabIndex        =   13
         Tag             =   "b0"
         Top             =   2700
         Width           =   1515
      End
      Begin VB.Frame Frame5 
         Caption         =   "Searching For Billing Cycle Date Match"
         Height          =   795
         Index           =   0
         Left            =   1230
         TabIndex        =   11
         Top             =   30
         Width           =   7425
         Begin MSComctlLib.ProgressBar pbBilling 
            Height          =   315
            Left            =   150
            TabIndex        =   12
            Top             =   330
            Width           =   7155
            _ExtentX        =   12621
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
         End
      End
      Begin VB.Image Image1 
         Height          =   960
         Index           =   2
         Left            =   90
         OLEDropMode     =   1  'Manual
         Picture         =   "frmBot2_re_new.frx":0884
         Stretch         =   -1  'True
         Top             =   90
         Width           =   975
      End
   End
   Begin VB.PictureBox picBots 
      BorderStyle     =   0  'None
      Height          =   3165
      Index           =   1
      Left            =   150
      ScaleHeight     =   3165
      ScaleWidth      =   8775
      TabIndex        =   5
      Top             =   450
      Visible         =   0   'False
      Width           =   8775
      Begin VB.Frame Frame1 
         Caption         =   "Statistics"
         Height          =   1695
         Left            =   1230
         TabIndex        =   28
         Top             =   900
         Width           =   7425
         Begin VB.TextBox txtStats 
            Height          =   1305
            Left            =   150
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   29
            Top             =   270
            Width           =   7155
         End
      End
      Begin VB.Timer tmrFile2 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   330
         Top             =   1590
      End
      Begin VB.CommandButton cmdUploadUsers 
         Caption         =   "Update Now"
         Height          =   390
         Left            =   7080
         TabIndex        =   8
         Top             =   2670
         Width           =   1515
      End
      Begin VB.Frame frameUsers 
         Caption         =   "Updating User Tables on Radius Server"
         Height          =   795
         Left            =   1230
         TabIndex        =   6
         Top             =   30
         Width           =   7425
         Begin MSComctlLib.ProgressBar pbRadUpdate 
            Height          =   315
            Left            =   150
            TabIndex        =   7
            Top             =   330
            Width           =   7155
            _ExtentX        =   12621
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
         End
      End
      Begin VB.Image Image1 
         Height          =   1110
         Index           =   1
         Left            =   150
         OLEDropMode     =   1  'Manual
         Picture         =   "frmBot2_re_new.frx":0CC6
         Stretch         =   -1  'True
         Top             =   60
         Width           =   975
      End
   End
   Begin MSComctlLib.TabStrip tsBots 
      Height          =   3645
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   6429
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Retrieve Radius Log"
            Object.ToolTipText     =   "Download and parses the radius log file."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Updating Radius"
            Object.ToolTipText     =   "Updates the Radius User Log File"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Billing Cycle"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Auto-generated Emails"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Virtual ISP"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Unpaid Client"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picBots 
      BorderStyle     =   0  'None
      Height          =   3165
      Index           =   0
      Left            =   120
      ScaleHeight     =   3165
      ScaleWidth      =   11085
      TabIndex        =   1
      Top             =   450
      Width           =   11085
      Begin VB.Frame Frame3 
         Caption         =   "Statistics"
         Height          =   2025
         Left            =   1230
         TabIndex        =   9
         Top             =   990
         Width           =   7605
         Begin VB.TextBox txtRadDown 
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1665
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   30
            Top             =   270
            Width           =   7365
         End
      End
      Begin VB.CommandButton cmdRadiusUpdate 
         Caption         =   "Update Now"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   8910
         TabIndex        =   4
         Tag             =   "radius"
         Top             =   2400
         Width           =   2085
      End
      Begin VB.Frame Frame2 
         Caption         =   "Scanning Database for Radius Information"
         Height          =   795
         Left            =   1230
         TabIndex        =   2
         Top             =   120
         Width           =   9825
         Begin MSComctlLib.ProgressBar pbTransmutate 
            Height          =   315
            Left            =   150
            TabIndex        =   3
            Top             =   330
            Width           =   9555
            _ExtentX        =   16854
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
         End
      End
      Begin VB.Image Image1 
         Height          =   1110
         Index           =   0
         Left            =   120
         OLEDropMode     =   1  'Manual
         Picture         =   "frmBot2_re_new.frx":1590
         Stretch         =   -1  'True
         Top             =   60
         Width           =   975
      End
   End
   Begin VB.PictureBox picBots 
      BorderStyle     =   0  'None
      Height          =   3165
      Index           =   3
      Left            =   150
      ScaleHeight     =   3165
      ScaleWidth      =   11025
      TabIndex        =   16
      Top             =   420
      Visible         =   0   'False
      Width           =   11025
      Begin VB.Frame frameEmails 
         Caption         =   "Searching and Sending Quota Messages"
         Height          =   795
         Left            =   1230
         TabIndex        =   20
         Top             =   30
         Width           =   9765
         Begin MSComctlLib.ProgressBar pbQuota 
            Height          =   315
            Left            =   150
            TabIndex        =   21
            Top             =   330
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Process Quota"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8160
         TabIndex        =   19
         Tag             =   "q0"
         Top             =   2670
         Width           =   2835
      End
      Begin VB.Frame Frame6 
         Caption         =   "Statistics"
         Height          =   1695
         Left            =   1230
         TabIndex        =   17
         Top             =   870
         Width           =   9765
         Begin VB.Label lblMessageStats 
            Height          =   1245
            Left            =   210
            TabIndex        =   18
            Top             =   330
            Width           =   9405
         End
      End
      Begin VB.Image Image1 
         Height          =   960
         Index           =   3
         Left            =   90
         OLEDropMode     =   1  'Manual
         Picture         =   "frmBot2_re_new.frx":19D2
         Stretch         =   -1  'True
         Top             =   90
         Width           =   975
      End
   End
   Begin VB.PictureBox picBots 
      BorderStyle     =   0  'None
      Height          =   3165
      Index           =   5
      Left            =   180
      ScaleHeight     =   3165
      ScaleWidth      =   8775
      TabIndex        =   31
      Top             =   480
      Visible         =   0   'False
      Width           =   8775
      Begin VB.Frame Frame8 
         Caption         =   "Results"
         Height          =   2205
         Left            =   1260
         TabIndex        =   34
         Top             =   870
         Width           =   7455
         Begin VB.CommandButton cmdUnpaid 
            Caption         =   "Run unpaid customer search"
            Height          =   345
            Left            =   180
            TabIndex        =   37
            Tag             =   "kunpaid"
            Top             =   1770
            Width           =   2415
         End
         Begin VB.CommandButton cmdReset 
            Caption         =   "Clear Results List"
            Height          =   345
            Left            =   5100
            TabIndex        =   36
            Top             =   1770
            Width           =   2235
         End
         Begin MSComctlLib.ListView lvUnpaid 
            Height          =   1455
            Left            =   150
            TabIndex        =   35
            Top             =   270
            Width           =   7185
            _ExtentX        =   12674
            _ExtentY        =   2566
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "VISP"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Account Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Username"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Search for Unpaid Customers"
         Height          =   795
         Left            =   1260
         TabIndex        =   32
         Top             =   60
         Width           =   7455
         Begin MSComctlLib.ProgressBar pbUnpaid 
            Height          =   285
            Left            =   120
            TabIndex        =   33
            Top             =   330
            Width           =   7245
            _ExtentX        =   12779
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
         End
      End
      Begin VB.Image Image2 
         Height          =   960
         Left            =   120
         Picture         =   "frmBot2_re_new.frx":1E14
         Stretch         =   -1  'True
         Top             =   60
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmBot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim bDownload As Boolean
Dim bUpload As Boolean
Dim oADORadius As adodb.Connection

Dim cSMTP As New smtp

Public iFrmState As Boolean

Public Sub cmdProcessBills_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdProcessBills_Click"
    Const ContainerName = "frmBot"
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


    tsBots.Tabs(3).Selected = True
    Frame5(0).Caption = "Searching For Services Billing Cycle"
    Call ProcessCycle
    Pause
    Frame5(0).Caption = "Searching for Invoice Cycle"
    Call ProcessStatements
    
    pbQuota.Value = 1
    pbQuota.Max = 1
    
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

'
'  This subroutine downloads the Log Flat File for users
'
'  This routine needs to be changed to accommodate a Database driven radius server not a
'  Flat file driven one.
'
'
'

Public Sub cmdRadiusUpdate_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdRadiusUpdate_Click"
    Const ContainerName = "frmBot"
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


    'On Error GoTo ErrorOccur
    
    Exit Sub
    
    'Dim rsPools As ADODB.Recordset
    
    Dim rsAccInfo As adodb.Recordset
    Dim rsRadius As adodb.Recordset
    Dim rsLog As adodb.Recordset
    Dim rsAlive As adodb.Recordset
    
    
    Dim lx As Long
    Dim br As Boolean
    Dim bFound As Boolean
        
    pbTransmutate.Value = 0
    pbTransmutate.Max = 1
    
    'If MySQL.DAdirectConn("Radius", sServer, sUID, sPWD, directConn, wrkODBC) = False Then Exit Sub
        
        
    'If MySQL.OpenTable(directConn, rsAlive, , "select RadiusLogTwo. RadiusLogTwo.AcctSessionTime, RadiusLogTwo.`Acct-Session-Id`, RadiusLogTwo.`Acct-Session-Id`, RadiusLogTwo.Username, RadiusLogTwo.SessionStart from RadiusLogTwo LEFT JOiN RadiusLogTwo ON RadiusLogTwo.Acct-Session-Id = RadiusLog.Acct-Session-Id WHERE RadiusLog.Acct-Session-Id is NULL") = True Then
    '    If rsAlive.RecordCount > 0 Then
    '        pbTransmutate.Max = pbTransmutate.Max + rsAlive.RecordCount
    '        While Not rsAlive.EOF
    '            pbTransmutate.Value = pbTransmutate.Value + 1
    '            directConn.Execute "Insert Into RadiusLogTwo (Acct-Status-Type,`Acct-Session-Id`, `Acct-Session-Id`, Username, SessionStart) VALUES ('Start','" & rsAlive!Acct-Session-Id & "','" & rsAlive!Acct-Session-Id & "','" & rsAlive!Username & "','" & Format(DateAdd("s", -IIf(IsNull(rsAlive!AcctSessionTime), 0, rsAlive!AcctSessionTime), IIf(IsNull(rsAlive!SessionStart), sysNow, rsAlive!SessionStart)), "yyyy-mm-dd Hh:Nn:Ss") & "')"
    '            rsAlive.MoveNext
    '        Wend
    '    End If
    '
    'End If
    
    ' This first section populate the class with any alive records that have occured in the system between runtimes of The Nexus
    
    br = MySQL.OpenTable(directConn, rsLog, , "select `Acct-Session-Id`, `User-Name`, SessionStart from RadiusLogTwo Where `Acct-Status-Type` = 'Start' or `Acct-Status-Type` = 'Alive'")
    
    Dim iCount As Integer
    
    If rsLog.RecordCount > 0 Then
        
        pbTransmutate.Max = pbTransmutate.Max + rsLog.RecordCount
        While Not rsLog.EOF And Err.Number = 0
            bFound = False
            If cRadius.Count > 0 Then
                For lx = 1 To cRadius.Count
                    If cRadius(lx).UniqueSessionID = rsLog("Acct-Session-ID") Then
                        bFound = True
                        Exit For
                    End If
                Next
            End If
            
            If bFound = False Then
                'br = MySQL.OpenTable(directConn, rsLoad, , "select `Acct-Session-Id`, `Acct-Session-Id`, Username, SessionStart from RadiusLogTwo Where RadAcctId = " & rsLog!RadAcctId)
                cRadius.Add rsLog("Acct-Session-ID"), rsLog("Acct-Session-ID"), rsLog("Acct-Session-ID"), rsLog("User-Name"), IIf(IsNull(rsLog!SessionStart), Format(sysnow, "dd-mm-yyyy Hh:Nn:Ss"), rsLog!SessionStart)
                iCount = iCount + 1
            End If
            pbTransmutate.Value = pbTransmutate.Value + 1
            rsLog.MoveNext
        Wend
    End If
    
    txtRadDown.Text = "" & iCount & " - New Start Records Found"
    iCount = 0
    
    
    ' This routine looks for updated alive records or alive records that have changed since last scanning the Database
    
    If cRadius.Count > 0 Then
        pbTransmutate.Max = pbTransmutate.Max + cRadius.Count
        For lx = cRadius.Count To 1 Step -1
            If MySQL.OpenTable(directConn, rsAlive, , "select chkSessionStart, Timestamp, SessionStart, `Acct-Input-Octets` , `chk-Input-Octets` , `Acct-Output-Octets`, `chk-Output-Octets` from RadiusLogTwo Where `Acct-Session-Id` = '" & cRadius(lx).UniqueSessionID & "' and (`chk-Output-Octets` <> `Acct-Output-Octets` or `chk-Input-Octets`  <> `Acct-Input-Octets` or chkSessionStart <> Timestamp)") = True Then
                If MySQL.OpenTable(directConn, rsLog, , "select RadAcctId as RecID, SessionStart, chkSessionStart, Timestamp, `Acct-Input-Octets` , `Acct-Output-Octets` from RadiusLogTwo Where `Acct-Session-Id` = '" & cRadius(lx).UniqueSessionID & "'") = True Then
            
                
                    If rsAlive.RecordCount > 0 Then
                        'pbTransmutate.Max = pbTransmutate.Max + rsAlive.RecordCount
                        br = MySQL.OpenTable(directConn, rsRadius, , "select * from radiusaccounts Where Username = '" & cRadius(lx).Username & "'")
                        If rsRadius.RecordCount > 0 Then
                            While Not rsAlive.EOF And Err.Number = 0
                                On Error Resume Next
                                
                                If rsAlive!TimeStamp <> rsAlive!chkSessionStart Then
                                    If IsDate(rsAlive!TimeStamp) And IsDate(rsAlive!chkSessionStart) Then
                                        MySQL.Execute directConn, "Update accountinfo set sfCycle_Mins = sfCycle_Mins + " & DateDiff("n", CDate(IIf(IsDate(rsLog!chkSessionStart), rsLog!chkSessionStart, IIf(IsNull(rsLog!SessionStart), rsAlive!TimeStamp, rsLog!SessionStart))), CDate(rsAlive!TimeStamp)) & " where RecID = " & rsRadius!acci_RecID
                                        MySQL.Execute directConn, "Update radiusaccounts set sfCycle_Mins = sfCycle_Mins + " & DateDiff("n", CDate(IIf(IsDate(rsLog!chkSessionStart), rsLog!chkSessionStart, IIf(IsNull(rsLog!SessionStart), rsAlive!TimeStamp, rsLog!SessionStart))), CDate(rsAlive!TimeStamp)) & " where RecID = " & rsRadius!RecID
                                    End If
                                    
                                    MySQL.Execute directConn, "Insert Into history_radius_datausage (RadiusID, Uploaded, Downloaded, NumMins) VALUES(" & IIf(IsNull(rsRadius!RecID), 0, rsRadius!RecID) & ", " & (Oct(rsAlive("Acct-Input-Octets")) - Oct(rsAlive("chk-Input-Octets"))) & ", " & (Oct(rsAlive("Acct-Output-Octets")) - Oct(rsAlive("chk-Output-Octets"))) & ", " & DateDiff("n", CDate(IIf(IsDate(rsLog!chkSessionStart), rsLog!chkSessionStart, IIf(IsNull(rsLog!SessionStart), rsAlive!TimeStamp, rsLog!SessionStart))), CDate(rsAlive!TimeStamp)) & ")"
                                    
                                Else
                                
                                    MySQL.Execute directConn, "Insert Into history_radius_datausage (RadiusID, Uploaded, Downloaded, NumMins) VALUES(" & IIf(IsNull(rsRadius!RecID), 0, rsRadius!RecID) & ", " & (Oct(rsAlive("Acct-Input-Octets")) - Oct(rsAlive("chk-Input-Octets"))) & ", " & (Oct(rsAlive("Acct-Output-Octets")) - Oct(rsAlive("chk-Output-Octets"))) & ", 0)"
                                
                                End If
                                
                                
                                MySQL.Execute directConn, "Update radiusaccounts set sfCycle_Upload = sfCycle_Upload + " & (Oct(rsAlive("Acct-Input-Octets")) - Oct(rsAlive("chk-Input-Octets"))) & " where RecID = " & rsRadius!RecID
                                MySQL.Execute directConn, "Update radiusaccounts set sfCycle_Download = sfCycle_Download + " & (Oct(rsAlive("Acct-Output-Octets")) - Oct(rsAlive("chk-Output-Octets"))) & " where RecID = " & rsRadius!RecID
                                
                                MySQL.Execute directConn, "Update accountinfo set sfCycle_Upload = sfCycle_Upload + " & (Oct(rsAlive("Acct-Input-Octets")) - Oct(rsAlive("chk-Input-Octets"))) & " where RecID = " & rsRadius!acci_RecID
                                MySQL.Execute directConn, "Update accountinfo set sfCycle_Download = sfCycle_Download + " & (Oct(rsAlive("Acct-Output-Octets")) - Oct(rsAlive("chk-Output-Octets"))) & " where RecID = " & rsRadius!acci_RecID
                                MySQL.Execute directConn, "Update accountinfo set sfStartTime = '" & rsAlive("SessionStart") & "' where RecID = " & rsRadius!acci_RecID
                                
                                MySQL.Execute oADORadius, "Update RadiusLogTwo set `chk-Input-Octets` =`Acct-Input-Octets` , `chk-Output-Octets`=`Acct-Output-Octets`, chkSessionStart=Timestamp Where `Acct-Session-Id` = '" & cRadius(lx).UniqueSessionID & "'"
                                
                                If Err.Number <> 0 Then
                                    cDebug "[" & Err.Number & "]" & Err.Description
                                    Err.Clear
                                End If
                                iCount = iCount + 1
                               ' pbTransmutate.Value = pbTransmutate.Value + 1
                                rsAlive.MoveNext
                            Wend
                            On Error GoTo ErrorOccur
                        End If
                    End If
                End If
            End If
            pbTransmutate.Value = pbTransmutate.Value + 1
        Next
    End If
    
    On Error GoTo 0
    
    txtRadDown.Text = txtRadDown.Text & vbCrLf & iCount & " - Alive Records Found"
    Dim rsload2 As adodb.Recordset
    Dim rsLoad3 As adodb.Recordset
    
    'TableRef Count
    'radiusaccount 68
    'accountinfo 279
    
    br = MySQL.OpenTable(directConn, rsLog, , "select RadAcctID as RecID, `Acct-Session-ID`, FlagID, `User-Name`, SessionStart, SessionStop, chkSessionStart, `Acct-Output-Octets`, `Acct-Input-Octets`, `chk-Input-Octets`, `chk-Output-Octets`, `Acct-Session-Time` from RadiusLogTwo Where `Acct-Status-Type` = 'Stop' and FlagID = 0")
    iCount = 0
    If rsLog.RecordCount > 0 Then
        pbTransmutate.Max = pbTransmutate.Max + rsLog.RecordCount
        While Not rsLog.EOF And Err.Number = 0
            If MySQL.OpenTable(directConn, rsRadius, , "select * from radiusaccounts Where Username = '" + rsLog("User-Name") + "'") = True Then
            
                If rsRadius.RecordCount > 0 Then
                    br = MySQL.OpenTable(directConn, rsAccInfo, , "select * from accountinfo Where RecID = " & rsRadius!acci_RecID & "")
                    'br = MySQL.OpenTable(directConn, rsLoad2, , "select `Acct-Session-ID` from RadiusLogTwo Where RadAcctID = '" & IIf(IsNull(rsLog!RecID), 0, rsLog!RecID) & "'")
                    
                    'br = MySQL.OpenTable(directConn, rsAlive, , "select chkSessionStart,  SessionStart, `Acct-Input-Octets` , `chk-Input-Octets` , `Acct-Output-Octets`, `chk-Output-Octets`  from RadiusLogTwo Where Acct-Session-Id = '" & rsLoad2("Acct-Session-ID") & "'")
                    
                    '    If rsAlive.RecordCount = 0 Then
                    '        rsAccInfo!sfCycle_Mins = IIf(IsNull(rsAccInfo!sfCycle_Mins), 0, rsAccInfo!sfCycle_Mins) + DateDiff("n", rsLog!SessionStart, rsLog!SessionStop)
                    '        sfCycle_Mins = IIf(IsNull(sfCycle_Mins), 0, sfCycle_Mins) + DateDiff("n", rsLog!SessionStart, rsLog!SessionStop)
                    '    Else
                    '        rsAccInfo!sfCycle_Mins = IIf(IsNull(rsAccInfo!sfCycle_Mins), 0, rsAccInfo!sfCycle_Mins) + DateDiff("n", rsAlive!SessionStart, rsLog!SessionStop)
                    '        sfCycle_Mins = IIf(IsNull(sfCycle_Mins), 0, sfCycle_Mins) + DateDiff("n", rsAlive!SessionStart, rsLog!SessionStop)
                    '    End If
    
                        'If rsAlive.RecordCount > 0 Then
                        '    sfCycle_Upload = IIf(IsNull(sfCycle_Upload), 0, sfCycle_Upload) + (Oct(IIf(IsNull(rsLog!Acct-Input-Octets ), 0, rsLog!Acct-Input-Octets )) - Oct(IIf(IsNull(rsAlive("Acct-Input-Octets") ), 0, rsAlive("Acct-Input-Octets") )))
                        '    sfCycle_Download = IIf(IsNull(sfCycle_Download), 0, sfCycle_Download) + (Oct(IIf(IsNull(rsLog("Acct-Output-Octets")), 0, rsLog("Acct-Output-Octets"))) - Oct(IIf(IsNull(rsAlive("Acct-Output-Octets")), 0, rsAlive("Acct-Output-Octets"))))
                        '    rsAccInfo!sfCycle_Upload = IIf(IsNull(rsAccInfo!sfCycle_Upload), 0, rsAccInfo!sfCycle_Upload) + (Oct(IIf(IsNull(rsLog!Acct-Input-Octets ), 0, rsLog!Acct-Input-Octets )) - Oct(IIf(IsNull(rsAlive("Acct-Input-Octets") ), 0, rsAlive("Acct-Input-Octets") )))
                        '    rsAccInfo!sfCycle_Download = IIf(IsNull(rsAccInfo!sfCycle_Download), 0, rsAccInfo!sfCycle_Download) + (Oct(IIf(IsNull(rsLog("Acct-Output-Octets")), 0, rsLog("Acct-Output-Octets"))) - Oct(IIf(IsNull(rsAlive("Acct-Output-Octets")), 0, rsAlive("Acct-Output-Octets"))))
                        'Else
                        '    sfCycle_Upload = IIf(IsNull(sfCycle_Upload), 0, sfCycle_Upload) + (Oct(IIf(IsNull(rsLog!Acct-Input-Octets ), 0, rsLog!Acct-Input-Octets )))
                        '    sfCycle_Download = IIf(IsNull(sfCycle_Download), 0, sfCycle_Download) + (Oct(IIf(IsNull(rsLog("Acct-Output-Octets")), 0, rsLog("Acct-Output-Octets"))))
                        '    rsAccInfo!sfCycle_Upload = IIf(IsNull(rsAccInfo!sfCycle_Upload), 0, rsAccInfo!sfCycle_Upload) + (Oct(IIf(IsNull(rsLog!Acct-Input-Octets ), 0, rsLog!Acct-Input-Octets )))
                        '    rsAccInfo!sfCycle_Download = IIf(IsNull(rsAccInfo!sfCycle_Download), 0, rsAccInfo!sfCycle_Download) + (Oct(IIf(IsNull(rsLog("Acct-Output-Octets")), 0, rsLog("Acct-Output-Octets"))))
                        'End If
                        
                        'rsradius.Update
                        'rsAccInfo.Update
                            If IsDate(rsLog!SessionStop) And IsDate(rsLog!chkSessionStart) Then
                                MySQL.Execute directConn, "Update accountinfo set sfCycle_Mins = sfCycle_Mins + " & DateDiff("n", CDate(rsLog!SessionStop), CDate(rsLog!chkSessionStart)) & " where RecID = " & rsRadius!acci_RecID
                                MySQL.Execute directConn, "Update radiusaccounts set sfCycle_Mins = sfCycle_Mins + " & DateDiff("n", CDate(rsLog!SessionStop), CDate(rsLog!chkSessionStart)) & " where RecID = " & rsRadius!RecID
                            End If
                            
                            MySQL.Execute directConn, "Insert Into history_radius_datausage (RadiusID, Uploaded, Downloaded, NumMins) VALUES(" & IIf(IsNull(rsRadius!RecID), 0, rsRadius!RecID) & ", " & (Oct(rsLog("Acct-Input-Octets")) - Oct(rsLog("chk-Input-Octets"))) & ", " & (Oct(rsLog("Acct-Output-Octets")) - Oct(rsLog("chk-Output-Octets"))) & ", " & DateDiff("n", CDate(rsLog!SessionStop), CDate(IIf(IsDate(rsLog!chkSessionStart), rsLog!chkSessionStart, IIf(IsDate(rsLog!SessionStart), rsLog!SessionStart, DateAdd("s", -rsLog("Acct-Session-Time"), rsLog!SessionStop))))) & ")"
                                                        
                            MySQL.Execute directConn, "Update radiusaccounts set sfCycle_Upload = sfCycle_Upload + " & (Oct(rsLog("Acct-Input-Octets")) - Oct(rsLog("chk-Input-Octets"))) & " where RecID = " & rsRadius!RecID
                            MySQL.Execute directConn, "Update radiusaccounts set sfCycle_Download = sfCycle_Download + " & (Oct(rsLog("Acct-Output-Octets")) - Oct(rsLog("chk-Output-Octets"))) & " where RecID = " & rsRadius!RecID
                            
                            MySQL.Execute directConn, "Update accountinfo set sfCycle_Upload = sfCycle_Upload + " & (Oct(rsLog("Acct-Input-Octets")) - Oct(rsLog("chk-Input-Octets"))) & " where RecID = " & rsRadius!acci_RecID
                            MySQL.Execute directConn, "Update accountinfo set sfCycle_Download = sfCycle_Download + " & (Oct(rsLog("Acct-Output-Octets")) - Oct(rsLog("chk-Output-Octets"))) & " where RecID = " & rsRadius!acci_RecID
                            MySQL.Execute directConn, "Update accountinfo set sfStartTime = '" & rsLog("SessionStart") & "' where RecID = " & rsRadius!acci_RecID
                        
                        iCount = iCount + 1
                
                End If
            End If
            MySQL.Execute oADORadius, "update radius.RadiusLogTwo set FlagID = 1 where `Acct-Session-ID` = '" & rsLog("Acct-Session-ID") + "'"
            
            If cRadius.Count > 0 Then
                For lx = cRadius.Count To 1 Step -1
                    If cRadius(lx).UniqueSessionID = rsLog("Acct-Session-ID") Then
                        
                        cRadius.Remove lx
                        Exit For
                    End If
                Next
            End If
            
            MySQL.Execute oADORadius, "Delete from RadiusLogTwo where `Acct-Session-ID` = '" & rsLog("Acct-Session-ID") & "'"
            
            rsLog.MoveNext
            pbTransmutate.Value = pbTransmutate.Value + 1
        Wend
    End If
    
    txtRadDown.Text = txtRadDown.Text & vbCrLf & iCount & " - Stop Records Found"
    
    On Error Resume Next
    pbTransmutate.Value = pbTransmutate.Value + 1
    
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



Private Sub cmdReset_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdReset_Click"
    Const ContainerName = "frmBot"
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


    lvUnpaid.ListItems.Clear
    
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

Public Sub cmdUnpaid_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdUnpaid_Click"
    Const ContainerName = "frmBot"
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


    
        Dim rsSearch As adodb.Recordset
        Dim itmX As ListItem
        
        'invoiceout 217
        'virtualisp 78
        'acci_services 468
        'Radius. 130
        
        If MySQL.OpenTable(directConn, rsSearch, , "select distinct invoiceout.AccI_RecID, accountinfo.AccountName, virtualisp.Description, acci_services.Username, radiusaccounts.Username as radUser from radiusaccounts, invoiceout, accountinfo, acci_services, virtualisp " + _
                                                "where acci_services.RadiusID = radiusaccounts.RecID and accountinfo.RecID = invoiceout.acci_RecID and virtualisp.RecID = accountinfo.VirtualID and acci_services.RecID = invoiceout.PlanServiceID AND " + _
                                                "(TotalDue > (AmountPaid + AmountRefunded + GSTRefunded)) and PaymentDue <= '" + Format(sysnow, "yyyy-mm-dd Hh:Nn:Ss") + "' and accountinfo.FlagA_RecID = 1") = True Then
            If Not rsSearch.EOF And Not rsSearch.BOF Then
                pbUnpaid.Value = 0
                pbUnpaid.Max = rsSearch.RecordCount
                While Not rsSearch.EOF And Err.Number = 0
                                
                    Set itmX = lvUnpaid.ListItems.Add(, , IIf(IsNull(rsSearch!Description), "", rsSearch!Description))
                    itmX.SubItems(1) = IIf(IsNull(rsSearch!AccountName), "", rsSearch!AccountName)
                    itmX.SubItems(2) = IIf(IsNull(rsSearch!radUser), IIf(IsNull(rsSearch!Username), "", rsSearch!Username), rsSearch!radUser)
                                
                    MySQL.Execute directConn, "UPDATE accountinfo SET FlagA_RecID = 4 where RecID = " & rsSearch!acci_RecID
                    
                    If Not IsNull(rsSearch!radUser) Then
                        MySQL.Execute directConn, "update radius.RadiusUserGroup SET groupname = 'unpaid' where username = '" & rsSearch!radUser & "'"
                    End If
                    
                    pbUnpaid.Value = pbUnpaid.Value + 1
                    rsSearch.MoveNext
                Wend
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

'
'  This subroutine builds the Radius Flat File for users and prepares it for upload to the server
'  using a timer control
'
'  This routine needs to be changed to accommodate a Database driven radius server not a
'  Flat file driven one.
'
'
'
Public Sub cmdUploadUsers_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdUploadUsers_Click"
    Const ContainerName = "frmBot"
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

    
    On Error GoTo ErrorOccur
    
    tsBots.Tabs(2).Selected = True
    
    Dim rsload As adodb.Recordset
    Dim rsPools As adodb.Recordset
    Dim rsPlans As adodb.Recordset
    Dim rsRadiusAcc As adodb.Recordset
    Dim rsRadius As adodb.Recordset
    
    
    Dim bResult As Boolean
    
    Exit Sub
    
        pbRadUpdate.Value = 0
        pbRadUpdate.Max = 3
        
        'RadiusPools 11
        
        Dim bAOK As Boolean
        
        frameUsers.Tag = frameUsers.Caption
        txtStats.Text = ""
        frameUsers.Caption = frameUsers.Tag + " - Opening Tables"
        frameUsers.Refresh
        
        bResult = MySQL.OpenTable(directConn, rsPools, , "select RecID, Description from radiuspools")
        
        'bResult = MySQL.OpenTable(directConn, rsLoad, , "select count(*) as RecCount from radius.radiusradcheck")
        
        bResult = MySQL.OpenTable(directConn, rsRadiusAcc, , "select projectalpha.radiusaccounts.*, AES_DECRYPT(projectalpha.radiusaccounts.password,'" & odb.colSalts.ReturnSalt("md5Password") & "') as DecPassword from projectalpha.radiusaccounts, projectalpha.accountinfo, projectalpha.Flags, projectalpha.plantypes Where projectalpha.plantypes.RecID = projectalpha.radiusaccounts.ptRecID AND projectalpha.radiusaccounts.acci_RecID = projectalpha.accountinfo.RecID AND projectalpha.accountinfo.FlagA_RecID = Flags.RecID AND projectalpha.radiusaccounts.Checked = -1 AND projectalpha.Flags.ListedOnRadius = -1 and projectalpha.accountinfo.Cancelled = 0")
        If Login.bTestBench = True Then
            bResult = MySQL.OpenTable(directConn, rsRadius, , "select * from radiusradcheck")
        Else
            bResult = MySQL.OpenTable(directConn, rsRadius, , "select * from radiusradcheck")
        End If
        
        frameUsers.Caption = frameUsers.Tag + " - Security Pass 1"
        frameUsers.Refresh
        Dim iCount As Long
        Dim rsPoolBuf As adodb.Recordset
        
        If rsRadiusAcc.RecordCount > 0 Then
            pbRadUpdate.Max = pbRadUpdate.Max + rsRadiusAcc.RecordCount
            While Not rsRadiusAcc.EOF And Err.Number = 0
                bAOK = False
                If Val(rsRadiusAcc!AutoActivateFlag) <> 0 Then
                    If DateDiff("s", Format(rsRadiusAcc!Activation, "HH:MM:SS"), Format(sysnow, "HH:MM:SS")) > 0 And DateDiff("s", Format(rsRadiusAcc!Deactivation, "HH:MM:SS"), Format(sysnow, "HH:MM:SS")) < 0 Then bAOK = True
                Else
                    bAOK = True
                End If
                
                If Login.bTestBench = True Then
                    bResult = MySQL.OpenTable(directConn, rsRadius, , "select * from radiusradcheck where username = '" + rsRadiusAcc!Username + "'")
                Else
                    bResult = MySQL.OpenTable(directConn, rsRadius, , "select * from radiusradcheck where username = '" + rsRadiusAcc!Username + "'")
                End If
                If rsRadius.RecordCount = 0 Then
                    If bAOK = True Then
                        MySQL.Execute directConn, "Insert into radius.radiusradcheck (RadiusID, Username, Attribute, Value) VALUES (" & rsRadiusAcc!RecID & ",'" & rsRadiusAcc!Username & "','Crypt-Password',encrypt('" & rsRadiusAcc!DecPassword & "'))"
                        If MySQL.OpenTable(directConn, rsPoolBuf, , "select RadiusID from plantypes where RecID = " & rsRadiusAcc!ptRecID) = True Then
                            rsPools.Filter = "RecID = " & rsPoolBuf!RadiusID
                            If rsPools.RecordCount > 0 Then
                                MySQL.Execute directConn, "Insert into radius.RadiusUserGroup (Username, GroupName) VALUES ('" & rsRadiusAcc!Username & "','" & MySQL.rGroupName(rsPools!Description, rsRadiusAcc!RecID, directConn) & "')"
                            End If
                        End If
                        iCount = iCount + 1
                    End If
                Else
                    If bAOK = False Then
                        MySQL.Execute directConn, "Delete from radius.radiusradcheck Where Username = '" & rsRadiusAcc!Username & "'"
                        MySQL.Execute directConn, "Delete from radius.RadiusUserGroup Where Username = '" & rsRadiusAcc!Username & "'"
                        iCount = iCount + 1
                    End If
                End If
                
                pbRadUpdate.Value = pbRadUpdate.Value + 1
                gSleep
                gSleep
                gSleep
                gSleep
                
                rsRadiusAcc.MoveNext
                
            Wend
            txtStats.Text = txtStats.Text + "" & rsRadiusAcc.RecordCount & " - Users found with active radius permissions" & vbCrLf
            txtStats.Text = txtStats.Text + "" & iCount & " - Users added/altered on radius server" & vbCrLf
        End If
            
        iCount = 0
        
        frameUsers.Caption = frameUsers.Tag + " - Requerying table"
        frameUsers.Refresh
        
        If Login.bTestBench = True Then
            bResult = MySQL.OpenTable(directConn, rsRadius, , "select * from radiusradcheck")
        Else
            bResult = MySQL.OpenTable(directConn, rsRadius, , "select * from radiusradcheck")
        End If
        
        pbRadUpdate.Value = pbRadUpdate.Value + 1
        
        bResult = MySQL.OpenTable(directConn, rsRadiusAcc, , "select projectalpha.radiusaccounts.Username from projectalpha.radiusaccounts, projectalpha.accountinfo, projectalpha.Flags, projectalpha.plantypes Where projectalpha.plantypes.RecID = projectalpha.radiusaccounts.ptRecID AND projectalpha.radiusaccounts.acci_RecID = projectalpha.accountinfo.RecID AND projectalpha.accountinfo.FlagA_RecID = Flags.RecID AND (projectalpha.radiusaccounts.Checked = 0 OR projectalpha.Flags.ListedOnRadius = 0 or projectalpha.accountinfo.Cancelled = 1)")
        
        frameUsers.Caption = frameUsers.Tag + " - Security Pass 2"
        frameUsers.Refresh
        
        'RadiusRadcheck 14
        
        If rsRadiusAcc.RecordCount > 0 Then
            pbRadUpdate.Max = pbRadUpdate.Max + rsRadiusAcc.RecordCount
            While Not rsRadiusAcc.EOF And Err.Number = 0
                bResult = MySQL.OpenTable(directConn, rsRadius, , "select * from radiusradcheck where username = '" + rsRadiusAcc!Username + "'")
                If rsRadius.RecordCount > 0 Then
                    MySQL.Execute directConn, "Delete from radius.radiusradcheck Where Username = '" & rsRadiusAcc!Username & "'"
                    MySQL.Execute directConn, "Delete from radius.RadiusUserGroup Where Username = '" & rsRadiusAcc!Username & "'"
                    iCount = iCount + 1
                End If
                
                pbRadUpdate.Value = pbRadUpdate.Value + 1
                rsRadiusAcc.MoveNext
            
            Wend
            
            txtStats.Text = txtStats.Text + "" & rsRadiusAcc.RecordCount & " - Users found with inactive radius permissions" & vbCrLf
            txtStats.Text = txtStats.Text + "" & iCount & " - Users removed from the radius server for permissioning" & vbCrLf
        
        End If
        
        iCount = 0
        
        bResult = MySQL.OpenTable(directConn, rsRadius, , "select * from radiusradcheck")
        bResult = MySQL.OpenTable(directConn, rsRadiusAcc, , "select projectalpha.radiusaccounts.*, AES_DECRYPT(projectalpha.radiusaccounts.password,'" & odb.colSalts.ReturnSalt("md5Password") & "') as DecPassword from projectalpha.radiusaccounts, projectalpha.accountinfo, projectalpha.Flags, projectalpha.plantypes Where projectalpha.plantypes.RecID = projectalpha.radiusaccounts.ptRecID AND projectalpha.radiusaccounts.acci_RecID = projectalpha.accountinfo.RecID AND projectalpha.accountinfo.FlagA_RecID = Flags.RecID AND projectalpha.radiusaccounts.Checked = -1 AND projectalpha.Flags.ListedOnRadius = -1")
        
        'frameUsers.Caption = " Checking for Anomalies - Security Pass 3 "
        'frameUsers.Refresh
        
        
        'pbRadUpdate.Value = pbRadUpdate.Value + 1
        'If rsradius.RecordCount > 0 Then
        '    pbRadUpdate.Max = pbRadUpdate.Max + rsradius.RecordCount
        '    While Not rsradius.EOF
        '        rsRadiusAcc.Filter = "Username = '" & rsRadius!Username & "'"
        '        If rsRadiusAcc.RecordCount = 0 Then
        '            MySQL.Execute directConn, "Delete from radius.radiusradcheck Where ID = " & rsRadius!ID
        '            MySQL.Execute directConn, "Delete from radius.RadiusUserGroup Where Username = '" & rsRadius!Username & "'"
        '            iCount = iCount + 1
        '            frameUsers.Caption = "Warning Anomalies Found! [" & rsRadius!Username & "]"
        '            txtStats.text = txtStats.text + "::Warning Anomalies Found:: " & iCount & " - User [" & rsRadius!Username & "] was removed from the radius server not found in The Nexus Database" & vbCrLf
        '            frameUsers.Refresh
        '        End If
        '
        '        pbRadUpdate.Value = pbRadUpdate.Value + 1
        '        rsradius.MoveNext
        '
        '    Wend
        'End If
        'pbRadUpdate.Value = pbRadUpdate.Value + 1
        'tmrFile2.Enabled = True
        'Do: gSleep: Loop Until bUpload = False
        'rsPools.MoveNext
    'Wend
        frameUsers.Caption = frameUsers.Tag
    
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

Private Sub cmdViSP_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdViSP_Click"
    Const ContainerName = "frmBot"
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


    Dim br As Boolean
    Dim rsVISP As adodb.Recordset
    Dim ix As Long
    Dim eFlag As enum_FlagID
    
    Dim rsinvoiceout As adodb.Recordset
    Dim rsTemplates As adodb.Recordset
    Dim rsPlans As adodb.Recordset
    Dim rsServices As adodb.Recordset
    Dim rsacciServ As adodb.Recordset
    Dim rsSum As adodb.Recordset
    
        
    Dim cTotal As Currency
    Dim cMargin As Currency
    Dim cCost As Currency
    Dim cCharge As Currency
    
    Dim cVISPTotal As Currency
    Dim cVISPMargin As Currency
    Dim cVISPCost As Currency
    Dim cVISPCharge As Currency
    
    Const lBytesPerMB = 1024 ^ 2
    
    Dim lHeader As Long
    
    Screen.MousePointer = vbHourglass
    
    br = MySQL.OpenTable(directConn, rsVISP, , "select * from virtualisp Where NextCycle <= '" & Format(sysnow, "yyyy-mm-dd Hh:Nn:Ss") & "'")
    
    If rsVISP.RecordCount > 0 Then
        pbViSP.Value = 0
        pbViSP.Max = rsVISP.RecordCount * 2
        
        While Not rsVISP.EOF And Err.Number = 0
                                
            br = MySQL.OpenTable(directConn, rsinvoiceout, , "select * from invoiceout Where TotalDue = (AmountPaid + AmountRefunded + GSTRefunded) and FlagID = 0 and VirtualID = " & rsVISP!RecID)
            
            On Error Resume Next
            Do
                Err.Clear
                lHeader = MySQL.GetTMPRecID("visp_accountheader", directConn)
                MySQL.Execute directConn, "insert into visp_accountheader (RecID, VirtualID, StartDate, EndDate, TotalIncome, TotalCost, Margin) VALUES (" & lHeader & "," & rsVISP!RecID & ",'" & rsVISP!PreviousCycle & "','" & rsVISP!NextCycle & "',0,0,0)"
                If Err.Number <> 0 Then cDebug Err.Description
            Loop Until Err.Number = 0
           

            pbViSP.Value = pbViSP.Value + 1
            If rsinvoiceout.RecordCount > 0 Then
                pbViSP.Max = pbViSP.Max + rsinvoiceout.RecordCount
                While Not rsinvoiceout.EOF And Err.Number = 0
                
                    cCost = 0
                    cCharge = 0
                
                    If IsNull(rsinvoiceout!PlanServiceID) Then
                        
                    Else
                    
                        'PlanTemplates 16
                        'PlanTypes 102
                        'servicetypes 59
                        
                        br = MySQL.OpenTable(directConn, rsacciServ, , "select * from acci_services Where RecID = " & rsinvoiceout!PlanServiceID)
                        If rsTemplates Is Nothing Then br = MySQL.OpenTable(directConn, rsTemplates, "plantemplates")
                        If rsPlans Is Nothing Then br = MySQL.OpenTable(directConn, rsPlans, "plantypes")
                        If rsServices Is Nothing Then br = MySQL.OpenTable(directConn, rsServices, "servicetypes")
                                           
                        If rsacciServ.RecordCount > 0 Then
                            rsPlans.Filter = "RecID = " & rsacciServ!ptRecID
                            
                            If rsPlans.RecordCount > 0 Then
                                cCharge = cCharge + rsinvoiceout!AmountPaid
                                
                                If rsPlans!TemplateID <> 0 Then
                                    rsTemplates.Filter = "RecID = " & rsPlans!TemplateID
                                    cCost = cCost + rsTemplates!PeriodFee
                                Else
                                    cCost = cCost + rsPlans!PeriodFee
                                End If
                                
                                If rsPlans!MBPerPeriod <> -1 Then
                                    If rsinvoiceout!sfCycle_Download / lBytesPerMB > rsPlans!MBPerPeriod Then
                                        cCharge = cCharge + ((rsinvoiceout!sfCycle_Download / lBytesPerMB - rsPlans!MBPerPeriod) / rsPlans!MBBlockSize) * rsPlans!FeePerBlock
                                        If rsPlans!TemplateID <> 0 Then
                                            cCost = cCost + ((rsinvoiceout!sfCycle_Download / lBytesPerMB - rsTemplates!MBPerPeriod) / rsTemplates!MBBlockSize) * rsTemplates!FeePerBlock
                                        End If
                                    End If
                                End If
                                
                                If rsPlans!HoursPerPeriod <> -1 Then
                                    If rsinvoiceout!sfCycle_Mins / 60 > rsPlans!HoursPerPeriod Then
                                        cCharge = cCharge + ((rsinvoiceout!sfCycle_Mins / 60) - rsinvoiceout!HoursPerPeriod) * rsPlans!ExtraPerHour
                                        If rsPlans!TemplateID <> 0 Then
                                            cCost = cCost + ((rsinvoiceout!sfCycle_Mins / 60) - rsinvoiceout!HoursPerPeriod) * rsTemplates!ExtraPerHour
                                        End If
                                    End If
                                End If
                            End If
                         End If
                        
                    End If
                    eFlag = bProcessed0
                    rsinvoiceout!FlagID = eFlag
                    rsinvoiceout.Update
                    rsinvoiceout.MoveNext
                    MySQL.Execute directConn, "update visp_accountheader set TotalIncome=TotalIncome+" & cCharge & ", TotalCost=TotalCost+" & cCost & ",Margin=Margin+" & (cCharge - cCost) & " Where RecID = " & lHeader
                    pbViSP.Value = pbViSP.Value + 1
                    cTotal = cTotal + cCharge
                    cMargin = cMargin + cCost
                Wend
                
              End If
                'sysops 231
                'visp_accountitem 1
                
            br = MySQL.OpenTable(directConn, rsSum, , "select Sum(PerVISP) as PerVISP from sysops")
            
            MySQL.Execute directConn, "insert into visp_accountitem (AmountDue,CostPrice,GSTCharged,InvoiceID) VALUES (" & rsSum!PerVISP * (rsVISP!Subscribed / 100) & "," & rsSum!PerVISP * (rsVISP!Subscribed / 100) & "," & (rsSum!PerVISP * (rsVISP!Subscribed / 100)) * oTax(Login.TaxCode, Login.TaxCountry) & "," & lHeader & ",'Administration Fee for " & (rsVISP!Subscribed / 100) & " User Licence Blocks')"
            pbViSP.Value = pbViSP.Value + 1
            
            Call MySQL.SetNextCycle(rsVISP, "m", "1")
            rsVISP.Update
            rsVISP.MoveNext
        Wend

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

Public Sub Command1_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Command1_Click"
    Const ContainerName = "frmBot"
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


    frameEmails.Caption = "Searching and Sending Quota Messages"
    Call ProcessQuota
    frameEmails.Caption = "Searching and Sending Purchase Orders"
    Call ProcessPO
    
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
    Const ContainerName = "frmBot"
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

    If Login.bMaster = False Then
        Unload Me
        Exit Sub
    End If
    
    If bBigFont = True Then
        tsBots.Font.Size = 16
    End If
    
    Call tsBots_Click
    
    
 '   Set wrkODBC = CreateWorkspace("The Nexus_robots", UID, PWD, dbUseODBC)
    'Call MySQLConnection("Radius", sServer, sUID, sPWD, oADORadius, wrkODBC)
    
'    Call MySQL.Connection("Radius", sServer, sUID, sPWD, oADORadius)
    
    Call GUI.LoadColWidths(lvUnpaid, Me)
    
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

Private Sub Form_Paint()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Paint"
    Const ContainerName = "frmBot"
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


    iFrmState = True
    
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

Private Sub lblStats_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lblStats_Click"
    Const ContainerName = "frmBot"
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

Private Sub Form_Unload(Cancel As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Unload"
    Const ContainerName = "frmBot"
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


    Call GUI.SaveColWidths(lvUnpaid, Me)
    
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

Private Sub tmFile1_Timer()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "tmFile1_Timer"
    Const ContainerName = "frmBot"
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


    bDownload = True
    
    Static frmFTP As Form
    
    Dim iFilenameList As String
    Dim iFileSize As Variant
    
    Static attemptnum As Variant
    On Error Resume Next
    
    If frmFTP Is Nothing Then
        'Set frmFTP = New frmFTPMain2
        If frmFTP.iState = 0 Then
            attemptnum = attemptnum + 1
            If attemptnum > 3 Then
                If MsgBox("This is attempted number " & attemptnum & " to get in contact with the server '" & reg.sRadiusFTPServer & "' it is possibly timing out. Do you wish to continue with this action?", vbQuestion + vbYesNo, "Attempt Number " & attemptnum) = vbNo Then
                    attemptnum = 0
                    tmFile1.Enabled = False
                    bDownload = False
                    Exit Sub
                End If
            End If
            Load frmFTP
            'frmFTP.Show
            frameDownload.Caption = "Download In Progress - Connecting to " & reg.sRadiusFTPServer
            frmFTP.bActive = True
            frmFTP.bBinary = False
            frmFTP.Connect reg.sRadiusFTPServer, reg.sRadiusFTPPort, reg.sRadiusFTPUsername, reg.sRadiusFTPPassword, reg.sTargetDir
        End If
    Else
        If frmFTP.iState = 0 Then
            attemptnum = attemptnum + 1
            If attemptnum > 3 Then
                If MsgBox("This is attempted number " & attemptnum & " to get in contact with the server '" & reg.sRadiusFTPServer & "' it is possibly timing out. Do you wish to continue with this action?", vbQuestion + vbYesNo, "Attempt Number " & attemptnum) = vbNo Then
                    attemptnum = 0
                    tmFile1.Enabled = False
                    bDownload = False
                    Exit Sub
                End If
            End If
            Unload frmFTP
            Set frmFTP = Nothing
            'Set frmFTP = New frmFTPMain2
            Load frmFTP
            'frmFTP.Show
            frameDownload.Caption = "Download In Progress - Connecting to " & reg.sRadiusFTPServer
            frmFTP.bActive = True
            frmFTP.bBinary = False
            frmFTP.Connect reg.sRadiusFTPServer, reg.sRadiusFTPPort, reg.sRadiusFTPUsername, reg.sRadiusFTPPassword, reg.sTargetDir
        End If
    End If
    
    
    Select Case frmFTP.iState
    
    Case FTP_DIRECTORY_INFO_COMPLETED
        
        If frmFTP.iLoadState = True Then
        
            tmFile1.Enabled = False
            
            Dim NewFilename As String
            Dim oldFilename As String
            NewFilename = "Radius_" & Format(sysnow, "ddmmmyyhhnnss") & ".log"
            oldFilename = IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\") + "radius.bak"
            If Dir(oldFilename, vbNormal) <> "" Then Kill oldFilename
            Open oldFilename For Output As #10
                Print #10, "|File created " & Format(sysnow, "dd-mm-yyyy Hh:Nn:Ss") & "|"
            Close #10
            
            
            'Sets Form Variables
            frmFTP.bActive = True
            frmFTP.bBinary = True
            Set frmFTP.pb = pbRadUpdate
            Set frmFTP.LbTimeLeft = lblTimeLeft
            Set frmFTP.LbSpeed = lblSpeedKbps
            Set frmFTP.KBTransfered = lblKBTransfered
            frmFTP.Dir1.Path = App.Path
            
            Dim ix As Variant
            For ix = 1 To frmFTP.ListView2.ListItems.Count
                If frmFTP.ListView2.ListItems(ix).Text = reg.sTargetFilename Then
                    If frmFTP.ListView2.ListItems(ix).SubItems(1) = "  .04Kb" Then
                        tmFile1.Enabled = False
                        Unload frmFTP
                        frameDownload.Caption = "File only .04Kb - Header Only"
                        attemptnum = 0
                        'Set frmFTP = Nothing
                        Exit Sub
                    End If
                End If
            Next
            
            frameDownload.Caption = "Upload In Progress - radius.bak"
            Dim iG As Variant
            
            For iG = 1 To frmFTP.ListView1.ListItems.Count
                If frmFTP.ListView1.ListItems(iG).Text = "radius.bak" Then
                    frmFTP.ListView1.ListItems(iG).Selected = True
                    Exit For
                End If
            Next iG
            
            Call frmFTP.zLokUp_Click        ' Calls Upload routine
            Pause
           
            ' Renames file on server
            frameDownload.Caption = "Download In Progress - Renaming Files"
            'Call FtpRenameFile(server, reg.sTargetFilename, NewFilename)
            'Call FtpRenameFile(server, "radius.bak", reg.sTargetFilename)
            
            
            ' Checks and load the download directory
            If Dir(IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\") & "RadiusLog", vbDirectory) = "" Then MkDir App.Path & "\RadiusLog"
            frmFTP.Drive1.Drive = Left(App.Path, 2)
            frmFTP.Dir1.Path = IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\") & "RadiusLog"
            Call frmFTP.LoadLocal
            
            frmFTP.List
            'Do: gSleep: Loop While frmFTP.iLoadState <> FTP_DIRECTORY_INFO_COMPLETED
            
            For ix = 1 To frmFTP.ListView2.ListItems.Count
                If frmFTP.ListView2.ListItems(ix).Text = NewFilename Then
                    If frmFTP.ListView2.ListItems(ix).SubItems(1) = "  .04Kb" Then
                        tmFile1.Enabled = False
                        Unload frmFTP
                        Set frmFTP = Nothing
                        Exit Sub
                    End If
                    frmFTP.ListView2.ListItems(ix).Selected = True
                    Exit For
                End If
            Next
            frameDownload.Caption = "Downloading File - " & NewFilename
            'Sets Form Variables
            frmFTP.bActive = True
            frmFTP.bBinary = True
            Set frmFTP.pb = pbRadUpdate
            Set frmFTP.LbTimeLeft = lblTimeLeft
            Set frmFTP.LbSpeed = lblSpeedKbps
            Set frmFTP.KBTransfered = lblKBTransfered
            
            Call frmFTP.zRemDown_Click      ' Calls download routine
            
            Pause
            
            pbRadUpdate.Value = pbRadUpdate.Max
            
            Unload frmFTP
            Set frmFTP = Nothing
            
            TransmuteRadiusData IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\") & "RadiusLog\" & NewFilename
            
            attemptnum = 0
            tmFile1.Enabled = False
            bDownload = False
        Else
            attemptnum = 0
            Unload frmFTP
            Set frmFTP = Nothing
            tmFile1.Enabled = False
            bDownload = False
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

Public Function TransmuteRadiusData(sFilename As String)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "TransmuteRadiusData"
    Const ContainerName = "frmBot"
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


    'On Error Resume Next
    
    
    On Error GoTo 0
    
    Dim lFreefile As Variant
    Dim lFreefile2 As Variant
    
    Dim rsSave As adodb.Recordset
    Dim rsload As adodb.Recordset
    Dim rsRadiusLog As adodb.Recordset
    Dim rsaccountinfo As adodb.Recordset
    
    Dim lRecsCreated As Variant
    Dim lRecsAmmended As Variant
    
    Dim bResult As Boolean
    Dim tmpInput As String
    Dim tmpVariant As Variant
    
    Dim itmpString As String
    
    Dim iFieldName As String
    Dim iDBFieldName(255) As String
    
    Dim bytePos As Byte
    Dim byteType As Byte
    
    Dim pbyteUsername As Byte
    Dim pbyteSessionID As Byte
    Dim pbyteUpload As Byte
    Dim pbyteDownload As Byte
    Dim pbyteTime As Byte
    
    Dim iDBValue(255) As Variant
    Dim Value As String
    Dim dblUpload As Double
    Dim dblDownload As Double

    lFreefile = FreeFile
    
    bResult = MySQL.OpenTable(directConn, rsSave, , "select * from RadiusLogTwo Limit 0")
    
    Dim rsLoadb As adodb.Recordset
    Dim iPos As Variant
    Open sFilename For Input As #lFreefile
    Line Input #lFreefile, tmpInput
    
    If InStr(tmpInput, Chr$(10)) > 0 Then
        
        Close #lFreefile
        Name sFilename As Left(sFilename, Len(sFilename) - 3) + "err"
        Open sFilename For Output As #lFreefile
        
        pbTransmutate.Max = FileLen(Left(sFilename, Len(sFilename) - 3) + "err")
        pbTransmutate.Value = 0
        lFreefile2 = FreeFile
        Open Left(sFilename, Len(sFilename) - 3) + "err" For Input As #lFreefile2
        Do
            Line Input #lFreefile2, tmpInput
            iPos = 1
            tmpVariant = tmpInput
            While InStr(iPos + 3, tmpVariant, Chr$(10)) <> 0
                iPos = InStr(iPos + 3, tmpVariant, Chr$(10))
                If iPos = 1 Then
                    tmpVariant = vbCr & Mid(tmpVariant, Len(Chr$(10)))
                Else
                    tmpVariant = Left(tmpVariant, iPos - 1) & vbCr & Mid(tmpVariant, iPos + Len(Chr$(10)))
                
                End If
            Wend
            pbTransmutate.Value = pbTransmutate.Value + Len(tmpInput)
            gSleep
            Print #lFreefile, CStr(tmpVariant);
        Loop Until EOF(lFreefile2) Or Err.Number <> 0
    End If
    
    Close #lFreefile
    
    pbTransmutate.Max = FileLen(sFilename)
    Open sFilename For Input As #lFreefile
    pbTransmutate.Value = 0
    
    Do
        Line Input #lFreefile, tmpInput
        
        If InStr(tmpInput, ":") > 0 And Not Left(tmpInput, 1) = "|" Then
            
            If Not bytePos = 0 And byteType <> 3 Then
                
                bResult = MySQL.OpenTable(directConn, rsload, , "select RecID, acci_RecID, sfAliveTime, Acct_Session_ID, sfCycle_Upload, sfCycle_Download, sfStartTime, sfStopTime, sfCycle_Mins from radiusaccounts Where Username = '" & iDBValue(pbyteUsername) & "' Limit 1")
                    
                    bytePos = bytePos + 1
                    If rsload.RecordCount > 0 And byteType <> 3 Then
                        bResult = MySQL.OpenTable(directConn, rsaccountinfo, , "select * from accountinfo Where RecID = " & rsload!acci_RecID & " Limit 1")
                        rsload.MoveFirst
                        rsload!Acct_Session_ID = iDBValue(pbyteSessionID)
                        rsload.Update
                        iDBFieldName(bytePos) = "Acci_RecID"
                        iDBValue(bytePos) = rsload!RecID
                        If pbyteUpload <> 255 Then rsload!sfCycle_Upload = IIf(IsNull(rsload!sfCycle_Upload), 0, rsload!sfCycle_Upload) + iDBValue(pbyteUpload)
                        If pbyteDownlaod <> 255 Then rsload!sfCycle_Download = IIf(IsNull(rsload!sfCycle_Download), 0, rsload!sfCycle_Download) + iDBValue(pbyteDownload)
                        Select Case byteType
                        Case 1 ' Start
                            If DateDiff("s", rsload!sfStartTime, rsload!sfStopTime) < 0 Then
                                bResult = MySQL.OpenTable(directConn, rsLoadb, , "select * from RadiusLogTwo Where User_Name = '" & iDBValue(pbyteUsername) & "' AND Acct_Session_Id = '" & rsload!Acct_Session_ID & "' AND Acct_Status_Type = 3 Limit 1")
                                rsload!sfStopTime = rsLoadb!LogEntryCreated
                            End If
                            rsload!sfStartTime = iDBValue(pbyteTime)
                            rsaccountinfo!sfStartTime = rsload!sfStartTime
                            rsload!sfAliveTime = iDBValue(pbyteTime)
                        Case 2 ' Stop
                            rsload!sfStopTime = iDBValue(pbyteTime)
                            rsload!sfCycle_Mins = IIf(IsNull(rsload!sfCycle_Mins), 0, rsload!sfCycle_Mins) + DateDiff("n", rsload!sfAliveTime, rsload!sfStopTime)
                            rsaccountinfo!sfCycle_Mins = IIf(IsNull(rsaccountinfo!sfCycle_Mins), 0, rsaccountinfo!sfCycle_Mins) + DateDiff("n", rsload!sfAliveTime, rsload!sfStopTime)
                            
                            bResult = MySQL.OpenTable(directConn, rsRadiusLog, , "select * from RadiusLogTwo Where User_Name = '" & iDBValue(pbyteUsername) & "' AND Acct_Session_Id = '" & iDBValue(pbyteSessionID) & "' Limit 1")
                            If rsRadiusLog.RecordCount > 0 Then
                                For X = 0 To bytePos
                                    If X = pbyteUpload Then
                                        dblUpload = IIf(IsNull(rsRadiusLog(iDBFieldName(X))), 0, rsRadiusLog(iDBFieldName(X)))
                                    ElseIf X = pbyteDownload Then
                                        dblDownload = IIf(IsNull(rsRadiusLog(iDBFieldName(X))), 0, rsRadiusLog(iDBFieldName(X)))
                                    End If
                                Next X
                            End If
                            
                            If pbyteUpload <> 255 Then
                                rsload!sfCycle_Upload = IIf(IsNull(rsload!sfCycle_Upload), 0, rsload!sfCycle_Upload) + (iDBValue(pbyteUpload) - dblUpload)
                                rsaccountinfo!sfCycle_Upload = IIf(IsNull(rsaccountinfo!sfCycle_Upload), 0, rsaccountinfo!sfCycle_Upload) + (iDBValue(pbyteUpload) - dblUpload)
                            End If
                            If pbyteDownlaod <> 255 Then
                                rsload!sfCycle_Download = IIf(IsNull(rsload!sfCycle_Download), 0, rsload!sfCycle_Download) + (iDBValue(pbyteDownload) - dblDownload)
                                rsaccountinfo!sfCycle_Download = IIf(IsNull(rsaccountinfo!sfCycle_Download), 0, rsaccountinfo!sfCycle_Download) + (iDBValue(pbyteDownload) - dblDownload)
                            End If
                        End Select
                        rsaccountinfo.Update
                        rsload.Update
                    Else
                        iDBFieldName(bytePos) = "Acci_RecID"
                        iDBValue(bytePos) = 0
                    End If
                    
                    rsSave.AddNew
                    lRecsCreated = lRecsCreated + 1
                    For X = 0 To bytePos
                        rsSave(iDBFieldName(X)) = iDBValue(X)
                    Next X
                    rsSave.Update
                    bytePos = 0
                    pbyteUpload = 255
                    pbyteDownload = 255
                
            ElseIf Not bytePos = 0 And byteType = 3 Then
                               
                dblUpload = 0
                dblDownload = 0
                bResult = MySQL.OpenTable(directConn, rsload, , "select * from RadiusLogTwo Where User_Name = '" & iDBValue(pbyteUsername) & "' AND Acct_Session_Id = '" & iDBValue(pbyteSessionID) & "' Limit 1")
                
                If rsload.RecordCount > 0 Then
                    lRecsAmmended = lRecsAmmended + 1
                    For X = 0 To bytePos
                        If X = pbyteUpload Then
                            dblUpload = IIf(IsNull(rsload(iDBFieldName(X))), 0, rsload(iDBFieldName(X)))
                        ElseIf X = pbyteDownload Then
                            dblDownload = IIf(IsNull(rsload(iDBFieldName(X))), 0, rsload(iDBFieldName(X)))
                        End If
                        rsload(iDBFieldName(X)) = iDBValue(X)
                    Next X
                    rsload.Update
                Else
                
                    rsSave.AddNew
                    lRecsCreated = lRecsCreated + 1
                    For X = 0 To bytePos
                        rsSave(iDBFieldName(X)) = iDBValue(X)
                    Next X
                    rsSave.Update
                End If
                
                bResult = MySQL.OpenTable(directConn, rsload, , "select RecID, acci_RecID, Acct_Session_ID, sfAliveTime, sfStopTime, sfCycle_Mins, sfCycle_Upload, sfCycle_Download from radiusaccounts Where Username = '" & iDBValue(pbyteUsername) & "' Limit 1")
                bytePos = bytePos + 1
                If rsload.RecordCount > 0 Then
                    rsload.MoveFirst
                    bResult = MySQL.OpenTable(directConn, rsaccountinfo, , "select * from accountinfo Where RecID = " & rsload!acci_RecID & " Limit 1")
                    iDBFieldName(bytePos) = "Acci_RecID"
                    iDBValue(bytePos) = rsload!RecID

                    If pbyteUpload <> 255 Then
                        rsload!sfCycle_Upload = IIf(IsNull(rsload!sfCycle_Upload), 0, rsload!sfCycle_Upload) + (iDBValue(pbyteUpload) - dblUpload)
                        rsaccountinfo!sfCycle_Upload = IIf(IsNull(rsaccountinfo!sfCycle_Upload), 0, rsaccountinfo!sfCycle_Upload) + (iDBValue(pbyteUpload) - dblUpload)
                    End If
                    If pbyteDownlaod <> 255 Then
                        rsload!sfCycle_Download = IIf(IsNull(rsload!sfCycle_Download), 0, rsload!sfCycle_Download) + (iDBValue(pbyteDownload) - dblDownload)
                        rsaccountinfo!sfCycle_Download = IIf(IsNull(rsaccountinfo!sfCycle_Download), 0, rsaccountinfo!sfCycle_Download) + (iDBValue(pbyteDownload) - dblDownload)
                    End If
                    
                    'bResult = MySQL.OpenTable(directConn, rsLoadb, , "select * from RadiusLogTwo Where User_Name = '" & iDBValue(pbyteUsername) & "' AND Acct_Session_Id = '" & rsLoad!Acct_Session_ID & "' AND Acct_Status_Type = 3 Limit 1")
                    rsload!sfStopTime = iDBValue(pbyteTime)
                    rsload!sfCycle_Mins = IIf(IsNull(rsload!sfCycle_Mins), 0, rsload!sfCycle_Mins) + DateDiff("n", rsload!sfAliveTime, iDBValue(pbyteTime))
                    
                    rsaccountinfo!sfCycle_Mins = IIf(IsNull(rsaccountinfo!sfCycle_Mins), 0, rsaccountinfo!sfCycle_Mins) + DateDiff("n", rsload!sfAliveTime, iDBValue(pbyteTime))
                    rsload!sfAliveTime = iDBValue(pbyteTime)
                    rsaccountinfo.Update
                    rsload.Update
                Else
                    iDBFieldName(bytePos) = "Acci_RecID"
                    iDBValue(bytePos) = 0
                End If
                bytePos = 0
                bytePos = 0
                pbyteUpload = 255
                pbyteDownload = 255
            End If
            
            'If rsSave.EOF And rsSave.BOF Then
            'ElseIf rsSave.EditMode = adEditAdd Then rsSave.Update
            'End If
            '
            
            'rsSave!LogEntryCreated = CDate(Trim(Mid(tmpInput, 9, 2)) + "-" + Trim(Mid(tmpInput, 5, 3)) + "-" + Mid(tmpInput, Len(tmpInput) - 3, 4) + " " + Left(Right(tmpInput, 13), 8))
            iDBValue(bytePos) = CDate(Trim(Mid(tmpInput, 9, 2)) + "-" + Trim(Mid(tmpInput, 5, 3)) + "-" + Mid(tmpInput, Len(tmpInput) - 3, 4) + " " + Left(Right(tmpInput, 13), 8))
            iDBFieldName(bytePos) = "LogEntryCreated"
            pbyteTime = bytePos
            
        ElseIf tmpInput = "" Then
        ElseIf Left(tmpInput, 1) = "|" Then
        Else
            bytePos = bytePos + 1
            iFieldName = Trim(Left(tmpInput, InStr(tmpInput, "=") - 1))
            Value = Trim(Mid(tmpInput, InStr(tmpInput, "=") + 1))
            
            iDBFieldName(bytePos) = ""
            For X = 1 To Len(iFieldName)
                iDBFieldName(bytePos) = iDBFieldName(bytePos) + IIf(Mid(iFieldName, X, 1) = "-", "_", Mid(iFieldName, X, 1))
            Next
            
            iDBValue(bytePos) = ""
            For X = 1 To Len(Value)
                iDBValue(bytePos) = iDBValue(bytePos) + IIf(Mid(Value, X, 1) = Chr$(34), "", Mid(Value, X, 1))
            Next
            
            If iDBFieldName(bytePos) = "Acct_Status_Type" Then
                Select Case iDBValue(bytePos)
                Case "Start"
                    iDBValue(bytePos) = 1
                Case "Stop"
                    iDBValue(bytePos) = 2
                Case "Alive"
                    iDBValue(bytePos) = 3
                End Select
                byteType = iDBValue(bytePos)
            ElseIf iDBFieldName(bytePos) = "Acct_Session_Id" Then
                

                pbyteSessionID = bytePos
                
            ElseIf iDBFieldName(bytePos) = "User_Name" Then

                pbyteUsername = bytePos
                            
            ElseIf InStr(iDBFieldName(bytePos), "Octet") > 0 Then
                iDBValue(bytePos) = Oct(iDBValue(bytePos))
                If iDBFieldName(bytePos) = "Acct_Output_Octets" Then
                    pbyteUpload = bytePos
                ElseIf iDBFieldName(bytePos) = "Acct_Input_Octets" Then
                    pbyteDownload = bytePos
                End If
            End If
            
            
        End If
        
        If pbTransmutate.Value + Len(tmpInput) + IIf(EOF(lFreefile), 0, 2) < pbTransmutate.Max Then pbTransmutate.Value = pbTransmutate.Value + Len(tmpInput) + IIf(EOF(lFreefile), 0, 2)
        gSleep
        
    Loop Until EOF(lFreefile) Or Err.Number <> 0
    
    Close #lFreefile
    
    If Not bytePos = 0 And byteType <> 3 Then
        
        bResult = MySQL.OpenTable(directConn, rsload, , "select RecID, acci_RecID, sfAliveTime, Acct_Session_ID, sfCycle_Upload, sfCycle_Download, sfStartTime, sfStopTime, sfCycle_Mins from radiusaccounts Where Username = '" & iDBValue(pbyteUsername) & "' Limit 1")
            
            bytePos = bytePos + 1
            If rsload.RecordCount > 0 And byteType <> 3 Then
                bResult = MySQL.OpenTable(directConn, rsaccountinfo, , "select * from accountinfo Where RecID = " & rsload!acci_RecID & " Limit 1")
                rsload.MoveFirst
                rsload!Acct_Session_ID = iDBValue(pbyteSessionID)
                rsload.Update
                iDBFieldName(bytePos) = "Acci_RecID"
                iDBValue(bytePos) = rsload!RecID
                If pbyteUpload <> 255 Then rsload!sfCycle_Upload = IIf(IsNull(rsload!sfCycle_Upload), 0, rsload!sfCycle_Upload) + iDBValue(pbyteUpload)
                If pbyteDownlaod <> 255 Then rsload!sfCycle_Download = IIf(IsNull(rsload!sfCycle_Download), 0, rsload!sfCycle_Download) + iDBValue(pbyteDownload)
                Select Case byteType
                Case 1 ' Start
                    If DateDiff("s", rsload!sfStartTime, rsload!sfStopTime) < 0 Then
                        bResult = MySQL.OpenTable(directConn, rsLoadb, , "select * from RadiusLogTwo Where User_Name = '" & iDBValue(pbyteUsername) & "' AND Acct_Session_Id = '" & rsload!Acct_Session_ID & "' AND Acct_Status_Type = 3 Limit 1")
                        rsload!sfStopTime = rsLoadb!LogEntryCreated
                    End If
                    rsload!sfStartTime = iDBValue(pbyteTime)
                    rsaccountinfo!sfStartTime = rsload!sfStartTime
                    rsload!sfAliveTime = iDBValue(pbyteTime)
                Case 2 ' Stop
                    rsload!sfStopTime = iDBValue(pbyteTime)
                    rsload!sfCycle_Mins = IIf(IsNull(rsload!sfCycle_Mins), 0, rsload!sfCycle_Mins) + DateDiff("n", rsload!sfAliveTime, rsload!sfStopTime)
                    rsaccountinfo!sfCycle_Mins = IIf(IsNull(rsaccountinfo!sfCycle_Mins), 0, rsaccountinfo!sfCycle_Mins) + DateDiff("n", rsload!sfAliveTime, rsload!sfStopTime)
                    
                    bResult = MySQL.OpenTable(directConn, rsRadiusLog, , "select * from RadiusLogTwo Where User_Name = '" & iDBValue(pbyteUsername) & "' AND Acct_Session_Id = '" & iDBValue(pbyteSessionID) & "' Limit 1")
                    If rsRadiusLog.RecordCount > 0 Then
                        For X = 0 To bytePos
                            If X = pbyteUpload Then
                                dblUpload = IIf(IsNull(rsRadiusLog(iDBFieldName(X))), 0, rsRadiusLog(iDBFieldName(X)))
                            ElseIf X = pbyteDownload Then
                                dblDownload = IIf(IsNull(rsRadiusLog(iDBFieldName(X))), 0, rsRadiusLog(iDBFieldName(X)))
                            End If
                        Next X
                    End If
                    
                    If pbyteUpload <> 255 Then
                        rsload!sfCycle_Upload = IIf(IsNull(rsload!sfCycle_Upload), 0, rsload!sfCycle_Upload) + (iDBValue(pbyteUpload) - dblUpload)
                        rsaccountinfo!sfCycle_Upload = IIf(IsNull(rsaccountinfo!sfCycle_Upload), 0, rsaccountinfo!sfCycle_Upload) + (iDBValue(pbyteUpload) - dblUpload)
                    End If
                    If pbyteDownlaod <> 255 Then
                        rsload!sfCycle_Download = IIf(IsNull(rsload!sfCycle_Download), 0, rsload!sfCycle_Download) + (iDBValue(pbyteDownload) - dblDownload)
                        rsaccountinfo!sfCycle_Download = IIf(IsNull(rsaccountinfo!sfCycle_Download), 0, rsaccountinfo!sfCycle_Download) + (iDBValue(pbyteDownload) - dblDownload)
                    End If
                End Select
                rsaccountinfo.Update
                rsload.Update
            Else
                iDBFieldName(bytePos) = "Acci_RecID"
                iDBValue(bytePos) = 0
            End If
            
            rsSave.AddNew
            lRecsCreated = lRecsCreated + 1
            For X = 0 To bytePos
                rsSave(iDBFieldName(X)) = iDBValue(X)
            Next X
            rsSave.Update
            bytePos = 0
            pbyteUpload = 255
            pbyteDownload = 255
        
    ElseIf Not bytePos = 0 And byteType = 3 Then
                       
        dblUpload = 0
        dblDownload = 0
        bResult = MySQL.OpenTable(directConn, rsload, , "select * from RadiusLogTwo Where User_Name = '" & iDBValue(pbyteUsername) & "' AND Acct_Session_Id = '" & iDBValue(pbyteSessionID) & "' Limit 1")
        
        If rsload.RecordCount > 0 Then
            lRecsAmmended = lRecsAmmended + 1
            For X = 0 To bytePos
                If X = pbyteUpload Then
                    dblUpload = IIf(IsNull(rsload(iDBFieldName(X))), 0, rsload(iDBFieldName(X)))
                ElseIf X = pbyteDownload Then
                    dblDownload = IIf(IsNull(rsload(iDBFieldName(X))), 0, rsload(iDBFieldName(X)))
                End If
                rsload(iDBFieldName(X)) = iDBValue(X)
            Next X
            rsload.Update
        Else
        
            rsSave.AddNew
            lRecsCreated = lRecsCreated + 1
            For X = 0 To bytePos
                rsSave(iDBFieldName(X)) = iDBValue(X)
            Next X
            rsSave.Update
        End If
        
        bResult = MySQL.OpenTable(directConn, rsload, , "select RecID, acci_RecID, Acct_Session_ID, sfAliveTime, sfStopTime, sfCycle_Mins, sfCycle_Upload, sfCycle_Download from radiusaccounts Where Username = '" & iDBValue(pbyteUsername) & "' Limit 1")
        bytePos = bytePos + 1
        If rsload.RecordCount > 0 Then
            rsload.MoveFirst
            bResult = MySQL.OpenTable(directConn, rsaccountinfo, , "select * from accountinfo Where RecID = " & rsload!acci_RecID & " Limit 1")
            iDBFieldName(bytePos) = "Acci_RecID"
            iDBValue(bytePos) = rsload!RecID

            If pbyteUpload <> 255 Then
                rsload!sfCycle_Upload = IIf(IsNull(rsload!sfCycle_Upload), 0, rsload!sfCycle_Upload) + (iDBValue(pbyteUpload) - dblUpload)
                rsaccountinfo!sfCycle_Upload = IIf(IsNull(rsaccountinfo!sfCycle_Upload), 0, rsaccountinfo!sfCycle_Upload) + (iDBValue(pbyteUpload) - dblUpload)
            End If
            If pbyteDownlaod <> 255 Then
                rsload!sfCycle_Download = IIf(IsNull(rsload!sfCycle_Download), 0, rsload!sfCycle_Download) + (iDBValue(pbyteDownload) - dblDownload)
                rsaccountinfo!sfCycle_Download = IIf(IsNull(rsaccountinfo!sfCycle_Download), 0, rsaccountinfo!sfCycle_Download) + (iDBValue(pbyteDownload) - dblDownload)
            End If
            
            'bResult = MySQL.OpenTable(directConn, rsLoadb, , "select * from RadiusLogTwo Where User_Name = '" & iDBValue(pbyteUsername) & "' AND Acct_Session_Id = '" & rsLoad!Acct_Session_ID & "' AND Acct_Status_Type = 3 Limit 1")
            rsload!sfStopTime = iDBValue(pbyteTime)
            rsload!sfCycle_Mins = IIf(IsNull(rsload!sfCycle_Mins), 0, rsload!sfCycle_Mins) + DateDiff("n", rsload!sfAliveTime, iDBValue(pbyteTime))
            
            rsaccountinfo!sfCycle_Mins = IIf(IsNull(rsaccountinfo!sfCycle_Mins), 0, rsaccountinfo!sfCycle_Mins) + DateDiff("n", rsload!sfAliveTime, iDBValue(pbyteTime))
            rsload!sfAliveTime = iDBValue(pbyteTime)
            rsaccountinfo.Update
            rsload.Update
        Else
            iDBFieldName(bytePos) = "Acci_RecID"
            iDBValue(bytePos) = 0
        End If
        bytePos = 0
        bytePos = 0
        pbyteUpload = 255
        pbyteDownload = 255
    End If
    
    pbTransmutate.Value = pbTransmutate.Max
    
    lblStats.Caption = "Records Created : " & lRecsCreated & vbCrLf & "Records Ammended: " & lRecsAmmended
    
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

Private Sub tmProgress_Timer()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "tmProgress_Timer"
    Const ContainerName = "frmBot"
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

    If Login.bMaster = False Then
        Unload Me
        Exit Sub
    End If
        
    If pbTransmutate.Value < pbTransmutate.Max Then
        If Not frmMDIMain.lvSchedule.ListItems(cmdRadiusUpdate.Tag).SubItems(3) = "" & Round((pbTransmutate.Value / pbTransmutate.Max * 100)) & "%" Then frmMDIMain.lvSchedule.ListItems(cmdRadiusUpdate.Tag).SubItems(3) = "" & Round((pbTransmutate.Value / pbTransmutate.Max * 100)) & "%"
    Else
        If Not frmMDIMain.lvSchedule.ListItems(cmdRadiusUpdate.Tag).SubItems(3) = "100%" Then frmMDIMain.lvSchedule.ListItems(cmdRadiusUpdate.Tag).SubItems(3) = "100%"
    End If
    
    If cmdUploadUsers.Tag <> "" Then
        If pbRadUpdate.Value < pbRadUpdate.Max Then
            If Not frmMDIMain.lvSchedule.ListItems(cmdUploadUsers.Tag).SubItems(3) = "" & Round((pbRadUpdate.Value / pbRadUpdate.Max * 100)) & "%" Then frmMDIMain.lvSchedule.ListItems(cmdUploadUsers.Tag).SubItems(3) = "" & Round((pbRadUpdate.Value / pbRadUpdate.Max * 100)) & "%"
        ElseIf pbRadUpdate.Value = pbRadUpdate.Max Then
            If Not frmMDIMain.lvSchedule.ListItems(cmdUploadUsers.Tag).SubItems(3) = "100%" Then frmMDIMain.lvSchedule.ListItems(cmdUploadUsers.Tag).SubItems(3) = "100%"
        End If
    End If
    
    If pbBilling.Value < pbBilling.Max Then
        If Not frmMDIMain.lvSchedule.ListItems(cmdProcessBills.Tag).SubItems(3) = "" & ((pbBilling.Value / pbBilling.Max * 100)) & "%" Then frmMDIMain.lvSchedule.ListItems(cmdProcessBills.Tag).SubItems(3) = "" & ((pbBilling.Value / pbBilling.Max * 100)) & "%"
    ElseIf pbBilling.Value = pbBilling.Max Then
        If Not frmMDIMain.lvSchedule.ListItems(cmdProcessBills.Tag).SubItems(3) = "100%" Then frmMDIMain.lvSchedule.ListItems(cmdProcessBills.Tag).SubItems(3) = "100%"
    End If

    If pbQuota.Value < pbQuota.Max Then
        If Not frmMDIMain.lvSchedule.ListItems(Command1.Tag).SubItems(3) = "" & ((pbQuota.Value / pbQuota.Max * 100)) & "%" Then frmMDIMain.lvSchedule.ListItems(Command1.Tag).SubItems(3) = "" & ((pbQuota.Value / pbQuota.Max * 100)) & "%"
    ElseIf pbQuota.Value = pbQuota.Max Then
        If Not frmMDIMain.lvSchedule.ListItems(Command1.Tag).SubItems(3) = "100%" Then frmMDIMain.lvSchedule.ListItems(Command1.Tag).SubItems(3) = "100%"
    End If
    
    If pbUnpaid.Value < pbUnpaid.Max Then
        If Not frmMDIMain.lvSchedule.ListItems(cmdUnpaid.Tag).SubItems(3) = "" & ((pbUnpaid.Value / pbUnpaid.Max * 100)) & "%" Then frmMDIMain.lvSchedule.ListItems(cmdUnpaid.Tag).SubItems(3) = "" & ((pbUnpaid.Value / pbUnpaid.Max * 100)) & "%"
    ElseIf pbUnpaid.Value = pbUnpaid.Max Then
        If Not frmMDIMain.lvSchedule.ListItems(cmdUnpaid.Tag).SubItems(3) = "100%" Then frmMDIMain.lvSchedule.ListItems(cmdUnpaid.Tag).SubItems(3) = "100%"
    End If
    
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

Private Sub tmrFile2_Timer()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "tmrFile2_Timer"
    Const ContainerName = "frmBot"
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


    Static frmFTP As Form
    
    Dim iFilenameList As String
    
    Static attemptnum As Variant
    
    bUpload = True
    If frmFTP Is Nothing Then
        'Set frmFTP = New frmFTPMain2
        If frmFTP.iState = 0 Then
            attemptnum = attemptnum + 1
            
            If attemptnum > 1 Then
                If MsgBox("This is attempted number " & attemptnum & " to get in contact with the server '" & reg.sRadiusFTPServer & "' it is possibly timing out. Do you wish to continue with this action?", vbQuestion + vbYesNo, "Attempt Number " & attemptnum) = vbNo Then
                    attemptnum = 0
                    tmFile1.Enabled = False
                    Kill tmrFile2.Tag
                    bUpload = False
                    Exit Sub
                End If
            End If
                                   
            Load frmFTP
            frameUpload.Caption = "Download In Progress - Connecting to " & reg.s2RadiusFTPServer
            frmFTP.Connect reg.s2RadiusFTPServer, reg.s2RadiusFTPPort, reg.s2RadiusFTPUsername, reg.s2RadiusFTPPassword, reg.s2TargetDir
        End If
    Else
        If frmFTP.iState = 0 Then
            
            attemptnum = attemptnum + 1
            
            If attemptnum > 1 Then
                If MsgBox("This is attempted number " & attemptnum & " to get in contact with the server '" & reg.sRadiusFTPServer & "' it is possibly timing out. Do you wish to continue with this action?", vbQuestion + vbYesNo, "Attempt Number " & attemptnum) = vbNo Then
                    attemptnum = 0
                    tmrFile2.Enabled = False
                    Kill tmrFile2.Tag
                    bUpload = False
                    Exit Sub
                End If
            End If
            
            Unload frmFTP
            Set frmFTP = Nothing
            'Set frmFTP = New frmFTPMain2
            Load frmFTP
            'frmFTP.Show
            frameUpload.Caption = "Download In Progress - Connecting to " & reg.s2RadiusFTPServer
            frmFTP.Connect reg.s2RadiusFTPServer, reg.s2RadiusFTPPort, reg.s2RadiusFTPUsername, reg.s2RadiusFTPPassword, reg.s2TargetDir
        End If
    End If
    
    
    
    Select Case frmFTP.iState
    
    Case FTP_DIRECTORY_INFO_COMPLETED
        
        gSleep
                
        'Sets Form Variables
        frmFTP.bActive = True
        frmFTP.bBinary = True
        Set frmFTP.pb = pbUploadUsers
        Set frmFTP.LbTimeLeft = lblTimeUpload
        Set frmFTP.LbSpeed = lblSpeedUpload
        Set frmFTP.KBTransfered = lblKBUpload
        frmFTP.Dir1.Path = App.Path
        
        Dim ix As Variant
        For ix = 1 To frmFTP.ListView1.ListItems.Count
            If frmFTP.ListView1.ListItems(ix).Text = reg.s2TargetFilename Then
                frmFTP.ListView1.ListItems(ix).Selected = True
                Exit For
            End If
        Next
        
        frameUpload.Caption = "Upload In Progress - " & reg.s2TargetFilename
        
        Call frmFTP.zLokUp_Click        ' Calls Upload routine
        Pause
        Unload frmFTP
        Set frmFTP = Nothing
        
        'Dim ix as Variant
        'For ix = 1 To frmFTP.ListView2.ListItems.Count
        '    If frmFTP.ListView2.ListItems(ix).Text = newFilename Then
        '        frmFTP.ListView2.ListItems(ix).selected = True
        '    End If
        'Next
        
        'Dim Item2 As ListItem
        'Set Item2 = frmFTP.ListView3.ListItems.Add(, , reg.s2TargetFilename)
        
        'Item2.SubItems(1) = tmrFile2.Tag
        'If iDoeventLev >= 2 Then gSleep
        'Item2.SubItems(2) = frmFTP.mFTP.GetFTPDirectory
        'If iDoeventLev >= 2 Then gSleep
        'Item2.SubItems(3) = frmFTP.TxtConnectedTo.Text
        'If iDoeventLev >= 2 Then gSleep
        'Item2.SubItems(4) = FileLen(tmrFile2.Tag)
        'ffrmFTP.TxtTotalBytesQueued.Text = TxtTotalBytesQueued.Text + iFileSize
        'ffrmFTP.Text1.Text = Text1.Text + iFileSize
        'If iDoeventLev >= 2 Then gSleep
        'Item2.SubItems(5) = "Upload"
        
        'frmFTP.bActive = True
        'frmFTP.bBinary = False
        
        'Set frmFTP.PB = pbUploadUsers
        'Set frmFTP.Label3 = lblTimeUpload
        'Set frmFTP.Label8 = lblSpeedUpload
        'Set frmFTP.Label11 = lblKBUpload
        
        'Call frmFTP.menudownload_Click
        'frameUpload.Caption = "Upload In Progress - " & reg.s2TargetFilename
        'Call frmFTP.menutransfer_Click
                            
        'Unload frmFTP
        'Set frmFTP = Nothing
        
        attemptnum = 0
        
        tmrFile2.Enabled = False
        
        Kill tmrFile2.Tag
        bUpload = False
        
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

Private Sub tsBots_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "tsBots_Click"
    Const ContainerName = "frmBot"
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
    
    For X = picBots.LBound To picBots.UBound
        If tsBots.SelectedItem.Index - 1 <> X Then picBots(X).Visible = False
    Next
    
    If tsBots.SelectedItem.Index - 1 <= picBots.UBound Then
        picBots(tsBots.SelectedItem.Index - 1).Move tsBots.ClientLeft, tsBots.ClientTop, tsBots.ClientWidth, tsBots.ClientHeight
        picBots(tsBots.SelectedItem.Index - 1).Visible = True
        picBots(tsBots.SelectedItem.Index - 1).ZOrder 0
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

Public Function ProcessCycle()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "ProcessCycle"
    Const ContainerName = "frmBot"
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
    
    Const SendMail = True
    
    Const lBytesPerMB = 1024 ^ 2
    
    Dim rsload As adodb.Recordset
    Dim rsload2 As adodb.Recordset
    Dim rsRadius As adodb.Recordset
    Dim rsSave As adodb.Recordset
    Dim rsPlanType As adodb.Recordset
    Dim rsAccInfo As adodb.Recordset
    Dim oConn As adodb.Connection
    
    Dim invout As invoiceout
    
    Dim StatementID As Single
    
    Dim lInvSubRecID As Long
    Dim cPrePaid As Currency
    Dim cTotalDue As Currency
    Dim cAmountPaid As Currency
    Dim bNewGroup As Boolean
    Dim bResult As Boolean
    Dim cCharge As Currency
    Dim cChargeTotal As Currency
    Dim lSubRecID As Long
    Dim rsRefer As adodb.Recordset
    
    Dim AddReceipt As Boolean
    
    Dim Att As Long
    
    pbBilling.Value = 0
    pbBilling.Max = 1
    
    
    Select Case Login.bTestBench
    Case False
        If MySQL.Connection(, sServer, , , oConn) = False Then
            MsgBox "Unable to Connect to MySQL Server, Please check your internet connection and attempt to restart the program.", vbCritical, "MySQL Server Not Found"
            End
        End If
    Case True
        
        sServer = "localhost"
        sUID = "pa2004"
        sPWD = "p0st41"
        If MySQL.Connection(, sServer, sUID, sPWD, oConn) = False Then
            MsgBox "Unable to Connect to MySQL Test Bench Server, Please check your LAN connection and attempt to restart the program.", vbCritical, "MySQL Test Bench Not Found"
            End
        End If
    
    End Select

    
    bResult = MySQL.OpenTable(oConn, rsload, , "select RecID from acci_services where Checked = 0 and Activation = NextCycle and Activation <= '" & Format(sysnow, "YYYY-MM-DD Hh:Nn:SS") & "' order by SubRecID")
    If rsload.RecordCount > 0 Then
        pbBilling.Max = rsload.RecordCount
        While Not rsload.EOF And Err.Number = 0
            MySQL.Execute oConn, "Update acci_services Set Checked=-1 where RecID = " & rsload!RecID
            rsload.MoveNext
            pbBilling.Value = pbBilling.Value + 1
            gSleep
        Wend
    End If
    
    bResult = MySQL.OpenTable(oConn, rsload, , "select RecID, SubRecID, ptRecID, acci_RecID, RadiusID, SysopID, NextCycle, PreviousCycle, Activation, JoiningFee, PeriodFee, PerMB, PerHour from acci_services where Checked = -1 and NextCycle <= '" & Format(sysnow, "YYYY-MM-DD Hh:Nn:SS") & "' order by SubRecID", adOpenStatic, adLockReadOnly)
    bResult = MySQL.OpenTable(oConn, rsPlanType, , "select * from plantypes")
    
    Dim rsObject As adodb.Recordset
    
    If rsload.RecordCount > 0 Then
        pbBilling.Max = pbBilling.Max + rsload.RecordCount
        While Not rsload.EOF And Err.Number = 0
            
            'If rsLoad!ptRecID = 1319583009 Then Stop
            If lSubRecID <> rsload!SubRecID Then
                bNewGroup = True
                lSubRecID = rsload!SubRecID
            End If

            
            If rsload!RadiusID <> 0 Then
            
                bResult = MySQL.OpenTable(oConn, rsRadius, , "select RecID, sfCycle_Upload , sfCycle_Download, sfCycle_Mins, VirtualID from radiusaccounts Where RecID = " & rsload!RadiusID)
            
                
                cCharge = 0
                rsPlanType.Filter = "RecID = " & rsload!ptRecID
                If rsPlanType.RecordCount > 0 Then
                    
                 
                    If MySQL.OpenTable(oConn, rsObject, , "select AccountActive from acci_dslconnections where acci_RecID = " & rsload!acci_RecID) = True Then
                        If rsObject.RecordCount > 0 Then
                            Select Case Val(rsObject!AccountActive)
                            Case 0
                                GoTo SkipBillCycle
                            Case Else
                            
                            
                            End Select
                        End If
                    End If
                    
                    If rsPlanType!BillOnce = 1 And rsload!Activation = rsload!NextCycle Then
                    
                    ElseIf rsPlanType!BillOnce = 1 And rsload!Activation <> rsload!NextCycle Then
                        GoTo SkipBillCycle
                    End If
                    
                    'bResult = MySQL.OpenTable(oConn, rsSave, , "select * from invoiceout Limit 1")
                    'bResult = MySQL.OpenTable(oConn, rsLoad2, , "select * from accountinfo Where RecID = " & rsLoad!acci_RecID)
                    
                    'rsSave.AddNew
                    On Error GoTo 0
                    invout.Description = rsPlanType!Description
                    
                    If rsload!Activation = rsload!NextCycle Then
                        
                        ' IF the account has been set to a latter Account Start Up Date/Activation Date
                        ' This section of the routine will search for whether it is a referal Item and
                        ' Apply the nessary Charges
                        
                        If MySQL.OpenTable(oConn, rsRefer, , "select RecID from acci_referedby where acciServiceID = " & rsload!RecID) = True Then
                            If rsRefer.RecordCount > 0 Then
                                invout.Description = invout.Description + " - Account Setup Free"
                            Else
                                invout.Description = invout.Description + " - Account Setup Fee [" & Format(Val(rsload!JoiningFee), "Currency") & "]"
                                cCharge = cCharge + Val(rsload!JoiningFee)
                            End If
                        Else
                            invout.Description = invout.Description + " - Account Setup Fee [" & Format(Val(rsload!JoiningFee), "Currency") & "]"
                            cCharge = cCharge + Val(rsload!JoiningFee)
                        End If
                    
                        MySQL.Execute oConn, "Update acci_services set Checked=-1 where RecID = " & rsload!RecID
                    End If
                    
                    If rsPlanType!MBPerPeriod <> -1 Then
                        If rsRadius!sfCycle_Download / lBytesPerMB > rsPlanType!MBPerPeriod Then
                            invout.Description = invout.Description + " " & (rsRadius!sfCycle_Download / lBytesPerMB) - rsPlanType!MBPerPeriod & "MB's Over"
                            
                            cCharge = cCharge + ((rsRadius!sfCycle_Download / 1048576 - rsPlanType!MBPerPeriod) / rsPlanType!MBBlockSize) * rsload!PerMB
                        End If
                    End If
                    
                    invout.sfCycle_Download = IIf(IsNull(rsRadius!sfCycle_Download), 0, rsRadius!sfCycle_Download)
                    invout.sfCycle_Upload = IIf(IsNull(rsRadius!sfCycle_Upload), 0, rsRadius!sfCycle_Upload)
                    
                    bResult = MySQL.OpenTable(oConn, rsAccInfo, , "select RecID, sfCycle_Upload, sfCycle_Download, sfCycle_Mins from accountinfo Where RecID = " & rsload!acci_RecID)
                    
                    If rsAccInfo.RecordCount > 0 Then
                        MySQL.Execute oConn, "Insert Into history_acci_datausage (acci_RecID, Uploaded, Downloaded, NumMins) VALUES(" & IIf(IsNull(rsAccInfo!RecID), 0, rsAccInfo!RecID) & ", " & IIf(IsNull(rsAccInfo!sfCycle_Upload), 0, rsAccInfo!sfCycle_Upload) & ", " & IIf(IsNull(rsAccInfo!sfCycle_Download), 0, rsAccInfo!sfCycle_Download) & ", " & IIf(IsNull(rsAccInfo!sfCycle_Mins), 0, rsAccInfo!sfCycle_Mins) & ")"
                        MySQL.Execute oConn, "Update accountinfo Set sfCycle_Upload =sfCycle_Upload - " & rsRadius!sfCycle_Upload & " where RecID = " & rsload!acci_RecID
                        MySQL.Execute oConn, "Update accountinfo Set sfCycle_Download =sfCycle_Download - " & rsRadius!sfCycle_Download & " where RecID = " & rsload!acci_RecID
                        MySQL.Execute oConn, "Update accountinfo Set sfCycle_Mins =sfCycle_Mins - " & rsRadius!sfCycle_Mins & " where RecID = " & rsload!acci_RecID
                    End If
                    
                    MySQL.Execute oConn, "UPDATE radiusaccounts Set sfCycle_Download = 0 where RecID = " & rsRadius!RecID
                    MySQL.Execute oConn, "UPDATE radiusaccounts Set sfCycle_Upload = 0 where RecID = " & rsRadius!RecID
                    
                    
                    If rsPlanType!HoursPerPeriod <> -1 Then
                        If rsRadius!sfCycle_Mins / 60 > rsPlanType!HoursPerPeriod Then
                            invout.Description = invout.Description + " " & (rsRadius!sfCycle_Mins - rsPlanType!HoursPerPeriod * 60) & "Min's Over"
                            cCharge = cCharge + (rsRadius!sfCycle_Mins / 60) * rsload!PerHour
                        End If
                    End If
                    
                    invout.sfCycle_Mins = rsRadius!sfCycle_Mins
                    MySQL.Execute oConn, "Update accountinfo set sfCycle_Download=sfCycle_Download-" & IIf(IsNull(invout.sfCycle_Download), 0, invout.sfCycle_Download) & ", sfCycle_Upload = sfCycle_Upload - " & IIf(IsNull(invout.sfCycle_Upload), 0, invout.sfCycle_Upload) & ", sfCycle_Mins = sfCycle_Mins - " & IIf(IsNull(invout.sfCycle_Mins), 0, invout.sfCycle_Mins) & " where RecID = " & rsload!acci_RecID
                    
                    MySQL.Execute oConn, "UPDATE radiusaccounts Set sfCycle_Mins = 0 where RecID = " & rsRadius!RecID
                    
                    cCharge = cCharge + IIf(rsload!PeriodFee = 0, rsPlanType!PeriodFee, rsload!PeriodFee)

                    MySQL.Execute oConn, "update acci_services set PreviousCycle = NextCycle where RecID = " & rsload!RecID
                    MySQL.Execute oConn, "update acci_services set NextCycle = '" & Format(DateAdd(IIf(IsNull(rsPlanType!chgIntervalType), "m", rsPlanType!chgIntervalType), IIf(IsNull(rsPlanType!chgInterval), 1, rsPlanType!chgInterval), sysnow), "yyyy-mm-dd ttttt") & "' where RecID = " & rsload!RecID

                    invout.StartCycle = rsload!NextCycle
                    invout.EndCycle = DateAdd(IIf(IsNull(rsPlanType!chgIntervalType), "m", rsPlanType!chgIntervalType), IIf(IsNull(rsPlanType!chgInterval), 1, rsPlanType!chgInterval), sysnow)
                    
                    'rsload.Update
                                 
                    ' Calculates any prepaid disposition with currency.
                    bResult = MySQL.OpenTable(oConn, rsload2, , "select * from invoicein Where AmountPaid > AmountUsed AND AccI_RecID = " & rsload!acci_RecID)
                    cPrePaid = 0
                    cTotalDue = 0
                    cAmountPaid = 0
                    AddReceipt = False
                    cTotalDue = cCharge + cCharge * oTax(Login.TaxCode, Login.TaxCountry)
                    If rsload2.RecordCount > 0 Then
                        Do While Not rsload2.EOF And Err.Number = 0
                            If rsload2!AmountPaid - rsload2!AmountUsed > 0 Then
                                cPrePaid = rsload2!AmountPaid - rsload2!AmountUsed
                                If cPrePaid < cTotalDue Then
                                    cTotalDue = cTotalDue - cPrePaid
                                    cAmountPaid = cAmountPaid + cPrePaid
                                    rsload2!AmountUsed = rsload2!AmountPaid
                                    rsload2.Update
                                    invout.PaidWhen = sysnow
                                ElseIf cPrePaid > cTotalDue Then
                                    rsload2!AmountUsed = rsload2!AmountUsed + cTotalDue
                                    cAmountPaid = cAmountPaid + cTotalDue
                                    cTotalDue = 0
                                    rsload2.Update
                                    invout.PaidWhen = sysnow
                                End If
                            End If
                            rsload2.MoveNext
                            If cTotalDue = 0 Then Exit Do
                            AddReceipt = True
                        Loop
                    End If
                    
                    ' Saves the entry for the account Invoice going out
                    invout.AmountDue = cCharge
                    invout.GSTCharged = cCharge * oTax(Login.TaxCode, Login.TaxCountry)
                    invout.TotalDue = cTotalDue
                    invout.AmountRefunded = 0
                    invout.GSTRefunded = 0
                    invout.AmountPaid = cAmountPaid
                    invout.acci_RecID = rsload!acci_RecID
                    invout.SysopID = rsload!SysopID
                    invout.PlanServiceID = rsload!RecID

                    bResult = MySQL.OpenTable(oConn, rsload2, , "select RecID, VirtualID, payIntervalType, payInterval from accountinfo Where RecID = " & rsload!acci_RecID)
                    If rsload2.RecordCount > 0 Then
                        invout.PaymentDue = DateAdd(IIf(IsNull(rsload2!PayIntervalType), "d", rsload2!PayIntervalType), IIf(IsNull(rsload2!PayInterval), 14, rsload2!PayInterval), sysnow)
                        invout.VirtualID = IIf(IsNull(rsload2!VirtualID), Login.lVirtualID, rsload2!VirtualID)
                    Else
                        invout.PaymentDue = DateAdd("d", 14, sysnow)
                        invout.VirtualID = Login.lVirtualID
                    End If
                    
                    

                    If bNewGroup = True Then
                        On Error Resume Next
                        Do
                            Err.Clear
                            lInvSubRecID = MySQL.GetTMPRecID("invoiceout", oConn, "RecID", False)
                            invout.RecID = lInvSubRecID
                            invout.SubRecID = lInvSubRecID
                            Call MySQL.Execute(oConn, "INSERT INTO invoiceout (RecID, Description, PaidWhen, AmountDue, GSTCharged, " + _
                                                    "TotalDue, AmountRefunded, GSTRefunded, AmountPaid, acci_RecID, SysopID, PlanServiceID, SubRecID, StartCycle, EndCycle, PaymentDue, VirtualID, StatementID, sfCycle_Download, sfCycle_Upload, sfCycle_Mins) " + _
                                                    "VALUES ('" & invout.RecID & "','" & MySQL.ESC(invout.Description) & "','" & Format(invout.PaidWhen, "YYYY-MM-DD ttttt") & "','" & invout.AmountDue & "','" & invout.GSTCharged & _
                                                    "','" & invout.TotalDue & "','" & invout.AmountRefunded & "','" & invout.GSTRefunded & "','" & invout.AmountPaid & "','" & invout.acci_RecID & "','" & invout.SysopID & _
                                                    "','" & invout.PlanServiceID & "','" & invout.SubRecID & "','" & Format(invout.StartCycle, "YYYY-MM-DD ttttt") & "','" & Format(invout.EndCycle, "YYYY-MM-DD ttttt") & "','" & Format(invout.PaymentDue, "YYYY-MM-DD ttttt") & "','" & invout.VirtualID & "','" & invout.StatementID & "','" & invout.sfCycle_Download & "','" & invout.sfCycle_Upload & "','" & invout.sfCycle_Mins & "')")
                    
                            If Err.Number <> 0 Then cDebug Err.Description
                        Loop Until Err.Number = 0
                        
                        Do
                            Err.Clear
                            StatementID = MySQL.GetTMPRecID("statementitems", oConn, "RecID", False)
                            MySQL.Execute oConn, "insert into statementitems (RecID, InvRecID, Items, Description, TotalDue) Values (" & StatementID & "," & lInvSubRecID & ",1,'" & sSTR.ReplaceString(invout.Description, "'", "\'") & "'," & cTotalDue & ")"
                            If Err.Number <> 0 Then cDebug Err.Description
                        Loop Until Err.Number = 0
                        If bDebug = True Then On Error GoTo 0 Else On Error GoTo ErrorOccur
                        MySQL.Execute oConn, "update invoiceout Set StatementID = " & StatementID & " where RecID = " & lInvSubRecID
                        
                        bNewGroup = False
                        If bDebug = True Then On Error GoTo 0 Else On Error GoTo ErrorOccur
                        
                    Else
                        On Error Resume Next
                            Do
                                Err.Clear
                                lInvSubRecID = MySQL.GetTMPRecID("invoiceout", oConn, "RecID", False)
                                invout.RecID = lInvSubRecID
                                invout.StatementID = StatementID
                                Call MySQL.Execute(oConn, "INSERT INTO invoiceout (RecID, Description, PaidWhen, AmountDue, GSTCharged, " + _
                                                        "TotalDue, AmountRefunded, GSTRefunded, AmountPaid, acci_RecID, SysopID, PlanServiceID, SubRecID, StartCycle, EndCycle, PaymentDue, VirtualID, StatementID, sfCycle_Download, sfCycle_Upload, sfCycle_Mins) " + _
                                                        "VALUES ('" & invout.RecID & "','" & MySQL.ESC(invout.Description) & "','" & Format(invout.PaidWhen, "YYYY-MM-DD ttttt") & "','" & invout.AmountDue & "','" & invout.GSTCharged & _
                                                        "','" & invout.TotalDue & "','" & invout.AmountRefunded & "','" & invout.GSTRefunded & "','" & invout.AmountPaid & "','" & invout.acci_RecID & "','" & invout.SysopID & _
                                                        "','" & invout.PlanServiceID & "','" & invout.SubRecID & "','" & Format(invout.StartCycle, "YYYY-MM-DD ttttt") & "','" & Format(invout.EndCycle, "YYYY-MM-DD ttttt") & "','" & Format(invout.PaymentDue, "YYYY-MM-DD ttttt") & "','" & invout.VirtualID & "','" & invout.StatementID & "','" & invout.sfCycle_Download & "','" & invout.sfCycle_Upload & "','" & invout.sfCycle_Mins & "')")
                                If Err.Number <> 0 Then cDebug Err.Description
                            Loop Until Err.Number = 0

                        MySQL.Execute oConn, "Update statementitems set Items=Items+1,TotalDue = TotalDue +" & cTotalDue & " where InvRecID = " & lInvSubRecID
                    End If
                    
                    If AddReceipt = True Then Call MySQL.AddReceiptItem(oConn, rsload!acci_RecID, invout.RecID, , rsload!RecID, , cAmountPaid, , "Account Payment")
                    
                Else
                
                End If
SkipBillCycle:
                cChargeTotal = cChargeTotal + cCharge
            Else
                cCharge = 0
                rsPlanType.Filter = "RecID = " & IIf(IsNull(rsload!ptRecID), 0, rsload!ptRecID)
                
                If rsPlanType.RecordCount > 0 Then
                            
                    If MySQL.OpenTable(oConn, rsObject, , "select AccountActive from acci_dslconnections where acci_RecID = " & rsload!acci_RecID) = True Then
                        If rsObject.RecordCount > 0 Then
                            Select Case Val(rsObject!AccountActive)
                            Case 0
                                GoTo SkipBillCycle2
                            Case Else
                            
                            
                            End Select
                        End If
                    End If
                    
                    If rsPlanType!BillOnce = 1 And rsload!Activation = rsload!NextCycle Then
                    
                    ElseIf rsPlanType!BillOnce = 1 And rsload!Activation <> rsload!NextCycle Then
                        GoTo SkipBillCycle2
                    End If
                    
                    If Val(rsPlanType!BillOnce) = False Then
                        'bResult = MySQL.OpenTable(oConn, rsSave, , "select * from invoiceout Limit 1")
                        'rsSave.AddNew
                        
                        invout.Description = rsPlanType!Description
                        
                        If rsload!Activation = rsload!NextCycle Then
                            
                            ' IF the account has been set to a latter Account Start Up Date/Activation Date
                            ' This section of the routine will search for whether it is a referal Item and
                            ' Apply the nessary Charges
                            
                            If MySQL.OpenTable(oConn, rsRefer, , "select RecID from acci_referedby where acciServiceID = " & rsload!RecID) = True Then
                                If rsRefer.RecordCount > 0 Then
                                    invout.Description = invout.Description + " - Account Setup Free"
                                Else
                                    invout.Description = invout.Description + " - Account Setup Fee [" & Format(rsload!JoiningFee, "Currency") & "]"
                                    cCharge = cCharge + Val(rsload!JoiningFee)
                                End If
                            Else
                                invout.Description = invout.Description + " - Account Setup Fee [" & Format(rsload!JoiningFee, "Currency") & "]"
                                cCharge = cCharge + rsload!JoiningFee
                            End If
                        
                            MySQL.Execute oConn, "Update acci_services set Checked=-1 where RecID = " & rsload!RecID
                        End If
                        
                        cCharge = cCharge + rsload!PeriodFee
                        
                        MySQL.Execute oConn, "update acci_services set PreviousCycle = NextCycle where RecID = " & rsload!RecID
                        MySQL.Execute oConn, "update acci_services set NextCycle = '" & Format(DateAdd(IIf(IsNull(rsPlanType!chgIntervalType), "m", rsPlanType!chgIntervalType), IIf(IsNull(rsPlanType!chgInterval), 1, rsPlanType!chgInterval), sysnow), "yyyy-mm-dd ttttt") & "' where RecID = " & rsload!RecID
                 
                        ' Calculates any prepaid disposition with currency.
                        bResult = MySQL.OpenTable(oConn, rsload2, , "select * from invoicein Where AmountPaid > AmountUsed AND AccI_RecID = " & rsload!acci_RecID)
                        cPrePaid = 0
                        cTotalDue = 0
                        cAmountPaid = 0
                        cTotalDue = cCharge + cCharge * oTax(Login.TaxCode, Login.TaxCountry)
                        AddReceipt = False
                        If rsload2.RecordCount > 0 Then
                            Do While Not rsload2.EOF And Err.Number = 0
                                If rsload2!AmountPaid - rsload2!AmountUsed > 0 Then
                                    cPrePaid = rsload2!AmountPaid - rsload2!AmountUsed
                                    If cPrePaid < cTotalDue Then
                                        cTotalDue = cTotalDue - cPrePaid
                                        cAmountPaid = cAmountPaid + cPrePaid
                                        'rsLoad2!AmountUsed = rsLoad2!AmountPaid
                                        Call MySQL.Execute(oConn, "Update invoicein AmountUsed = AmountPaid where RecID = " & rsload2!RecID)
                                        rsload2.Resync adAffectCurrent
                                        invout.PaidWhen = sysnow
                                    ElseIf cPrePaid > cTotalDue Then
                                        'rsLoad2!AmountUsed = rsLoad2!AmountUsed + cTotalDue
                                        Call MySQL.Execute(oConn, "Update invoicein AmountUsed = AmountUsed + " & cTotalDue & " where RecID = " & rsload2!RecID)
                                        cAmountPaid = cAmountPaid + cTotalDue
                                        cTotalDue = 0
                                        rsload2.Resync adAffectCurrent
                                        invout.PaidWhen = sysnow
                                    End If
                                End If
                                rsload2.MoveNext
                                If cTotalDue = 0 Then Exit Do
                            Loop
                            AddReceipt = True
                        End If
                        
                        ' Saves the entry for the account Invoice going out
                        invout.AmountDue = cCharge
                        invout.GSTCharged = cCharge * oTax(Login.TaxCode, Login.TaxCountry)
                        invout.TotalDue = cTotalDue
                        invout.PlanServiceID = rsload!RecID
                        invout.AmountPaid = cAmountPaid
                        invout.acci_RecID = rsload!acci_RecID
                        
                        bResult = MySQL.OpenTable(oConn, rsload2, , "select * from accountinfo Where RecID = " & rsload!acci_RecID)
                        If rsload2.RecordCount > 0 Then
                            invout.PaymentDue = DateAdd(IIf(IsNull(rsload2!PayIntervalType), "d", rsload2!PayIntervalType), IIf(IsNull(rsload2!PayInterval), 14, rsload2!PayInterval), sysnow)
                            invout.VirtualID = IIf(IsNull(rsload2!VirtualID), Login.lVirtualID, rsload2!VirtualID)
                        Else
                            invout.PaymentDue = DateAdd("d", 14, sysnow)
                            invout.VirtualID = Login.lVirtualID
                        End If
                        
                        
                        
                        If bNewGroup = True Then
                            On Error Resume Next
                            Do
                                Err.Clear
                                lInvSubRecID = MySQL.GetTMPRecID("invoiceout", oConn, "RecID", False)
                                invout.RecID = lInvSubRecID
                                invout.SubRecID = lInvSubRecID
                                Call MySQL.Execute(oConn, "INSERT INTO invoiceout (RecID, Description, PaidWhen, AmountDue, GSTCharged, " + _
                                                        "TotalDue, AmountRefunded, GSTRefunded, AmountPaid, acci_RecID, SysopID, PlanServiceID, SubRecID, StartCycle, EndCycle, PaymentDue, VirtualID, StatementID, sfCycle_Download, sfCycle_Upload, sfCycle_Mins) " + _
                                                        "VALUES ('" & invout.RecID & "','" & MySQL.ESC(invout.Description) & "','" & Format(invout.PaidWhen, "YYYY-MM-DD ttttt") & "','" & invout.AmountDue & "','" & invout.GSTCharged & _
                                                        "','" & invout.TotalDue & "','" & invout.AmountRefunded & "','" & invout.GSTRefunded & "','" & invout.AmountPaid & "','" & invout.acci_RecID & "','" & invout.SysopID & _
                                                        "','" & invout.PlanServiceID & "','" & invout.SubRecID & "','" & Format(invout.StartCycle, "YYYY-MM-DD ttttt") & "','" & Format(invout.EndCycle, "YYYY-MM-DD ttttt") & "','" & Format(invout.PaymentDue, "YYYY-MM-DD ttttt") & "','" & invout.VirtualID & "','" & invout.StatementID & "','" & invout.sfCycle_Download & "','" & invout.sfCycle_Upload & "','" & invout.sfCycle_Mins & "')")
                                If Err.Number <> 0 Then cDebug Err.Description
                            Loop Until Err.Number = 0
                            If bDebug = True Then On Error GoTo 0 Else On Error GoTo ErrorOccur
                            
                            Do
                                Err.Clear
                                StatementID = MySQL.GetTMPRecID("statementitems", oConn, "RecID", False)
                                MySQL.Execute oConn, "insert into statementitems (RecID, InvRecID, Items, Description, TotalDue) Values (" & StatementID & "," & lInvSubRecID & ",1,'" & sSTR.ReplaceString(invout.Description, "'", "\'") & "'," & cTotalDue & ")"
                                If Err.Number <> 0 Then cDebug Err.Description
                            Loop Until Err.Number = 0
                            If bDebug = True Then On Error GoTo 0 Else On Error GoTo ErrorOccur
                            MySQL.Execute oConn, "update invoiceout Set StatementID = " & StatementID & " where RecID = " & lInvSubRecID
                            
                            bNewGroup = False
                        Else
                            On Error Resume Next
                            Do
                                Err.Clear
                                lInvSubRecID = MySQL.GetTMPRecID("invoiceout", oConn, "RecID", False)
                                invout.RecID = lInvSubRecID
                                Call MySQL.Execute(oConn, "INSERT INTO invoiceout (RecID, Description, PaidWhen, AmountDue, GSTCharged, " + _
                                                        "TotalDue, AmountRefunded, GSTRefunded, AmountPaid, acci_RecID, SysopID, PlanServiceID, SubRecID, StartCycle, EndCycle, PaymentDue, VirtualID, StatementID, sfCycle_Download, sfCycle_Upload, sfCycle_Mins) " + _
                                                        "VALUES ('" & invout.RecID & "','" & MySQL.ESC(invout.Description) & "','" & Format(invout.PaidWhen, "YYYY-MM-DD ttttt") & "','" & invout.AmountDue & "','" & invout.GSTCharged & _
                                                        "','" & invout.TotalDue & "','" & invout.AmountRefunded & "','" & invout.GSTRefunded & "','" & invout.AmountPaid & "','" & invout.acci_RecID & "','" & invout.SysopID & _
                                                        "','" & invout.PlanServiceID & "','" & invout.SubRecID & "','" & Format(invout.StartCycle, "YYYY-MM-DD ttttt") & "','" & Format(invout.EndCycle, "YYYY-MM-DD ttttt") & "','" & Format(invout.PaymentDue, "YYYY-MM-DD ttttt") & "','" & invout.VirtualID & "','" & invout.StatementID & "','" & invout.sfCycle_Download & "','" & invout.sfCycle_Upload & "','" & invout.sfCycle_Mins & "')")
                                If Err.Number <> 0 Then cDebug Err.Description
                            Loop Until Err.Number = 0
                            
                            MySQL.Execute oConn, "Update statementitems set Items=Items+1,TotalDue = TotalDue +" & cTotalDue & " where InvRecID = " & lInvSubRecID
                        End If
                    Else
                    
                    End If
                    
                    If AddReceipt = True Then Call MySQL.AddReceiptItem(oConn, rsload!acci_RecID, invout.RecID, , rsload!RecID, , cAmountPaid, , "", "Account Payment")
                    
                End If
SkipBillCycle2:
                cChargeTotal = cChargeTotal + cCharge
                
            End If
            rsPlanType.Filter = ""
            rsload.MoveNext
            pbBilling.Value = pbBilling.Value + 1
            gSleep
        Wend
        
    Else
        pbBilling.Max = 1
        pbBilling.Value = 1
    End If
    
    Dim HTML As Variant
    Dim rsTraxr As adodb.Recordset
    Dim rseMail As adodb.Recordset
    Dim iTraxrID As Variant
     Dim AddFilename As String
    
    
    bResult = MySQL.OpenTable(oConn, rsload, , "select distinct AccI_RecID, accountinfo.payInterval, accountinfo.VirtualID, accountinfo.payIntervalType  from invoiceout, accountinfo where accountinfo.RecID = invoiceout.AccI_RecID and accountinfo.Cancelled = 0 and InvoiceFlagID = 0")
    If rsload.RecordCount > 0 Then
        pbBilling.Max = pbBilling.Max + rsload.RecordCount
        While Not rsload.EOF And Err.Number = 0
                
            InvCreated = InvCreated + 1
            'bResult = MySQL.OpenTable(oConn, rsTraxr, , "select * from invoicetraxr Limit 0")
            'rsTraxr.AddNew
            
            'InvoiceTraxr 41
            
            On Error Resume Next
            Do
                Err.Clear
                iTraxrID = MySQL.SetInvoiceSerial(oConn)
                MySQL.Execute oConn, "Insert Into invoicetraxr (RecID, InvoiceSerial, acci_RecID, PaymentDue) VALUES ('" & iTraxrID & "','" & Hex(iTraxrID) & "','" & rsload!acci_RecID & "','" & Format(DateAdd(IIf(IsNull(rsload!PayIntervalType), "d", rsload!PayIntervalType), IIf(IsNull(rsload!PayInterval), 14, rsload!PayInterval), sysnow), "yyyy-mm-dd Hh:Nn:Ss") & "')"

                If Err.Number <> 0 Then cDebug Err.Description

            Loop Until Err.Number = 0
            
            If bDebug = True Then On Error GoTo 0 Else On Error GoTo ErrorOccur
            
            HTML = cSMTP.sendInvoiceHTML(oConn, rsload!acci_RecID, iTraxrID)
            
            SMTP1.MailDate = Format(sysnow, "ddddd ttttt")
            SMTP1.MailFrom = Chr$(34) + "Exitstencil Press Australia Billing Services" & Chr$(34) + " <jcosti@ep.net.au>"
            SMTP1.MessageSubject = "Invoice #" & Hex(iTraxrID) & " for Internet Services"
            HTML = MySQL.ReplaceString(HTML, "/InvoiceNumber/", "" & iTraxrID & "")
            
            cDebug "Sending " & SMTP1.MessageSubject
            SMTP1.SendTo = ""
            AddFilename = ""
            If MySQL.OpenTable(oConn, rseMail, , "select AES_DECRYPT(EmailAddress,'" & odb.colSalts.ReturnSalt(EMAILSalt) & "') as Emailaddress from acci_emailaddresses where Checked = -1 and AccI_RecID = " & rsload!acci_RecID) = True Then
                If rseMail.RecordCount > 0 Then
                 While Not rseMail.EOF And Err.Number = 0
                    If IsNull(rseMail!EmailAddress) Then AddFilename = "*"
                    SMTP1.SendTo = SMTP1.SendTo + IIf(IsNull(rseMail!EmailAddress), "jcosti@ep.net.au", rseMail!EmailAddress) + ", "
                    rseMail.MoveNext
                 Wend
                Else
                    SMTP1.SendTo = "jcosti@ep.net.au,"
                End If
            Else
                SMTP1.SendTo = "jcosti@ep.net.au,"
            End If
            SMTP1.SendTo = Left(SMTP1.SendTo, Len(SMTP1.SendTo) - 1) & ", " & oResell.ReturnCatch(IIf(IsNull(rsload!VirtualID), Login.lVirtualID, rsload!VirtualID), "invoice")
            If Login.bTestBench = True Or SMTP1.SendTo = "" Then SMTP1.SendTo = "jcosti@ep.net.au, " & oResell.ReturnCatch(IIf(IsNull(rsload!VirtualID), Login.lVirtualID, rsload!VirtualID), "invoice")
            SMTP1.SendTo = MySQL.ReplaceString(SMTP1.SendTo, ",,", ",")
            SMTP1.Server = reg.smtpServer
            SMTP1.Port = reg.smtpPort
            SMTP1.Username = reg.smtpUsername
            SMTP1.Password = reg.smtpPassword
            
            If IsNull(HTML) Then GoTo SkipSend
            
            SMTP1.MessageHTML = MySQL.ReplaceString(HTML, "/EmailAddy/", SMTP1.SendTo)
            
            SMTP1.Attachments.Add cSMTP.SaveHTML(SMTP1.MessageHTML, AddFilename & SMTP1.MessageSubject + ".html", sysnow)
            If Login.bTestBench = False Then
                SMTP1.SendEmail
                
                For Att = SMTP1.Attachments.Count To 1 Step -1
                    SMTP1.Attachments.Remove Att
                Next
                If Not Frame5(0).Caption = SMTP1.Status + " - " & SMTP1.MessageSubject Then Frame5(0).Caption = SMTP1.Status + " - " & SMTP1.MessageSubject
                Do
                    'SMTP1.SendEmail
                    gSleep
                    If Not Frame5(0).Caption = SMTP1.Status + " - " & SMTP1.MessageSubject Then Frame5(0).Caption = SMTP1.Status + " - " & SMTP1.MessageSubject
                Loop Until SMTP1.Status = "SMTP session closed" Or SMTP1.Status = "" Or Err.Number <> 0
            End If
            
SkipSend:

            rsload.MoveNext
            pbBilling.Value = pbBilling.Value + 1
            gSleep
        Wend
    End If
    
    Dim iReceiptNo As Double
    
    bResult = MySQL.OpenTable(oConn, rsload, , "select distinct accountinfo.VirtualID, receipts.acci_RecID from receipts inner join accountinfo on receipts.acci_RecID = accountinfo.RecID where ReceiptNo = 0")
    If rsload.RecordCount > 0 Then
        pbBilling.Max = pbBilling.Max + rsload.RecordCount
        While Not rsload.EOF And Err.Number = 0
            
            HTML = cSMTP.sendReceiptHTML(oConn, rsload!acci_RecID, iReceiptNo, False)
            
            SMTP1.MailDate = Format(sysnow, "ddddd ttttt")
            SMTP1.MailFrom = Chr$(34) + "Exitstencil Press Australia Billing Services" & Chr$(34) + " <jcosti@ep.net.au>"
            SMTP1.MessageSubject = "Receipt #" & iReceiptNo & " for Internet Services"
            cDebug "Sending " & SMTP1.MessageSubject
            
            SMTP1.SendTo = ""
            
            AddFilename = ""
            If MySQL.OpenTable(oConn, rseMail, , "select AES_DECRYPT(EmailAddress,'" & odb.colSalts.ReturnSalt(EMAILSalt) & "') as Emailaddress from acci_emailaddresses where Checked = -1 and AccI_RecID = " & rsload!acci_RecID) = True Then
                If IsNull(rseMail!EmailAddress) Then AddFilename = "*"
                If rseMail.RecordCount > 0 Then
                 While Not rseMail.EOF
                    SMTP1.SendTo = SMTP1.SendTo + IIf(IsNull(rseMail!EmailAddress), "jcosti@ep.net.au", rseMail!EmailAddress) + ","
                    rseMail.MoveNext
                 Wend
                Else
                    SMTP1.SendTo = "jcosti@ep.net.au,"
                End If
            Else
                SMTP1.SendTo = "jcosti@ep.net.au,"
            End If
            SMTP1.SendTo = Left(SMTP1.SendTo, Len(SMTP1.SendTo) - 1) & ", " & oResell.ReturnCatch(IIf(IsNull(rsload!VirtualID), Login.lVirtualID, rsload!VirtualID), "receipts")
            If Login.bTestBench = True Or SMTP1.SendTo = "" Then SMTP1.SendTo = "jcosti@ep.net.au" & ", " & oResell.ReturnCatch(IIf(IsNull(rsload!VirtualID), Login.lVirtualID, rsload!VirtualID), "receipts")
            SMTP1.SendTo = MySQL.ReplaceString(SMTP1.SendTo, ",,", ",")
            SMTP1.Server = reg.smtpServer
            SMTP1.Port = reg.smtpPort
            SMTP1.Username = reg.smtpUsername
            SMTP1.Password = reg.smtpPassword
            
            If IsNull(HTML) Then GoTo SkipSend1
            
            SMTP1.MessageHTML = MySQL.ReplaceString(HTML, "/EmailAddy/", SMTP1.SendTo)
            
            SMTP1.Attachments.Add cSMTP.SaveHTML(SMTP1.MessageHTML, AddFilename & SMTP1.MessageSubject + ".html", sysnow)
            If Login.bTestBench = False Then
                SMTP1.SendEmail
                
                For Att = SMTP1.Attachments.Count To 1 Step -1
                    SMTP1.Attachments.Remove Att
                Next
                If Not Frame5(0).Caption = SMTP1.Status + " - " & SMTP1.MessageSubject Then Frame5(0).Caption = SMTP1.Status + " - " & SMTP1.MessageSubject
                Do
                    'SMTP1.SendEmail
                    gSleep
                    If Not Frame5(0).Caption = SMTP1.Status + " - " & SMTP1.MessageSubject Then Frame5(0).Caption = SMTP1.Status + " - " & SMTP1.MessageSubject
                Loop Until SMTP1.Status = "SMTP session closed" Or SMTP1.Status = "" Or Err.Number <> 0
            End If
            

SkipSend1:
            pbBilling.Value = pbBilling.Value + 1
            rsload.MoveNext
        Wend
    End If
    
    pbBilling.Value = pbBilling.Max
    
    lblBillingStats.Caption = Format(cChargeTotal, "Currency") & " - Processed Into Account Receivable" + vbCrLf
    lblBillingStats.Caption = lblBillingStats.Caption + Format(cChargeTotal * oTax(Login.TaxCode, Login.TaxCountry), "Currency") & " - GST Processed Into Account Receivable"
    
    If Err.Number = 0 Then Exit Function
    

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

Public Function ProcessQuota()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "ProcessQuota"
    Const ContainerName = "frmBot"
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
    
    Dim rsload As adodb.Recordset
    Dim rsload2 As adodb.Recordset
    Dim rsPlanType As adodb.Recordset
    Dim rsRadius As adodb.Recordset
    Dim rsBooleanMSG As adodb.Recordset
    Dim bx As Byte
    Dim lRecCount As Variant
    
    bResult = MySQL.OpenTable(directConn, rsload, , "select Count(*) As RecordCount from acci_services where RadiusID <> 0")
    bResult = MySQL.OpenTable(directConn, rsPlanType, , "select * from plantypes")
    
    lRecCount = rsload!RecordCount
    
    'Call MySQL.DAOConn("projectalpha", sServer, sUID, sPWD, dConn, wrkODBC)
        
    Dim ix As Variant
    
    If lRecCount > 0 Then
        pbQuota.Max = lRecCount
        pbQuota.Value = 0
        For ix = 0 To lRecCount Step 30
            bResult = MySQL.OpenTable(directConn, rsload, , "select * from acci_services where RadiusID <> 0 Limit " & ix & ",30")
            While Not rsload.EOF And Err.Number = 0
                rsPlanType.Filter = "RecID = " & rsload!ptRecID
                
                If Not rsPlanType.EOF And Not rsPlanType.BOF Then
                    bResult = MySQL.OpenTable(directConn, rsRadius, , "select * from radiusaccounts where RecID = " & rsload!RadiusID & " Limit 1")
                    
                    If Not rsRadius.EOF And Not rsRadius.BOF Then
                        If rsRadius.RecordCount > 0 Then
                            If rsPlanType!MBPerPeriod <> -1 Then
                                If (rsRadius!sfCycle_Download / 1024 ^ 2) / rsPlanType!MBPerPeriod > 0.85 Then
                                    Call MySQL.AddQuotaItem(directConn, rsload!acci_RecID, "Data Limits exceeding 85 percent of Quota", rsPlanType!MBPerPeriod - (rsRadius!sfCycle_Download / 1024 ^ 2), (rsRadius!sfCycle_Download / 1024 ^ 2), "Mb", 2)
                                End If
                            End If
                            
                            If rsPlanType!HoursPerPeriod <> -1 Then
                                If (rsRadius!sfCycle_Mins / 60) / rsPlanType!HoursPerPeriod > 0.7 Then
                                    Call MySQL.AddQuotaItem(directConn, rsload!acci_RecID, "Time Limits exceeding 85 percent of Quota", rsPlanType!HoursPerPeriod * 60 - (rsRadius!sfCycle_Mins), rsRadius!sfCycle_Mins, "min(s)", 3)
                                End If
                            End If
                        End If
                    End If
                End If
                pbQuota.Value = pbQuota.Value + 1
                rsload.MoveNext
            Wend
        Next ix
    Else
        pbQuota.Value = 1
        pbQuota.Max = 1
    End If
    
    pbQuota.Value = pbQuota.Max
    
    Dim rseMail As adodb.Recordset
    Dim rsMaxi As adodb.Recordset
    Dim bDir As Boolean
    Dim iQuotaID As Long
    
    bResult = MySQL.OpenTable(directConn, rsload, , "select distinct acci_RecID from acci_quotareceipt where QuotaMSGSent = 0")
    
    If rsload.RecordCount > 0 Then
    
        Select Case bDir
        Case False
            bResult = MySQL.OpenTable(directConn, rsMaxi, , "select max(QuotaMSGID) as nResult from acci_quotareceipt")
            iQuotaID = IIf(IsNull(rsMaxi!nResult), Rnd * 9999999, rsMaxi!nResult) + 1
            bDir = True
        Case True
            bDir = False
            bResult = MySQL.OpenTable(directConn, rsMaxi, , "select min(QuotaMSGID) as nResult from acci_quotareceipt")
            iQuotaID = IIf(IsNull(rsMaxi!nResult), -Rnd * 9999999, rsMaxi!nResult) - 1
        End Select
        
        MySQL.Execute directConn, "update acci_quotareciept set QuotaMSGID = '" & iQuotaID & "' where acci_RecID = '" & rsload!acci_RecID & "' and QuotaMSGSent = 0"
        
        pbQuota.Max = pbQuota.Max + rsload.RecordCount
        While Not rsload.EOF And Err.Number = 0
            
            HTML = cSMTP.SendQuotaHTML(directConn, rsload!acci_RecID, iQuotaID)
            
            SMTP2.MailDate = Format(sysnow, "ddddd ttttt")
            SMTP2.MailFrom = Chr$(34) + "Exitstencil Press Australia Billing Services" & Chr$(34) + " <jcosti@ep.net.au>"
            SMTP2.MessageSubject = "Customer No #" & Hex(rsload!acci_RecID) & " - Quota Warning " & iQuotaID
            
            'Call cSMTP.SaveHTML(html, SMTP2.MessageSubject + ".html", Sysnow)
            AddFilename = ""
            If MySQL.OpenTable(directConn, rseMail, , "select AES_DECRYPT(EmailAddress,'" & odb.colSalts.ReturnSalt(EMAILSalt) & "') as Emailaddress from acci_emailaddresses where Checked = -1 and AccI_RecID = " & rsload!acci_RecID) = True Then
                If IsNull(rseMail!EmailAddress) Then AddFilename = "*"
                If rseMail.RecordCount > 0 Or Not rseMail Is Nothing Then
                 While Not rseMail.EOF And Err.Number = 0
                    SMTP2.SendTo = SMTP2.SendTo + IIf(IsNull(rseMail!EmailAddress), "jcosti@ep.net.au,", rseMail!EmailAddress + ",")
                    rseMail.MoveNext
                 Wend
                Else
                    SMTP2.SendTo = "jcosti@ep.net.au,"
                End If
            Else
                SMTP2.SendTo = "jcosti@ep.net.au,"
            End If
            SMTP2.SendTo = Left(SMTP2.SendTo, Len(SMTP2.SendTo) - 1)
            If Login.bTestBench = True Or SMTP2.SendTo = "" Then SMTP2.SendTo = "jcosti@ep.net.au" 'Left(SMTP2.SendTo, Len(SMTP2.SendTo) - 1)
            SMTP2.Server = reg.smtpServer
            SMTP2.Port = reg.smtpPort
            If IsNull(HTML) Then GoTo SkipSend1
            SMTP2.MessageHTML = MySQL.ReplaceString(HTML, "/EmailAddy/", SMTP2.SendTo)
            SMTP2.Attachments.Add cSMTP.SaveHTML(SMTP2.MessageHTML, AddFilename & SMTP2.MessageSubject + ".html", sysnow)
            SMTP2.Username = reg.smtpUsername
            SMTP2.Password = reg.smtpPassword
            If Login.bTestBench = False Then
                If Not frameEmails.Caption = SMTP2.Status + " - " & SMTP2.MessageSubject Then Frame5(0).Caption = SMTP2.Status + " - " & SMTP2.MessageSubject
                Do
                    SMTP2.SendEmail
                    gSleep
                    If Not frameEmails.Caption = SMTP2.Status + " - " & SMTP2.MessageSubject Then Frame5(0).Caption = SMTP2.Status + " - " & SMTP2.MessageSubject
                Loop Until SMTP2.Status = "SMTP session closed" Or SMTP2.Status = "" Or Err.Number <> 0
            End If
            SMTP2.Attachments.Remove 1
SkipSend1:
            pbQuota.Value = pbQuota.Value + 1
            rsload.MoveNext
        Wend
    End If
    
    pbQuota.Value = pbQuota.Max
    
    If Err.Number = 0 Then Exit Function
    
Exit Function



ErrorOccur:
pbQuota.Value = pbQuota.Max
Select Case oErr.chkError(directConn, Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function

Public Function ProcessStatements()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "ProcessStatements"
    Const ContainerName = "frmBot"
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
    
    Const SendMail = True
    
    Dim bResult As Boolean
    Dim rsload As adodb.Recordset
    Dim rsInvOut As adodb.Recordset
    Dim rsTraxr As adodb.Recordset
    Dim cAmountDue As Currency
    Dim cAmountPaid As Currency
    Dim iTraxrID As Variant
    Dim InvCreated As Variant
    Dim Att As Long
    
    bResult = MySQL.OpenTable(directConn, rsload, , "select * from accountinfo where Cancelled = 0 and BillingDate <= '" & Format(sysnow, "YYYY-MM-DD Hh:Nn:SS") & "'")
    Dim dStatementNo As Double
    Dim rseMail As adodb.Recordset
    
    If rsload.RecordCount > 0 Then
    
        pbBilling.Max = pbBilling.Max + rsload.RecordCount
        While Not rsload.EOF And Err.Number = 0
            
            If MySQL.OpenTable(directConn, rsTraxr, , "select RecID from acci_services where acci_RecID = " & rsload!RecID) = True Then
                If rsTraxr.RecordCount = 0 Then GoTo SkipSend
            End If
            
            HTML = cSMTP.sendStatementHTML(directConn, rsload!RecID, dStatementNo)
            
            SMTP3.MailDate = Format(sysnow, "ddddd ttttt")
            SMTP3.MailFrom = Chr$(34) + "Exitstencil Press Australia Billing Services" & Chr$(34) + " <jcosti@ep.net.au>"
            SMTP3.MessageSubject = "Statement #" & dStatementNo & " for Internet Services"
            
            AddFilename = ""
            If MySQL.OpenTable(directConn, rseMail, , "select AES_DECRYPT(EmailAddress,'" & odb.colSalts.ReturnSalt(EMAILSalt) & "') as Emailaddress from acci_emailaddresses where Checked = -1 and AccI_RecID = " & rsload!RecID) = True Then
                If IsNull(rseMail!EmailAddress) Then AddFilename = "*"
                If rseMail.RecordCount > 0 Then
                 While Not rseMail.EOF And Err.Number = 0
                    SMTP3.SendTo = SMTP3.SendTo + IIf(IsNull(rseMail!EmailAddress), "jcosti@ep.net.au,", rseMail!EmailAddress + ",")
                    rseMail.MoveNext
                 Wend
                Else
                    SMTP3.SendTo = "jcosti@ep.net.au,"
                End If
            Else
                SMTP3.SendTo = "jcosti@ep.net.au,"
            End If
            SMTP3.SendTo = Left(SMTP3.SendTo, Len(SMTP3.SendTo) - 1)
            If Login.bTestBench = True Or SMTP3.SendTo = "" Then SMTP3.SendTo = "jcosti@ep.net.au"
            SMTP3.Server = reg.smtpServer
            SMTP3.Port = reg.smtpPort
            If IsNull(HTML) Then GoTo SkipSend
            
            SMTP3.MessageHTML = MySQL.ReplaceString(HTML, "/EmailAddy/", SMTP3.SendTo)
            SMTP3.Attachments.Add cSMTP.SaveHTML(SMTP3.MessageHTML, AddFilename + SMTP3.MessageSubject + ".html", sysnow)
            
            If SendMail = True And Len(SMTP3.MessageHTML) > 0 Then
                SMTP3.SendEmail
                If Not Frame5(0).Caption = SMTP3.Status + " - " & SMTP3.MessageSubject Then Frame5(0).Caption = SMTP3.Status + " - " & SMTP3.MessageSubject
                Do
                    SMTP3.SendEmail
                    gSleep
                    If Not Frame5(0).Caption = SMTP3.Status + " - " & SMTP3.MessageSubject Then Frame5(0).Caption = SMTP3.Status + " - " & SMTP3.MessageSubject
                Loop Until SMTP3.Status = "SMTP session closed" Or SMTP3.Status = "" Or Err.Number <> 0
            End If
            
            For Att = SMTP3.Attachments.Count To 1 Step -1
                SMTP3.Attachments.Remove Att
            Next
SkipSend:
            'pbBilling.Value = pbBilling.Value + 1

            MySQL.Execute directConn, "update accountinfo set BillingDate = '" & Format(DateAdd("m", 1, IIf(IsNull(rsload!BillingDate), sysnow, rsload!BillingDate)), "yyyy-mm-dd ttttt") & "' where RecID = " & rsload!RecID
            'rsLoad.Update
            rsload.Cancel
            rsload.MoveNext
            InvCreated = InvCreated + 1
            pbBilling.Value = pbBilling.Value + 1
        Wend
    End If
    
    lblBillingStats.Caption = lblBillingStats.Caption + vbCrLf + vbCrLf & InvCreated & " Statements Created"
    
    pbBilling.Value = pbBilling.Max
    
    Frame5(0).Caption = "Process Completed"
    
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


Public Function ProcessPO()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "ProcessPO"
    Const ContainerName = "frmBot"
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
    Dim SQL As String
    
    Call MySQL.OpenTable(directConn, rsload, , "select count(distinct RecID) as RecCount from acci_services where POID = 0")
    
    If rsload!RecCount > 0 Then
        pbQuota.Max = pbQuota.Max + rsload!RecCount
        
        SQL = "select distinct acci_services.RecID, acci_services.DefaultShippingID as ShippingID, vendors.RecID as VendorID, accountinfo.RecID as AccI_RecID, servicetypes.ServiceKey, plantemplates.VendorPartID, " + _
                "plantemplates.SubPartID, plantemplates.Description, vendors.vName, plantemplates.CostPrice, plantemplates.MBQuota, acci_services.Username, AES_DECRYPT(Password, 'salt') as Password " + _
                "from plantemplates left join plantypes on plantemplates.RecID = plantypes.TemplateID " + _
                "left join vendors on vendors.RecID = plantypes.VendorID " + _
                "left join acci_services on acci_services.ptRecID = plantypes.RecID " + _
                "left join accountinfo on acci_services.AccI_RecID = accountinfo.RecID " + _
                "left join servicetypes on acci_services.ServiceID = servicetypes.RecID " + _
                "where acci_services.POID = 0 order by accountinfo.RecID, vendors.RecID "
               
                
        Call MySQL.OpenTable(directConn, rsload, , SQL)
        
        Dim VendorID As Long
        Dim acci_RecID As Long
        Dim POValue As Single
        Dim POGST As Single
        Dim SavePO As Boolean
        Dim POID As Double
        Dim NEWPO As Boolean
        
        If rsload.RecordCount > 0 Then
            While Not rsload.EOF And Err.Number = 0
                
                VendorID = IIf(IsNull(rsload!VendorID), 0, rsload!VendorID)
                acci_RecID = IIf(IsNull(rsload!acci_RecID), 0, rsload!acci_RecID)
                
                If POID = 0 Or NEWPO = True Then
                    On Error Resume Next
                    Do
                        POID = MySQL.GetTMPRecID("purchaseorder", directConn)
                        Call MySQL.Execute(directConn, "insert into purchaseorder (RecID, VendorID, acci_RecID) VALUES('" & POID & "','" & VendorID & "','" & acci_RecID & "')")
                        If Err.Number = 0 Then Exit Do
                        gSleep
                    Loop
                    NEWPO = False
                End If
                
                Call MySQL.Execute(directConn, "Update acci_services set POID = '" & POID & "' where RecID = " & rsload!RecID)
                
                rsload.MoveNext
                pbQuota.Value = pbQuota.Value + 1
                
                If Not rsload.EOF Then
                    If Not Val(acci_RecID) = Val(rsload!acci_RecID) Then SavePO = True
                    If Not Val(VendorID) = Val(rsload!VendorID) Then SavePO = True
                Else
                    SavePO = True
               End If
                
                Call MySQL.Execute(directConn, "Update purchaseorder set POValue = POValue + " & IIf(IsNull(rsload!CostPrice), 0, rsload!CostPrice) & " where RecID = " & POID)
                Call MySQL.Execute(directConn, "Update purchaseorder set POGST = POGST + " & IIf(IsNull(rsload!CostPrice), 0, rsload!CostPrice) * oTax(Login.TaxCode, Login.TaxCountry) & " where RecID = " & POID)
                
                If SavePO = True Then
                    Call MySQL.Execute(directConn, "Update purchaseorder set DateSent = '1899-12-31 12:00:00', ShippingID = '" & IIf(IsNull(rsload!ShippingID), 0, rsload!ShippingID) & "' where RecID = " & POID)
                    NEWPO = True
                    SavePO = False
                    POGST = 0
                End If
                
            Wend
        End If
        
        Dim sql2 As String
        Dim rssend As adodb.Recordset
        Dim rsTMP As adodb.Recordset
        Dim Att As Long
        
        sql2 = "select vendors.*, purchaseorder.* from purchaseorder, vendors where purchaseorder.VendorID = vendors.RecID and purchaseorder.DateSent = '1899-12-31 12:00:00'"
        
        
            Call MySQL.OpenTable(directConn, rssend, , sql2)
            If rssend.RecordCount > 0 Then
                pbQuota.Max = pbQuota.Max + rssend.RecordCount
                While Not rssend.EOF And Err.Number = 0
                    MySQL.Execute directConn, "update purchaseorder set DateSent = '" + "NOW()" + "' where RecID = " & rssend!RecID
                    
                    Call MySQL.OpenTable(directConn, rsTMP, , "select DefaultShippingID from acci_services where POID = " & rssend!RecID & " limit 1")
                    
                    SMTP2.MailDate = Format(sysnow, "ddddd ttttt")
                    SMTP2.MailFrom = Chr$(34) + "Exitstencil Press Australia Billing Services" & Chr$(34) + " <jcosti@ep.net.au>"
                    SMTP2.MessageSubject = "Purchase Order Number #" & MySQL.ReplaceString("" & rssend!RecID, "-", "J-") & ""
                    
                    SMTP2.SendTo = oResell.ReturnCatch(IIf(IsNull(rssend!VirtualID), Login.lVirtualID, rssend!VirtualID), "po") & ", " & IIf(IsNull(rssend!poemailaddy), "jcosti@ep.net.au", rssend!poemailaddy)
                                                
                    If Login.bTestBench = True Or SMTP2.SendTo = "" Then SMTP2.SendTo = "jcosti@ep.net.au"
                    SMTP2.Server = reg.smtpServer
                    SMTP2.Port = reg.smtpPort
                    SMTP2.Username = reg.smtpUsername
                    SMTP2.Password = reg.smtpPassword
                    HTML = "" & cSMTP.GetPOHTML(rssend!RecID, rssend!acci_RecID, IIf(IsNull(rsTMP!DefaultShippingID), 0, rsTMP!DefaultShippingID), directConn)
                     SMTP2.MessageHTML = MySQL.ReplaceString(HTML, "/EmailAddy/", SMTP2.SendTo)
                     
                    For Att = SMTP2.Attachments.Count To 1 Step -1
                        SMTP2.Attachments.Remove Att
                    Next
                    
                    SMTP2.Attachments.Add cSMTP.SaveHTML(SMTP2.MessageHTML, AddFilename & SMTP2.MessageSubject + ".html", sysnow), "a1"
                    'SMTP2.SendTo = "jcosti@ep.net.au"
                    
                    
                    If Login.bTestBench = False Then SMTP2.SendEmail
                    



                    pbQuota.Value = pbQuota.Value + 1
                    rssend.MoveNext
                Wend
                pbQuota.Value = pbQuota.Max
        End If
    End If
    
    pbQuota.Value = pbQuota.Max
    
Exit Function



ErrorOccur:
pbQuota.Value = pbQuota.Max
Select Case oErr.chkError(directConn, Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function
