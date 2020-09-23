VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H004FD2F9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sysops Manifest"
   ClientHeight    =   10095
   ClientLeft      =   5145
   ClientTop       =   2895
   ClientWidth     =   10395
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10095
   ScaleWidth      =   10395
   Begin MSComctlLib.ImageList ilSysops 
      Left            =   6090
      Top             =   750
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
            Picture         =   "frmMain.frx":0442
            Key             =   "k100"
            Object.Tag             =   "100% Security Access"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":075C
            Key             =   "k080_099"
            Object.Tag             =   "High Level Access"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A76
            Key             =   "k060_079"
            Object.Tag             =   "Server Administrator"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1350
            Key             =   "k040_059"
            Object.Tag             =   "Network Admin"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17A2
            Key             =   "k001_020"
            Object.Tag             =   "Service Plans Editor"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BF4
            Key             =   "k020_039"
            Object.Tag             =   "Rainbow Warrior"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2046
            Key             =   "k000"
            Object.Tag             =   "Low Level Access"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0047ADE4&
      Caption         =   "Sysops"
      Height          =   9825
      Left            =   120
      TabIndex        =   5
      Top             =   90
      Width           =   10155
      Begin VB.CommandButton Command1 
         Caption         =   "Change P&ermissioning"
         Height          =   435
         Index           =   4
         Left            =   6420
         TabIndex        =   6
         Top             =   330
         Width           =   1725
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Refresh"
         Height          =   435
         Index           =   3
         Left            =   120
         TabIndex        =   1
         Top             =   330
         Width           =   1725
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Change &Password"
         Height          =   435
         Index           =   2
         Left            =   3720
         TabIndex        =   3
         Top             =   330
         Width           =   1725
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Delete Sysop"
         Height          =   435
         Index           =   1
         Left            =   8220
         TabIndex        =   4
         Top             =   330
         Width           =   1725
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Create Sysop"
         Height          =   435
         Index           =   0
         Left            =   1920
         TabIndex        =   2
         Top             =   330
         Width           =   1725
      End
      Begin MSComctlLib.ListView lvSysops 
         Height          =   8775
         Left            =   150
         TabIndex        =   0
         Top             =   870
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   15478
         View            =   3
         Arrange         =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ilSysops"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Level"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Username"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fullname"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Create Sysops"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Create VISPS"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Create Agencies"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Create Templates"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Primary Sysop"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Virtual ISP"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "Web Register"
            Object.Width           =   1764
         EndProperty
         Picture         =   "frmMain.frx":2498
      End
   End
   Begin VB.Menu mnuAct 
      Caption         =   "Action"
      Begin VB.Menu mnuAct_End 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuAct_Debug 
         Caption         =   "Debug Window"
         Shortcut        =   ^D
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)

    Select Case Index
    Case 0
        Dim oSysop As New frmSysop
        oSysop.RecID = 0
        oSysop.Show
    Case 1
        Select Case MsgBox("Are you sure you wish to delete this sysop?", vbYesNo, "Remove Sysop")
        Case vbYes
            oMySQL.Execute oConn, "update sysops set bDeleted = '-1' Where RecID = " & Mid(lvSysops.SelectedItem.Key, 2)
            lvSysops.ListItems.Remove lvSysops.SelectedItem.Index
        End Select
    Case 2
    
        Dim oPWD As New frmPWD
        oPWD.inPassword = lvSysops.SelectedItem.Tag
        oPWD.Show 1
        lvSysops.SelectedItem.Tag = oPWD.outPassword
        oMySQL.Execute oConn, "update sysops set Password = encode('" + oPWD.outPassword + "','" + PasswordSalt + "') Where RecID = " & Mid(lvSysops.SelectedItem.Key, 2)
    Case 3
        Call Form_Load
    Case 4
        Dim oPerm As New frmPerm
        oPerm.RecID = Val(Mid(lvSysops.SelectedItem.Key, 2))
        oPerm.Show
    
    End Select
    
End Sub

Private Sub Form_Load()

       
        
   Dim rsLoad As ADODB.Recordset
   
    Command1(0).Enabled = Login.bCreateSysop
    Command1(1).Enabled = Login.bCreateSysop
    Command1(4).Enabled = Login.bCreateSysop
     
   Call oGUI.LoadColWidths(lvSysops, Me)
        
        Me.Caption = "Searching for ViSP MAP"
    
        Dim rsPriMap As ADODB.Recordset
        
        Call oMySQL.OpenTable(oConn, rsPriMap, , "select distinct vispa.RecID as RecIDa, vispb.RecID as RecIDb, vispb.Description from accountviewer, virtualisp as vispa, virtualisp as vispb where vispa.RecID = vispb.VirtualID  and  vispb.VirtualID = '" & Login.lVirtualID & "'")
        
        ViSPMAP.Clear
        
        If rsPriMap.State = adStateOpen Then
        
            If rsPriMap.RecordCount > 0 Then
                While Not rsPriMap.EOF
                    If rsPriMap!RecIDb <> Login.lVirtualID Then ViSPMAP.Add "r" & ViSPMAP.Count + 1, Val(rsPriMap!RecIDa), Val(rsPriMap!RecIDb), IIf(IsNull(rsPriMap!Description), "\-\", rsPriMap!Description), "r" & ViSPMAP.Count + 1
                    rsPriMap.MoveNext
                Wend
            End If
            
         End If
        
        rsPriMap.Close
        
        Dim allchecked As Long
        
        If ViSPMAP.Count > 0 Then
            Do
                Me.Caption = "ViSP Map [ " & ViSPMAP.Count & " visp found so far ]"
                allchecked = allchecked + 1
                
                Call oMySQL.OpenTable(oConn, rsPriMap, , "select distinct vispa.RecID as RecIDa, vispb.RecID as RecIDb, vispb.Description from accountviewer, virtualisp as vispa, virtualisp as vispb where vispa.RecID = vispb.VirtualID  and vispb.VirtualID = '" & Val(ViSPMAP(allchecked).RecIDb) & "'")
                
                'ViSPMAP.Clear
                
                If rsPriMap.State = adStateOpen Then
                
                    If rsPriMap.RecordCount > 0 Then
                        While Not rsPriMap.EOF
                            If rsPriMap!RecIDb <> Login.lVirtualID Then ViSPMAP.Add "r" & ViSPMAP.Count + 1, Val(rsPriMap!RecIDa), Val(rsPriMap!RecIDb), IIf(IsNull(rsPriMap!Description), "\-\", rsPriMap!Description), "r" & ViSPMAP.Count + 1
                            rsPriMap.MoveNext
                        Wend
                    End If
                    
                 End If
                
                rsPriMap.Close
                            
                'lvSchedule.ListItems(ix).SubItems(3) = ((allchecked / ViSPMAP.Count) * 100) & "%"
            Loop Until ViSPMAP.Count = allchecked
        End If
            
        
        
        Call oMySQL.OpenTable(oConn, rsPriMap, , "select RecID as RecIDa, RecID as RecIDb, Description from virtualisp where RecID = '" & Login.lVirtualID & "'")
        If rsPriMap.State = adStateOpen Then
            If rsPriMap.RecordCount >= 1 Then
                
                ViSPMAP.Add "r0", Val(rsPriMap!RecIDa), Val(rsPriMap!RecIDb), IIf(IsNull(rsPriMap!Description), "\-\", rsPriMap!Description), "r0"
                
            End If
         End If
        rsPriMap.Close
        
        Me.Caption = "Sysop Control Centre - version [" & App.Major & "." & App.Minor & ".0." & App.Revision & "]"
        
        
        Dim sql As String
        
        If ViSPMAP.Count > 0 Then
            sql = "("
            Dim iCnt As Long
            
            For iCnt = 1 To ViSPMAP.Count
                sql = sql + "sysops.VirtualID = '" & ViSPMAP(iCnt).RecIDb & "' Or "
            Next iCnt
            
            sql = Left(sql, Len(sql) - 4) + ")"
        Else
            sql = "sysops.VirtualID = " & Login.lVirtualID
        End If
        
        bResult = oMySQL.OpenTable(oConn, rsLoad, , "Select sysops.Checked, sysops.SecurityLevel, sysops.RecID, sysops.Username, sysops.Description, sysops.bWEBAccount, bVISP, bCreateSysop, bVISPFiscal, bTemplates, bPrimary, bAgency, Firstname, Surname, Master, decode(Password,'" + PasswordSalt + "') as Password, virtualisp.description as VrtDESC from sysops, virtualisp where bDeleted = 0 and sysops.VirtualID = virtualisp.RecID and " & sql)

    
    lvSysops.ListItems.Clear

    If rsLoad.BOF And rsLoad.EOF Then
    
    Else
    
        If rsLoad.RecordCount > 0 Then
            rsLoad.MoveFirst
            While Not rsLoad.EOF
                
                Set itmX = lvSysops.ListItems.Add(, "r" & rsLoad!RecID, String(3 - Len("" & rsLoad!SecurityLevel), "0") + "" & rsLoad!SecurityLevel)
                itmX.SubItems(1) = IIf(IsNull(rsLoad!Username), "", rsLoad!Username)
                itmX.SubItems(2) = IIf(IsNull(rsLoad!Firstname), "", rsLoad!Firstname) + " " + IIf(IsNull(rsLoad!Surname), "", rsLoad!Surname)
                itmX.SubItems(3) = IIf(Val(rsLoad!bCreateSysop) = False, "NO", "YES")
                itmX.SubItems(4) = IIf(Val(rsLoad!bVISP) = False, "NO", "YES")
                itmX.SubItems(5) = IIf(Val(rsLoad!bAgency) = False, "NO", "YES")
                itmX.SubItems(6) = IIf(Val(rsLoad!bTemplates) = False, "NO", "YES")
                itmX.SubItems(7) = IIf(Val(rsLoad!bPrimary) = False, "NO", "YES")
                itmX.SubItems(8) = IIf(IsNull(rsLoad!VrtDESC), "", rsLoad!VrtDESC)
                itmX.SubItems(9) = IIf(Val(rsLoad!bWEBAccount) = 0, "No", "Yes")
        
                itmX.Tag = rsLoad!Password
                itmX.Checked = IIf(rsLoad!Checked <> 0, True, False)
                
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
                    If CByte(rsLoad!SecurityLevel) >= imgMin And CByte(rsLoad!SecurityLevel) <= imgMax Then
                        itmX.SmallIcon = imgX
                    End If
                Next imgX
                
                rsLoad.MoveNext
            Wend
        End If
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    End
    Call oGUI.SaveColWidths(lvSysops, Me)

End Sub

Private Sub lvSysops_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Call oGUI.ColumnSort(ColumnHeader, lvSysops)
    
End Sub

Private Sub lvSysops_DblClick()

    
    If lvSysops.SelectedItem Is Nothing Then
    
    Else
        
        If Login.bMaster = False And Login.bPrimary = False Then
            If Login.lSysopID <> Val(Mid(lvSysops.SelectedItem.Key, 2)) Then Exit Sub
        End If
        
        Dim oSysop As New frmSysop
        oSysop.RecID = Val(Mid(lvSysops.SelectedItem.Key, 2))
        oSysop.Show
    End If
    
End Sub

Private Sub mnuAct_Debug_Click()

    frmDebug.Show
    
End Sub

Private Sub mnuAct_End_Click()

    End
    
End Sub
