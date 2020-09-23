VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDOC 
   BackColor       =   &H00A9C9AE&
   Caption         =   "Document Browser"
   ClientHeight    =   10875
   ClientLeft      =   1695
   ClientTop       =   2400
   ClientWidth     =   15900
   Icon            =   "frmDOC_refresh.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   10875
   ScaleWidth      =   15900
   Begin VB.PictureBox SMTP1 
      Height          =   480
      Left            =   6600
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   11
      Top             =   900
      Width           =   1200
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   7170
      Top             =   900
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ilDocs 
      Left            =   12180
      Top             =   9120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDOC_refresh.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDOC_refresh.frx":0894
            Key             =   "selfolder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDOC_refresh.frx":0CE6
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDOC_refresh.frx":1138
            Key             =   "visp"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDOC_refresh.frx":158A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDOC_refresh.frx":19DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDOC_refresh.frx":1E2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDOC_refresh.frx":2280
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDOC_refresh.frx":26D2
            Key             =   "item"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDOC_refresh.frx":2B24
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDOC_refresh.frx":2F76
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDOC_refresh.frx":33C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDOC_refresh.frx":381A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDOC_refresh.frx":3B34
            Key             =   "year"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDOC_refresh.frx":3F86
            Key             =   "month"
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmDocs 
      BackColor       =   &H00A9C9AE&
      Caption         =   "History"
      Height          =   9015
      Index           =   3
      Left            =   6090
      TabIndex        =   6
      Top             =   1500
      Width           =   10005
      Begin MSComctlLib.ListView www 
         Height          =   8565
         Left            =   180
         TabIndex        =   7
         Top             =   270
         Width           =   9705
         _ExtentX        =   17119
         _ExtentY        =   15108
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
   Begin VB.Frame frmDocs 
      BackColor       =   &H00A9C9AE&
      Caption         =   "Document Generated So Far"
      Height          =   10665
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   5565
      Begin VB.CommandButton cmbResend 
         Caption         =   "Resend Document to Email address Specified"
         Height          =   345
         Left            =   150
         TabIndex        =   5
         Top             =   10230
         Width           =   5295
      End
      Begin VB.Frame frmDocs 
         BackColor       =   &H00A9C9AE&
         Caption         =   "Document Tree"
         Height          =   7785
         Index           =   2
         Left            =   150
         TabIndex        =   3
         Top             =   2340
         Width           =   5265
         Begin MSComctlLib.TreeView tvDocs 
            Height          =   6435
            Left            =   150
            TabIndex        =   4
            Top             =   300
            Width           =   4995
            _ExtentX        =   8811
            _ExtentY        =   11351
            _Version        =   393217
            Indentation     =   441
            LabelEdit       =   1
            LineStyle       =   1
            Sorted          =   -1  'True
            Style           =   7
            HotTracking     =   -1  'True
            SingleSel       =   -1  'True
            ImageList       =   "ilDocs"
            Appearance      =   1
         End
      End
      Begin VB.Frame frmDocs 
         BackColor       =   &H00A9C9AE&
         Caption         =   "Document Type"
         Height          =   1725
         Index           =   1
         Left            =   150
         TabIndex        =   1
         Top             =   270
         Width           =   5265
         Begin VB.CommandButton cmbAction 
            Caption         =   "&Save Document to a local destination..."
            Enabled         =   0   'False
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   10
            Top             =   1230
            Width           =   4995
         End
         Begin VB.CommandButton cmbAction 
            Caption         =   "&Email Document"
            Enabled         =   0   'False
            Height          =   375
            Index           =   1
            Left            =   2640
            TabIndex        =   9
            Top             =   810
            Width           =   2475
         End
         Begin VB.CommandButton cmbAction 
            Caption         =   "&View Document"
            Enabled         =   0   'False
            Height          =   375
            Index           =   0
            Left            =   150
            TabIndex        =   8
            Top             =   810
            Width           =   2445
         End
         Begin VB.ComboBox cmbDoc 
            BackColor       =   &H00A9C9AE&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   270
            Width           =   5025
         End
      End
   End
End
Attribute VB_Name = "frmDOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rsStationary As adodb.Recordset

Private Sub cmbAction_Click(Index As Integer)

    Dim HTML
    
    Dim HTMLIn As New smtp
    Dim rsload As adodb.Recordset
    
    Call MySQL.OpenTable(directConn, rsload, , "select * from " & rsStationary!Tablename & " where " & rsStationary!FieldName & " = '" & Mid(tvDocs.SelectedItem.Key, 2) & "'")
    
    If rsload.State = adStateOpen Then
        Select Case rsStationary!StationaryCode
        Case "INVOICE"
            HTML = HTMLIn.sendInvoiceHTML(directConn, rsload!acci_RecID, rsload!RecID, 1)
            HTML = MySQL.ReplaceString(HTML, "/InvoiceNumber/", "" & rsload!RecID & "")
            SMTP1.MessageSubject = "[PASSON] Invoice " & rsload!RecID & " - generated by The Nexus"
        Case "RECEIPT"
            HTML = HTMLIn.sendReceiptHTML(directConn, rsload!acci_RecID, rsload!receiptNo, True)
            SMTP1.MessageSubject = "[PASSON] Receipt " & rsload!receiptNo & " - generated by The Nexus"
        Case "STATEMENT"
            HTML = HTMLIn.sendStatementHTML(directConn, rsload!acci_RecID, rsload!RecID, True)
            SMTP1.MessageSubject = "[PASSON] Statement " & rsload!RecID & " - generated by The Nexus"
        Case "QUOTA"
            HTML = HTMLIn.SendQuotaHTML(directConn, rsload!acci_RecID, rsload!QuotaMSGID)
            SMTP1.MessageSubject = "[PASSON] Quota Warnings " & rsload!QuotaMSGID & " - generated by The Nexus"
        Case "PURCHASEORDER"
            HTML = HTMLIn.GetPOHTML(rsload!RecID, rsload!acci_RecID, rsload!ShippingID, directConn)
            SMTP1.MessageSubject = "[PASSON] Purchase Order " & rsload!RecID & " - generated by The Nexus"
        End Select
        
        Select Case Index
        
        Case 0 ' View
        
            If InStr(tvDocs.SelectedItem.Text, "[") > 0 Then
                trp = Mid(tvDocs.SelectedItem.Text, 1, InStr(tvDocs.SelectedItem.Text, "[") - 1)
            End If
            ShellExecute Me.hWnd, vbNullString, HTMLIn.SaveHTML(HTML, rsStationary!StationaryCode & " " & IIf(InStr(tvDocs.SelectedItem.Text, "[") > 0, trp, tvDocs.SelectedItem.Text) & ".html", sysnow), vbNullString, "C:\", SW_SHOWNORMAL

        Case 1 ' Email Document
            
            Dim fAddy As New frmEmailAddies
            Dim rseMail As adodb.Recordset
            
            
            
            cDebug "Sending " & SMTP1.MessageSubject
            
            AddFilename = ""
            If MySQL.OpenTable(directConn, rseMail, , "select AES_DECRYPT(EmailAddress,'" & odb.colSalts.ReturnSalt(EMAILSalt) & "') as Emailaddress from acci_emailaddresses where Checked = -1 and AccI_RecID = " & rsload!acci_RecID) = True Then
                If rseMail.State = adStateOpen Then
                    If rseMail.RecordCount > 0 Then
                     While Not rseMail.EOF And Err.Number = 0
                        fAddy.sTo = fAddy.sTo + IIf(IsNull(rseMail!EmailAddress), "jcosti@ep.net.au", rseMail!EmailAddress) + "; "
                        rseMail.MoveNext
                     Wend
                    Else
                        fAddy.sTo = "jcosti@ep.net.au; "
                    End If
                End If
            Else
                fAddy.sTo = "jcosti@ep.net.au; "
            End If
            fAddy.sTo = Left(fAddy.sTo, Len(fAddy.sTo) - 2)
            
            fAddy.Show
            
            If fAddy.ButPush = 1 Then Exit Sub
            
            SMTP1.SendTo = fAddy.sCC
            SMTP1.CC = fAddy.sCC
            SMTP1.BCC = fAddy.sBC
            
            SMTP1.UserName = reg.smtpUsername
            SMTP1.Password = reg.smtpPassword
            SMTP1.Server = reg.smtpServer
            SMTP1.Port = reg.smtpPort
            
            SMTP1.MessageHTML = MySQL.ReplaceString(HTML, "/EmailAddy/", fAddy.sTo)
            SMTP1.Attachments.Add cSMTP.SaveHTML(SMTP1.MessageHTML, AddFilename & SMTP1.MessageSubject + ".html", sysnow)
            
            SMTP1.SendEmail
            
            For Att = SMTP1.Attachments.Count To 1 Step -1
                SMTP1.Attachments.Remove Att
            Next
        
            Me.Caption = "Finished sending (" & SMTP1.MessageSubject & ")"
            
        Case 2 ' Save to local Destination
            cd.Filter = "HTML File (*.html)|*.html|HTML File (*.htm)|*.htm"
            cd.FilterIndex = 1
            cd.Filename = ""
            cd.ShowSave
            
            If cd.Filename <> "" Then
                Dim iFileNum As Integer
                
                iFileNum = FreeFile
                
                If Dir(cd.Filename, vbNormal) <> "" Then
                    Select Case MsgBox("The file " & cd.Filename & " already exists!! Are you sure you wish to over write it?", vbYesNo + vbCritical, "File Exists")
                    Case vbYes
                        Kill cd.Filename
                    Case vbNo
                        Exit Sub
                    End Select
                End If
                
                Open cd.Filename For Output As #iFileNum
                Print #iFileNum, HTML
                Close #iFileNum
                
            
            End If
        
        End Select
    End If
    
    
End Sub

Private Sub cmbDoc_Click()

    
    
    If tvDocs.Tag = cmbDoc.Text Then Exit Sub
    tvDocs.Tag = cmbDoc.Text
    
    tvDocs.Nodes.Clear
    
    rsStationary.Filter = "RecID = '" & cmbDoc.ItemData(cmbDoc.ListIndex) & "'"
    
    On Error Resume Next
    
    If rsStationary.RecordCount = 1 Then
        
        Dim NodeX As Node
        Dim NodX As Node
        
        Dim allchecked As Long
        
        For allchecked = 0 To ViSPMAP.Count - 1
        
            If ViSPMAP("r" & allchecked).RecIDb = Login.lVirtualID Then
                Set NodX = tvDocs.Nodes.Add(, , "v" & ViSPMAP("r" & allchecked).RecIDb, ViSPMAP("r" & allchecked).Desc, "visp")
            Else
                Set NodX = tvDocs.Nodes("v" & ViSPMAP("r" & allchecked).RecIDa)
                Set NodeX = tvDocs.Nodes.Add(NodX, tvwChild, "v" & ViSPMAP("r" & allchecked).RecIDb, ViSPMAP("r" & allchecked).Desc, "visp")
            End If
        
        Next

        SQL = rsStationary!selectStatement
        
        Dim rsDocs As adodb.Recordset
        
        Call MySQL.OpenTable(directConn, rsDocs, , MySQL.virtualisp(SQL, rsStationary!PrimaryTable, True, Login.bMaster))
        
        On Error Resume Next
        
        If rsDocs.State = adStateOpen Then
            If rsDocs.RecordCount > 0 Then
                Do
ReDo:
                    Set NodX = tvDocs.Nodes("v" & rsDocs!VirtualID & "_" & Format(rsDocs!Created, "yyyy") & "_" & Format(rsDocs!Created, "mm"))
                    If Err.Number <> 0 Then
                        Err.Clear
                        Set NodX = tvDocs.Nodes("v" & rsDocs!VirtualID & "_" & Format(rsDocs!Created, "yyyy"))
                        If Err.Number <> 0 Then
                            Set NodX = tvDocs.Nodes("v" & rsDocs!VirtualID)
                            Set NodeX = tvDocs.Nodes.Add(NodX, tvwChild, "v" & rsDocs!VirtualID & "_" & Format(rsDocs!Created, "yyyy"), Format(rsDocs!Created, "yyyy"), "year")
                            Err.Clear
                            GoTo ReDo
                        Else
                            Set NodeX = tvDocs.Nodes.Add(NodX, tvwChild, "v" & rsDocs!VirtualID & "_" & Format(rsDocs!Created, "yyyy") & "_" & Format(rsDocs!Created, "mm"), Format(rsDocs!Created, "mmmm"), "month")
                            Err.Clear
                            GoTo ReDo
                        End If
                    End If
                    
                    Set NodeX = tvDocs.Nodes.Add(NodX, tvwChild, "r" & Val(rsDocs("" & rsStationary!FieldName)), MySQL.OP(rsDocs, rsStationary!DisplayFormat), "item")
                    
                    rsDocs.MoveNext
                Loop Until rsDocs.EOF Or Err.Number <> 0
                rsDocs.Close
            End If
        End If
        
    End If
    
End Sub

Private Sub Form_Load()

    
    Call MySQL.OpenTable(directConn, rsStationary, , "select * from stationary")
    
    If rsStationary.State = adStateOpen Then
        If rsStationary.RecordCount > 0 Then
            Do
                cmbDoc.AddItem rsStationary!DOCShortname
                cmbDoc.ItemData(cmbDoc.ListCount - 1) = rsStationary!RecID
                rsStationary.MoveNext
            Loop Until rsStationary.EOF Or Err.Number <> 0
        End If
        rsStationary.MoveFirst
    End If
    
End Sub

Private Sub Form_Resize()

    If Me.WindowState = vbMinimized Then Exit Sub
    
    If (4565 / 16005) * Me.ScaleWidth < 2000 Then Exit Sub
    If Me.ScaleHeight < 2200 Then Exit Sub
    
    frmDocs(0).Move 120, 120, (4565 / 11000) * Me.ScaleWidth, Me.ScaleHeight - 240
    frmDocs(1).Move 120, 240, frmDocs(0).Width - 240, frmDocs(1).Height
    cmbDoc.Move 120, 240, frmDocs(1).Width - 240
    cmbAction(0).Width = cmbDoc.Width / 2 - 480
    cmbAction(1).Width = cmbAction(0).Width
    cmbAction(1).Left = cmbDoc.Left + cmbDoc.Width - cmbAction(0).Width
    cmbAction(2).Width = cmbDoc.Width
    cmbResend.Move 120, frmDocs(0).Height - cmbResend.Height - 120, frmDocs(0).Width - 240
    frmDocs(2).Move frmDocs(1).Left, frmDocs(1).Top + frmDocs(1).Height + 60, frmDocs(0).Width - 240, cmbResend.Top - 120 - frmDocs(1).Top - frmDocs(1).Height
    tvDocs.Move 120, 240, frmDocs(2).Width - 240, frmDocs(2).Height - 360
    frmDocs(3).Move frmDocs(0).Left + frmDocs(0).Width + 240, 120, Me.ScaleWidth - frmDocs(0).Width - 480, Me.ScaleHeight - 240
    www.Move 240, 360, frmDocs(3).Width - 480, frmDocs(3).Height - 600
'    www.
End Sub

Private Sub tvDocs_NodeClick(ByVal Node As MSComctlLib.Node)

    If Left(Node.Key, 1) = "r" Then
        cmbAction(0).Enabled = True
        cmbAction(1).Enabled = True
        cmbAction(2).Enabled = True
    
        If IsNull(rsStationary!formcode) Then
            frmDocs(3).Caption = "No History for this document standard"
        Else
        
            Call MySQL.SetColumnHeaders(rsStationary!formcode, www, "", directConn)
            Dim rsload As adodb.Recordset
            
            Call MySQL.OpenTable(directConn, rsload, , rsStationary!HistorySQL & " " & Mid(Node.Key, 2))
            
            If rsload.State = adStateOpen Then
            
                Call MySQL.fillLV(directConn, rsload, www, False)
                
            End If
            
        End If
        
    Else
        
        cmbAction(0).Enabled = False
        cmbAction(1).Enabled = False
        cmbAction(2).Enabled = False
        
    End If
        
        
End Sub

