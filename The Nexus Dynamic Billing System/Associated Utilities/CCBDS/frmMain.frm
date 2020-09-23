VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "C.C.B.D.S."
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8340
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrProcess 
      Interval        =   5000
      Left            =   7350
      Top             =   60
   End
   Begin VB.Timer tmrResfresh 
      Interval        =   500
      Left            =   7830
      Top             =   60
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   1500
      TabIndex        =   5
      Top             =   90
      Width           =   1395
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   90
      TabIndex        =   3
      Top             =   6570
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmbProcess 
      Caption         =   "Process"
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   90
      Width           =   1395
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   5385
      Left            =   150
      ScaleHeight     =   5385
      ScaleWidth      =   8025
      TabIndex        =   1
      Top             =   990
      Width           =   8025
      Begin MSComctlLib.ListView lvItems 
         Height          =   2415
         Left            =   120
         TabIndex        =   4
         Top             =   180
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   4260
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Account Name"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Amount Refunded"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Card Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Card Type"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin MSComctlLib.TabStrip tbs 
      Height          =   5925
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   10451
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Items"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oConn As New ADODB.Connection
Dim CC As New ADODB.Recordset
Dim InvOut As New ADODB.Recordset

Private Sub cmbProcess_Click()

    ProgressBar1.Value = 0
    
    tmrResfresh.Enabled = False
    cmdRefresh.Enabled = False
    
    Dim trans As CCTransaction
    Dim card As CreditCard
    Dim ret As CCSummaryStatusTypes
    Dim rsLoad As ADODB.Recordset
    Dim rsTraxr As ADODB.Recordset
    
    Dim InvOut2 As ADODB.Recordset
    Dim lx As Long
    Dim ly As Long
    Dim cAmount As Currency
    
    Dim bSkip As Boolean
    
    
    ret = 5
    
    Call MySQL.OpenTable(oConn, CC, , "Select Decode(CardNumber,'9b658a7e2ce494d53e79392ed7400f6873c25ae0c1769d5f9224c42918b1e02c68b301efe483d4ce753affacee934630e7cc3424568c934016c9f7c0f754c77beae2c761bbe0c18dd95d73c7e94d7c74d5d072540f69cdcae1ddec6f116ea65a9294a3b9e7a06578ac320812736fb357a88346c7d3c20df8ee796012330b6fc2eafa2804a87078afc643f8148dd8ec78') as CardNumber, RecID, SecurityNumber, ExpiryDate, Name, AccI_RecID, bDefault, bType " + _
                                                      "from CreditCard where ExpiryDate >= '" & Format(Now, "yyyy-mm-dd Hh:Nn:Ss") + "'")
    
    ProgressBar1.Value = 0
    ProgressBar1.Max = lvItems.ListItems.count
    For lx = 1 To lvItems.ListItems.count
        If lvItems.ListItems(lx).Checked = True Then
            cAmount = 0
            For ly = 1 To lvItems.ListItems.count
                If lvItems.ListItems(lx).Tag = lvItems.ListItems(ly).Tag Then
                    cAmount = cAmount + CCur(Mid(lvItems.ListItems(ly).SubItems(2), 2))
                End If
            Next
            
            If cAmount > 0 Then
                CC.Filter = "bDefault = -1 and AccI_RecID = " & lvItems.ListItems(lx).Tag
                
                Do While Not CC.EOF
                    Set trans = New CCTransaction
                    Set card = New CreditCard
                    
                    Set trans = New CCTransaction
                    Set card = New CreditCard
                    trans.Amount = "" & cAmount
                    trans.CCNumber = MySQL.NumDecrypt(CC!CardNumber)
                    trans.CCCode = CC!SecurityNumber
                    trans.CCExpiryMonth = Format(CC!ExpiryDate, "mm")
                    trans.CCExpiryYear = Format(CC!ExpiryDate, "yyyy")
                    trans.TransactionCode = MySQL.GetTMPRecID("cctransactions", oConn)
                    trans.RequesterIPAddress = "192.168.0.1"
                    
                    trans.CCName = CC!Name
                    trans.CCType = CC!bType
                    trans.CurrencyCode = "AUD"
                    ret = card.MakePayment(trans)
                    oConn.Execute "update cctransactions set Amount=" & trans.Amount & ",CCRecID=" & CC!RecID & ",CurrencyCode='" & trans.CurrencyCode & _
                                  "',ECIType=" & trans.ECIType & ",ErrorCode=" & trans.ErrorCode & ",ErrorMessage='" & trans.ErrorMessage & "'," & _
                                  "ID=" & trans.ID & ",OwnerID=" & trans.OwnerID & ",ProcessDate='" & Format(Now, "yyyy-mm-dd Hh:Nn:Ss") & "'," & _
                                  "ReceiptNumber='" & trans.ReceiptNumber & "',RequesterIPAddress='" & trans.RequesterIPAddress & "',ResponseText='" & _
                                  trans.ResponseText & "',SecureCode='" & trans.SecureCode & "',Status=" & trans.Status & ",SummaryCode=" & trans.SummaryCode & "," & _
                                  "TransactionCode='" & trans.TransactionCode & "',TransactionType=" & trans.TransactionType & " Where RecID = " & trans.TransactionCode
                    CC.MoveNext
                Loop
                
                If ret <> Approved Then
                    CC.Filter = "bDefault = 0 and AccI_RecID = " & lvItems.ListItems(lx).Tag
                    
                    Do While Not CC.EOF
                        lvItems.ListItems(lx).SubItems(5) = MySQL.NumDecrypt(CC!CardNumber)
                        Set trans = New CCTransaction
                        Set card = New CreditCard
                        
                        bSkip = False
                        
                        Select Case Val(CC!bType)
                        Case 0 ' EFTPOS
                            lvItems.ListItems(lx).SubItems(6) = "EFTPOS"
                            bSkip = True
                        Case 1 ' VISA
                            lvItems.ListItems(lx).SubItems(6) = "Visa"
                            bSkip = False
                        Case 2 ' MASTERCARD
                            lvItems.ListItems(lx).SubItems(6) = "Mastercard"
                            bSkip = False
                        Case 3 ' AMEX
                            lvItems.ListItems(lx).SubItems(6) = "AMEX"
                            bSkip = True
                        Case 4 ' Dinners
                            lvItems.ListItems(lx).SubItems(6) = "Dinner"
                            bSkip = True
                        Case 5 ' DiSCOVER
                            lvItems.ListItems(lx).SubItems(6) = "Discover"
                        Case 6 ' JCB
                            bSkip = True
                            lvItems.ListItems(lx).SubItems(6) = "JCB"
                        End Select
                        
                        If bSkip = True Then
                            lvItems.ListItems(lx).SubItems(6) = lvItems.ListItems(lx).SubItems(6) + " - " + "Skipped Unsupported"
                        Else
                        
                            Set trans = New CCTransaction
                            Set card = New CreditCard
                            
                            trans.Amount = "" & cAmount
                            trans.CCNumber = MySQL.NumDecrypt(CC!CardNumber)
                            trans.CCCode = CC!SecurityNumber
                            trans.CCExpiryMonth = Format(CC!ExpiryDate, "mm")
                            trans.CCExpiryYear = Format(CC!ExpiryDate, "yyyy")
                            trans.TransactionCode = MySQL.GetTMPRecID("cctransactions", oConn)
                            trans.RequesterIPAddress = "192.168.0.1"
                            trans.CCName = CC!Name
                            trans.CCType = CC!bType
                            trans.CurrencyCode = "AUD"
                            ret = card.MakePayment(trans)
    
                            
    
                            oConn.Execute "update cctransactions set Amount=" & trans.Amount & ",CCRecID=" & CC!RecID & ",CurrencyCode='" & trans.CurrencyCode & _
                                              "',ECIType=" & trans.ECIType & ",ErrorCode=" & trans.ErrorCode & ",ErrorMessage='" & trans.ErrorMessage & "'," & _
                                              "ID=" & trans.ID & ",OwnerID=" & trans.OwnerID & ",ProcessDate='" & Format(Now, "yyyy-mm-dd Hh:Nn:Ss") & "'," & _
                                              "ReceiptNumber='" & trans.ReceiptNumber & "',RequesterIPAddress='" & trans.RequesterIPAddress & "',ResponseText='" & _
                                              trans.ResponseText & "',SecureCode='" & trans.SecureCode & "',Status=" & trans.Status & ",SummaryCode=" & trans.SummaryCode & "," & _
                                              "TransactionCode='" & trans.TransactionCode & "',TransactionType=" & trans.TransactionType & " Where RecID = " & trans.TransactionCode
                        End If
                        
                        CC.MoveNext
                        If ret = Approved Then Exit Do
                    Loop
                End If
                
                If ret = Approved Then
                    For ly = 1 To lvItems.ListItems.count
                        If lvItems.ListItems(lx).Tag = lvItems.ListItems(ly).Tag Then
                            
                            
                            Select Case ret
                            Case FailedSystemError
                                lvItems.ListItems(ly).SubItems(4) = "System Error"
                            Case DeclinedBadCard
                                lvItems.ListItems(ly).SubItems(4) = "Declined Bad Card"
                            Case Declined
                                lvItems.ListItems(ly).SubItems(4) = "Declined"
                            Case Approved
                                lvItems.ListItems(ly).SubItems(4) = "Approved"
                                oConn.Execute "Update InvoiceOut Set AmountPaid=TotalDue - AmountRefunded - GSTRefunded Where RecID = " & Mid(lvItems.ListItems(ly).Key, 2)
                                                                    
                                bResult = MySQL.OpenTable(oConn, rsLoad, , "Select TraxrID from InvoiceOut Where RecID = " & Mid(lvItems.ListItems(ly).Key, 2))
                                If rsLoad!TraxrID <> 0 Then
                                    bResult = MySQL.OpenTable(oConn, rsTraxr, , "Select * from InvoiceTraxr Where RecID = " & rsLoad!TraxrID)
                                    If rsTraxr.RecordCount > 0 Then
                                        rsTraxr!AmountPaid = rsTraxr!AmountPaid + trans.Amount
                                        If rsTraxr!AmountPaid >= rsTraxr!TotalDue Then rsTraxr!Finalised = True
                                        rsTraxr.Update
                                    End If
                                End If
                            End Select
                            
                            Select Case ret
                            Case FailedSystemError
                                MySQL.AddReceiptItem oConn, lvItems.ListItems(ly).Tag, Val(Mid(lvItems.ListItems(ly).Key, 2)), , , , CCur(Mid(lvItems.ListItems(ly).SubItems(2), 2)), , "cc: " + "System Error " + Right(MySQL.NumDecrypt(trans.CCNumber), 5)
                            Case DeclinedBadCard
                                MySQL.AddReceiptItem oConn, lvItems.ListItems(ly).Tag, Val(Mid(lvItems.ListItems(ly).Key, 2)), , , , CCur(Mid(lvItems.ListItems(ly).SubItems(2), 2)), , "cc: " + "Declined Bad Card " + Right(MySQL.NumDecrypt(trans.CCNumber), 5)
                            Case Declined
                                MySQL.AddReceiptItem oConn, lvItems.ListItems(ly).Tag, Val(Mid(lvItems.ListItems(ly).Key, 2)), , , , CCur(Mid(lvItems.ListItems(ly).SubItems(2), 2)), , "cc: " + "Declined " + Right(MySQL.NumDecrypt(trans.CCNumber), 5)
                            Case Approved
                                MySQL.AddReceiptItem oConn, lvItems.ListItems(ly).Tag, Val(Mid(lvItems.ListItems(ly).Key, 2)), , , , CCur(Mid(lvItems.ListItems(ly).SubItems(2), 2)), , "cc: " + "Approved " + Right(MySQL.NumDecrypt(trans.CCNumber), 5)
                            End Select

                        End If
                    Next
                Else
                
                    For ly = 1 To lvItems.ListItems.count
                        If lvItems.ListItems(lx).Tag = lvItems.ListItems(ly).Tag Then
                            
                            CC.Filter = "bDefault = -1 and AccI_RecID = " & lvItems.ListItems(lx).Tag
                            Do While Not CC.EOF
                                Set trans = New CCTransaction
                                Set card = New CreditCard
                                lvItems.ListItems(ly).SubItems(5) = MySQL.NumDecrypt(CC!CardNumber)
                                
                                bSkip = False
                                
                                Select Case Val(CC!bType)
                                Case 0 ' EFTPOS
                                    lvItems.ListItems(ly).SubItems(6) = "EFTPOS"
                                    bSkip = True
                                Case 1 ' VISA
                                    lvItems.ListItems(ly).SubItems(6) = "Visa"
                                    bSkip = False
                                Case 2 ' MASTERCARD
                                    lvItems.ListItems(ly).SubItems(6) = "Mastercard"
                                    bSkip = False
                                Case 3 ' AMEX
                                    lvItems.ListItems(ly).SubItems(6) = "AMEX"
                                    bSkip = True
                                Case 4 ' Dinners
                                    lvItems.ListItems(ly).SubItems(6) = "Dinner"
                                    bSkip = True
                                Case 5 ' DiSCOVER
                                    lvItems.ListItems(ly).SubItems(6) = "Discover"
                                Case 6 ' JCB
                                    bSkip = True
                                    lvItems.ListItems(ly).SubItems(6) = "JCB"
                                End Select
                                lvItems.Refresh
                                
                                If bSkip = True Then
                                    lvItems.ListItems(ly).SubItems(6) = lvItems.ListItems(lx).SubItems(6) + " - " + "Skipped Unsupported"
                                Else
                                
                                
                                
                                     Set trans = New CCTransaction
                                     Set card = New CreditCard
                                     
                                     trans.Amount = Mid(lvItems.ListItems(ly).SubItems(2), 2)
                                     trans.CCNumber = MySQL.NumDecrypt(CC!CardNumber)
                                     trans.CCCode = CC!SecurityNumber
                                     trans.CCExpiryMonth = Format(CC!ExpiryDate, "mm")
                                     trans.CCExpiryYear = Format(CC!ExpiryDate, "yyyy")
                                     trans.TransactionCode = MySQL.GetTMPRecID("cctransactions", oConn)
                                     trans.RequesterIPAddress = "192.168.0.1"
                                     trans.CCName = CC!Name
                                     trans.CCType = CC!bType
                                     trans.CurrencyCode = "AUD"
                                     ret = card.MakePayment(trans)
                                     oConn.Execute "update cctransactions set Amount=" & trans.Amount & ",CCRecID=" & CC!RecID & ",CurrencyCode='" & trans.CurrencyCode & _
                                                   "',ECIType=" & trans.ECIType & ",ErrorCode=" & trans.ErrorCode & ",ErrorMessage='" & trans.ErrorMessage & "'," & _
                                                   "ID=" & trans.ID & ",OwnerID=" & trans.OwnerID & ",ProcessDate='" & Format(Now, "yyyy-mm-dd Hh:Nn:Ss") & "'," & _
                                                   "ReceiptNumber='" & trans.ReceiptNumber & "',RequesterIPAddress='" & trans.RequesterIPAddress & "',ResponseText='" & _
                                                   trans.ResponseText & "',SecureCode='" & trans.SecureCode & "',Status=" & trans.Status & ",SummaryCode=" & trans.SummaryCode & "," & _
                                                   "TransactionCode='" & trans.TransactionCode & "',TransactionType=" & trans.TransactionType & " Where RecID = " & trans.TransactionCode
                                    
                                     
                                     Select Case ret
                                     Case FailedSystemError
                                         MySQL.AddReceiptItem oConn, lvItems.ListItems(lx).Tag, Val(Mid(lvItems.ListItems(lx).Key, 2)), , , , CCur(Mid(lvItems.ListItems(ly).SubItems(2), 2)), , "cc: " + "System Error " + Right(MySQL.NumDecrypt(trans.CCNumber), 5)
                                     Case DeclinedBadCard
                                         MySQL.AddReceiptItem oConn, lvItems.ListItems(lx).Tag, Val(Mid(lvItems.ListItems(lx).Key, 2)), , , , CCur(Mid(lvItems.ListItems(ly).SubItems(2), 2)), , "cc: " + "Declined Bad Card " + Right(MySQL.NumDecrypt(trans.CCNumber), 5)
                                     Case Declined
                                         MySQL.AddReceiptItem oConn, lvItems.ListItems(lx).Tag, Val(Mid(lvItems.ListItems(lx).Key, 2)), , , , CCur(Mid(lvItems.ListItems(ly).SubItems(2), 2)), , "cc: " + "Declined " + Right(MySQL.NumDecrypt(trans.CCNumber), 5)
                                     Case Approved
                                         MySQL.AddReceiptItem oConn, lvItems.ListItems(lx).Tag, Val(Mid(lvItems.ListItems(lx).Key, 2)), , , , CCur(Mid(lvItems.ListItems(ly).SubItems(2), 2)), , "cc: " + "Approved " + Right(MySQL.NumDecrypt(trans.CCNumber), 5)
                                         oConn.Execute "Update InvoiceOut Set AmountPaid=TotalDue - AmountRefunded - GSTRefunded Where RecID = " & Mid(lvItems.ListItems(ly).Key, 2)
                                                                             
                                         bResult = MySQL.OpenTable(oConn, rsLoad, , "Select TraxrID from InvoiceOut Where RecID = " & Mid(lvItems.ListItems(ly).Key, 2))
                                         
                                         If rsLoad!TraxrID <> 0 Then
                                         
                                             bResult = MySQL.OpenTable(oConn, rsTraxr, , "Select * from InvoiceTraxr Where RecID = " & rsLoad!TraxrID)
                                             If rsTraxr.RecordCount > 0 Then
                                                 rsTraxr!AmountPaid = rsTraxr!AmountPaid + trans.Amount
                                                 If rsTraxr!AmountPaid >= rsTraxr!TotalDue Then rsTraxr!Finalised = True
                                                 rsTraxr.Update
                                             End If
                                         End If
                                     End Select
                                End If
                                CC.MoveNext
                            
                            Loop
                            
                            If ret <> Approved Then
                                CC.Filter = "bDefault = 0 and AccI_RecID = " & lvItems.ListItems(lx).Tag
                                Do While Not CC.EOF
                                    
                                    lvItems.ListItems(ly).SubItems(5) = MySQL.NumDecrypt(CC!CardNumber)
                                    
                                    bSkip = False
                                    
                                    Select Case Val(CC!bType)
                                    Case 0 ' EFTPOS
                                        lvItems.ListItems(ly).SubItems(6) = "EFTPOS"
                                        bSkip = True
                                    Case 1 ' VISA
                                        lvItems.ListItems(ly).SubItems(6) = "Visa"
                                        bSkip = False
                                    Case 2 ' MASTERCARD
                                        lvItems.ListItems(ly).SubItems(6) = "Mastercard"
                                        bSkip = False
                                    Case 3 ' AMEX
                                        lvItems.ListItems(ly).SubItems(6) = "AMEX"
                                        bSkip = True
                                    Case 4 ' Dinners
                                        lvItems.ListItems(ly).SubItems(6) = "Dinner"
                                        bSkip = True
                                    Case 5 ' DiSCOVER
                                        lvItems.ListItems(ly).SubItems(6) = "Discover"
                                    Case 6 ' JCB
                                        bSkip = True
                                        lvItems.ListItems(ly).SubItems(6) = "JCB"
                                    End Select
                                    lvItems.Refresh
                                    
                                    If bSkip = True Then
                                        lvItems.ListItems(ly).SubItems(6) = lvItems.ListItems(lx).SubItems(6) + " - " + "Skipped Unsupported"
                                    Else
                                    
                                    
                                        Set trans = New CCTransaction
                                        Set card = New CreditCard
                                        
                                        trans.Amount = Mid(lvItems.ListItems(ly).SubItems(2), 2)
                                        trans.CCNumber = MySQL.NumDecrypt(CC!CardNumber)
                                        'MsgBox MySQL.NumDecrypt(CC!CardNumber) & vbCrLf & MySQL.NumCrypt("1234567890") & vbCrLf & MySQL.NumDecrypt(MySQL.NumCrypt("1234567890"))
                                        trans.CCCode = CC!SecurityNumber
                                        trans.CCExpiryMonth = Format(CC!ExpiryDate, "mm")
                                        trans.CCExpiryYear = Format(CC!ExpiryDate, "yyyy")
                                        trans.TransactionCode = MySQL.GetTMPRecID("cctransactions", oConn)
                                        trans.RequesterIPAddress = "192.168.0.1"
                                        trans.CCName = CC!Name
                                        trans.CCType = CC!bType
                                        trans.CurrencyCode = "AUD"
                                        ret = card.MakePayment(trans)
                                        oConn.Execute "update cctransactions set Amount=" & trans.Amount & ",CCRecID=" & CC!RecID & ",CurrencyCode='" & trans.CurrencyCode & _
                                                      "',ECIType=" & trans.ECIType & ",ErrorCode=" & trans.ErrorCode & ",ErrorMessage='" & trans.ErrorMessage & "'," & _
                                                      "ID=" & trans.ID & ",OwnerID=" & trans.OwnerID & ",ProcessDate='" & Format(Now, "yyyy-mm-dd Hh:Nn:Ss") & "'," & _
                                                      "ReceiptNumber='" & trans.ReceiptNumber & "',RequesterIPAddress='" & trans.RequesterIPAddress & "',ResponseText='" & _
                                                      trans.ResponseText & "',SecureCode='" & trans.SecureCode & "',Status=" & trans.Status & ",SummaryCode=" & trans.SummaryCode & "," & _
                                                      "TransactionCode='" & trans.TransactionCode & "',TransactionType=" & trans.TransactionType & " Where RecID = " & trans.TransactionCode
                                                                       
                                        Select Case ret
                                        Case FailedSystemError
                                            MySQL.AddReceiptItem oConn, lvItems.ListItems(lx).Tag, Val(Mid(lvItems.ListItems(ly).Key, 2)), , , , CCur(Mid(lvItems.ListItems(ly).SubItems(2), 2)), , "cc: " + "System Error " + Right(MySQL.NumDecrypt(trans.CCNumber), 5)
                                        Case DeclinedBadCard
                                            MySQL.AddReceiptItem oConn, lvItems.ListItems(lx).Tag, Val(Mid(lvItems.ListItems(ly).Key, 2)), , , , CCur(Mid(lvItems.ListItems(ly).SubItems(2), 2)), , "cc: " + "Declined Bad Card " + Right(MySQL.NumDecrypt(trans.CCNumber), 5)
                                        Case Declined
                                            MySQL.AddReceiptItem oConn, lvItems.ListItems(lx).Tag, Val(Mid(lvItems.ListItems(ly).Key, 2)), , , , CCur(Mid(lvItems.ListItems(ly).SubItems(2), 2)), , "cc: " + "Declined " + Right(MySQL.NumDecrypt(trans.CCNumber), 5)
                                        Case Approved
                                            MySQL.AddReceiptItem oConn, lvItems.ListItems(lx).Tag, Val(Mid(lvItems.ListItems(ly).Key, 2)), , , , CCur(Mid(lvItems.ListItems(ly).SubItems(2), 2)), , "cc: " + "Approved " + Right(MySQL.NumDecrypt(trans.CCNumber), 5)
                                        End Select
                                        
                                    End If
                                    
                                    
                                    If CC.EOF = False Or ret <> Approved Then CC.MoveNext
                                    If ret = Approved Or CC.EOF Then Exit Do
                                Loop
                            End If
                            
                            If bSkip = False Then
                                Select Case ret
                                Case FailedSystemError
                                    lvItems.ListItems(ly).SubItems(4) = "System Error"
                                    MySQL.AddReceiptItem oConn, lvItems.ListItems(lx).Tag, , , , , , , "cc: System Error - " + Right(MySQL.NumDecrypt(trans.CCNumber), 5)
                                Case DeclinedBadCard
                                    lvItems.ListItems(ly).SubItems(4) = "Declined Bad Card"
                                    MySQL.AddReceiptItem oConn, lvItems.ListItems(lx).Tag, , , , , , , "cc: Declined Bad Card - " + Right(MySQL.NumDecrypt(trans.CCNumber), 5)
                                Case Declined
                                    lvItems.ListItems(ly).SubItems(4) = "Declined"
                                    MySQL.AddReceiptItem oConn, lvItems.ListItems(lx).Tag, , , , , , , "cc: Declined - " & Right(MySQL.NumDecrypt(trans.CCNumber), 5)
                                Case Approved
                                    MySQL.AddReceiptItem oConn, lvItems.ListItems(lx).Tag, , , , , , , "cc: Approved - " + Right(MySQL.NumDecrypt(trans.CCNumber), 5)
                                    lvItems.ListItems(ly).SubItems(4) = "Approved"
                                    oConn.Execute "Update InvoiceOut Set AmountPaid=TotalDue - AmountRefunded - GSTRefunded Where RecID = " & Mid(lvItems.ListItems(ly).Key, 2)
                                                                        
                                    bResult = MySQL.OpenTable(oConn, rsLoad, , "Select TraxrID from InvoiceOut Where RecID = " & Mid(lvItems.ListItems(ly).Key, 2))
                                    
                                    If rsLoad!TraxrID <> 0 Then
                                        
                                        bResult = MySQL.OpenTable(oConn, rsTraxr, , "Select * from InvoiceTraxr Where RecID = " & rsLoad!TraxrID)
                                        If rsTraxr.RecordCount > 0 Then
                                            rsTraxr!AmountPaid = rsTraxr!AmountPaid + trans.Amount
                                            If rsTraxr!AmountPaid >= rsTraxr!TotalDue Then rsTraxr!Finalised = True
                                            rsTraxr.Update
                                        End If
                                    End If
                                End Select
                            End If
                        End If
                    Next
                
                End If
                
            End If
        End If
        ProgressBar1.Value = lx
    
    Next
    
    tmrResfresh.Enabled = True
    cmdRefresh.Enabled = True
    
    
    Exit Sub

End Sub


Private Sub cmdRefresh_Click()

       Me.Caption = "Opening tables"
        Call MySQL.OpenTable(oConn, CC, , "Select Decode(CardNumber,'9b658a7e2ce494d53e79392ed7400f6873c25ae0c1769d5f9224c42918b1e02c68b301efe483d4ce753affacee934630e7cc3424568c934016c9f7c0f754c77beae2c761bbe0c18dd95d73c7e94d7c74d5d072540f69cdcae1ddec6f116ea65a9294a3b9e7a06578ac320812736fb357a88346c7d3c20df8ee796012330b6fc2eafa2804a87078afc643f8148dd8ec78') as CardNumber, RecID, SecurityNumber, ExpiryDate, Name, AccI_RecID, bDefault, bType " + _
                                                          "from CreditCard where ExpiryDate >= '" & Format(Now, "yyyy-mm-dd Hh:Nn:Ss") + "'")
        Call MySQL.OpenTable(oConn, InvOut, , "Select Count(*) as RecordCount from InvoiceOut, AccountInfo Where AccountInfo.RecID = InvoiceOut.AccI_RecID AND (InvoiceOut.AmountRefunded + InvoiceOut.GSTRefunded + InvoiceOut.AmountPaid) < InvoiceOut.TotalDue Order By AccountInfo.AccountName Limit 1")
        
        Me.Caption = "CCBDS"
        
        Me.Show
        DoEvents
        
        iRecCount = IIf(IsNull(InvOut!RecordCount), 0, InvOut!RecordCount)
        
        ProgressBar1.Value = 0
        
        If iRecCount <> 0 Then
            ProgressBar1.Max = iRecCount
            lvItems.ListItems.Clear
            For iX = 0 To iRecCount Step 30
                bResult = MySQL.OpenTable(oConn, InvOut, , "Select AccountInfo.AccountName, InvoiceOut.* from InvoiceOut, AccountInfo Where AccountInfo.RecID = InvoiceOut.AccI_RecID AND (InvoiceOut.AmountRefunded + InvoiceOut.GSTRefunded + InvoiceOut.AmountPaid) < InvoiceOut.TotalDue Order By AccountInfo.AccountName Limit " & iX & ",30")
                Select Case InvOut.RecordCount
                Case 0, -1
                Case Else
                    InvOut.MoveFirst
                    While Not InvOut.EOF
                        ProgressBar1.Value = ProgressBar1.Value + 1
                        CC.Filter = "AccI_RecID = " & InvOut!AccI_RecID
                        If CC.RecordCount > 0 Then
                            DoEvents
                            Set itmX = lvItems.ListItems.Add(, "r" & InvOut!RecID, InvOut!AccountName)
                            'itmX.SubItems(1) = Format(IIf(IsNull(InvOut!AmountDue), 0, InvOut!AmountDue), "Currency")
                            'itmX.SubItems(2) = Format(IIf(IsNull(InvOut!GSTCharged), 0, InvOut!GSTCharged), "Currency")
                            itmX.SubItems(2) = Format(IIf(IsNull(InvOut!TotalDue), 0, InvOut!TotalDue - InvOut!AmountRefunded - InvOut!GSTRefunded - InvOut!AmountPaid), "Currency")
                            'itmX.SubItems(4) = Format(IIf(IsNull(InvOut!PaymentDue), #9/19/1950#, InvOut!PaymentDue), "ddddd ttttt")
                            'itmX.SubItems(5) = Format(IIf(IsNull(InvOut!AmountPaid), 0, InvOut!AmountPaid), "Currency")
                            itmX.SubItems(1) = IIf(IsNull(InvOut!Description), "", InvOut!Description)
                            itmX.SubItems(3) = Format(IIf(IsNull(InvOut!AmountRefunded), 0, InvOut!AmountRefunded), "Currency")
                            itmX.Checked = True
                            itmX.Tag = InvOut!AccI_RecID
                        End If
                        InvOut.MoveNext
                    Wend
                End Select
            Next iX
        End If
        
End Sub

Private Sub Form_Load()

    Me.Visible = True
    Me.Caption = "Connecting to MySQL Server"
    
    GUI.LoadColWidths lvItems, Me
    
    If InStr(LCase(Command), "/test") = 0 Then
        Call MySQL.Connection(, , , , oConn)
    Else
        Call MySQL.Connection("no1_billing", "192.168.0.193", "sroberts", "process", oConn)
    End If
    
    Me.Caption = "Opening tables"
    Call MySQL.OpenTable(oConn, CC, , "Select Decode(CardNumber,'9b658a7e2ce494d53e79392ed7400f6873c25ae0c1769d5f9224c42918b1e02c68b301efe483d4ce753affacee934630e7cc3424568c934016c9f7c0f754c77beae2c761bbe0c18dd95d73c7e94d7c74d5d072540f69cdcae1ddec6f116ea65a9294a3b9e7a06578ac320812736fb357a88346c7d3c20df8ee796012330b6fc2eafa2804a87078afc643f8148dd8ec78') as CardNumber, RecID, SecurityNumber, ExpiryDate, Name, AccI_RecID, bDefault, bType " + _
                                                      "from CreditCard where ExpiryDate >= '" & Format(Now, "yyyy-mm-dd Hh:Nn:Ss") + "'")
    Call MySQL.OpenTable(oConn, InvOut, , "Select Count(*) as RecordCount from InvoiceOut, AccountInfo Where AccountInfo.RecID = InvoiceOut.AccI_RecID AND (InvoiceOut.AmountRefunded + InvoiceOut.GSTRefunded + InvoiceOut.AmountPaid) < InvoiceOut.TotalDue Order By AccountInfo.AccountName Limit 1")
    
    Me.Caption = "CCBDS"
    
    Me.Show
    DoEvents
    
    iRecCount = IIf(IsNull(InvOut!RecordCount), 0, InvOut!RecordCount)
    
    ProgressBar1.Value = 0
    
    If iRecCount <> 0 Then
        ProgressBar1.Max = iRecCount
        lvItems.ListItems.Clear
        For iX = 0 To iRecCount Step 30
            bResult = MySQL.OpenTable(oConn, InvOut, , "Select AccountInfo.AccountName, InvoiceOut.* from InvoiceOut, AccountInfo Where AccountInfo.RecID = InvoiceOut.AccI_RecID AND (InvoiceOut.AmountRefunded + InvoiceOut.GSTRefunded + InvoiceOut.AmountPaid) < InvoiceOut.TotalDue Order By AccountInfo.AccountName Limit " & iX & ",30")
            Select Case InvOut.RecordCount
            Case 0, -1
            Case Else
                InvOut.MoveFirst
                While Not InvOut.EOF
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    CC.Filter = "AccI_RecID = " & InvOut!AccI_RecID
                    If CC.RecordCount > 0 Then
                        DoEvents
                        Set itmX = lvItems.ListItems.Add(, "r" & InvOut!RecID, IIf(IsNull(InvOut!AccountName), "Account Name not Set", InvOut!AccountName))
                        'itmX.SubItems(1) = Format(IIf(IsNull(InvOut!AmountDue), 0, InvOut!AmountDue), "Currency")
                        'itmX.SubItems(2) = Format(IIf(IsNull(InvOut!GSTCharged), 0, InvOut!GSTCharged), "Currency")
                        itmX.SubItems(2) = Format(IIf(IsNull(InvOut!TotalDue), 0, InvOut!TotalDue - InvOut!AmountRefunded - InvOut!GSTRefunded - InvOut!AmountPaid), "Currency")
                        'itmX.SubItems(4) = Format(IIf(IsNull(InvOut!PaymentDue), #9/19/1950#, InvOut!PaymentDue), "ddddd ttttt")
                        'itmX.SubItems(5) = Format(IIf(IsNull(InvOut!AmountPaid), 0, InvOut!AmountPaid), "Currency")
                        itmX.SubItems(1) = IIf(IsNull(InvOut!Description), "", InvOut!Description)
                        itmX.SubItems(3) = Format(IIf(IsNull(InvOut!AmountRefunded), 0, InvOut!AmountRefunded), "Currency")
                        itmX.Checked = True
                        itmX.Tag = InvOut!AccI_RecID
                    End If
                    InvOut.MoveNext
                Wend
            End Select
        Next iX
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    GUI.SaveColWidths lvItems, Me
    
End Sub

Private Sub Form_Resize()

    If Me.WindowState = vbMinimized Then Exit Sub
    
    tbs.Move tbs.Left, tbs.Top, Me.ScaleWidth - tbs.Left * 2, Me.ScaleHeight - ProgressBar1.Height - 180 - tbs.Top
    pic.Move tbs.ClientLeft, tbs.ClientTop, tbs.ClientWidth, tbs.ClientHeight
    ProgressBar1.Move 60, Me.ScaleHeight - ProgressBar1.Height - 60, Me.ScaleWidth - 120
    
End Sub

Private Sub pic_Resize()

    lvItems.Move 60, 60, pic.ScaleWidth - 120, pic.ScaleHeight - 120
    
End Sub

Private Sub tmrProcess_Timer()

    Const iTotalMin = 120
    
    Static lLoop As Long
    Static StartTime As Date
    
    
    If lLoop = 0 Then
        lLoop = iTotalMin
        StartTime = Now
    ElseIf lLoop = 1 Then
        lLoop = lLoop - 1
        cmbProcess_Click
    Else
        lLoop = lLoop - 1
        'lvItems.Refresh
        cmbProcess.Caption = "&Process [" & (iTotalMin / 2) + DateDiff("n", Now, StartTime) & "]m"
        cmbProcess.Refresh
        DoEvents
    End If
    
End Sub

Private Sub tmrResfresh_Timer()

    Const imax = 120
    
    Static count As Long
    Static ResetTime As Date
    
    If count = 0 Then
        count = imax
        ResetTime = Now
    ElseIf count = 1 Then
        count = count - 1
        Call cmdRefresh_Click
    Else
        count = count - 1
        'lvItems.Refresh
        cmdRefresh.Caption = "&Refresh [" & (imax / 2) + DateDiff("s", Now, ResetTime) & "]s"
        cmdRefresh.Refresh
        DoEvents
    End If
            
End Sub
