VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colReseller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"clsReseller"
Attribute VB_Ext_KEY = "Member0" ,"clsReseller"

'local variable to hold collection
Private mCol As Collection
'local variable(s) to hold property value(s)
Private mvarCurrentReseller As clsReseller 'local copy
Private mvarCurrentVirtualID As Long 'local copy
'local variable(s) to hold property value(s)
'local variable(s) to hold property value(s)
Private mvarFetchStatus As enumFetchStatus 'local copy
Public Property Let FetchStatus(ByVal vData As enumFetchStatus)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FetchStatus = 5
    mvarFetchStatus = vData
End Property


Public Property Get FetchStatus() As enumFetchStatus
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FetchStatus
    FetchStatus = mvarFetchStatus
End Property









Public Property Let CurrentVirtualID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CurrentVirtualID = 5
    mvarCurrentVirtualID = vData
End Property


Public Property Get CurrentVirtualID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CurrentVirtualID
    CurrentVirtualID = mvarCurrentVirtualID
End Property



Public Property Set CurrentReseller(ByVal vData As clsReseller)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.CurrentReseller = Form1
    Set mvarCurrentReseller = vData
    mvarCurrentVirtualID = vData.RecID
    
End Property

Public Function Clear(Optional SESSIONKey As String = "")

    Dim ix As Long
    
    If Me.Count > 0 Then
        For ix = Me.Count To 1 Step -1
            If Len(SESSION) > 0 Then
                If Me(ix).SESSION = SESSIONKey Then
                    Me.Remove ix
                End If
            Else
                Me.Remove ix
            End If
        Next
    End If
    
End Function
Public Property Get CurrentReseller() As clsReseller
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CurrentReseller
    Set CurrentReseller = mvarCurrentReseller
    Me.CurrentVirtualID = mvarCurrentReseller.RecID
    
End Property

Public Function IsKeyInstance(sKey As String) As Long

    Static LastsKey As String
    Static Intialised As Boolean
    Static LastIndex As Long
    
    IsKeyInstance = -1
    
    If Me.Count > 0 Then
        If Not sKey = LastsKey Or Not Initialised = True Then
            Dim lk As Long
            For lk = 1 To Me.Count
                If Me(lk).Key = sKey Then
                    LastIndex = lk
                    IsKeyInstance = lk
                    Initialised = True
                    Exit For
                End If
            Next
        Else
            IsKeyInstance = LastIndex
        End If
    Else
        IsKeyInstance = -1
    End If

End Function


Public Function Add(Key As String, RecID As Long, AgencyID As Long, VirtualID As Long, Description As String, BriefDesc As Variant, Realm As String, CreationDate As Date, ABN As String, Subscribed As Long, SysopID As Long, ACN As String, NoSub As Long, CreatedBy_SysopID As Long, NextCycle As Date, PreviousCycle As Date, JoiningFee As Single, LogoURL As String, Icon As Integer, Manager As Variant, Manager_SysopID As Long, _
                    AssistanceManager_SysopID As Long, Comment As Variant, MISCFee As Single, bTaxMode As Integer, cTaxCode As String, cTaxCountry As String, cTaxExemptCode As String, ftpHostName As String, ftpUsername As String, _
                    ftpPassword As String, ftpPort As String, ftpBasePath As String, ftpTotalMB As Long, ftpCostPerMB As Single, ftpURLPath As String, FetchStatus As enumFetchStatus, NumSales_Minimum As Integer, NumSales_Maximum As Integer, NumSales_CappingAt As Integer, NumSales_AdminFeePerSale As Single, Cycle_IntervalType As String, Cycle_IntervalLength As Integer, SESSION As String, Optional sKey As String) As clsReseller
    
    If Me.IsKeyInstance(sKey) = -1 Then
        
        'create a new object
        Dim objNewMember As clsReseller
        Set objNewMember = New clsReseller
    
    
        'set the properties passed into the method
        objNewMember.Key = Key
        objNewMember.RecID = RecID
        objNewMember.AgencyID = AgencyID
        objNewMember.VirtualID = VirtualID
        objNewMember.Description = Description
        If IsObject(BriefDesc) Then
            Set objNewMember.BriefDesc = BriefDesc
        Else
            objNewMember.BriefDesc = BriefDesc
        End If
        objNewMember.Realm = Realm
        objNewMember.CreationDate = CreationDate
        objNewMember.ABN = ABN
        objNewMember.Subscribed = Subscribed
        objNewMember.SysopID = SysopID
        objNewMember.ACN = ACN
        objNewMember.NoSub = NoSub
        objNewMember.CreatedBy_SysopID = CreatedBy_SysopID
        objNewMember.NextCycle = NextCycle
        objNewMember.PreviousCycle = PreviousCycle
        objNewMember.JoiningFee = JoiningFee
        objNewMember.LogoURL = LogoURL
        objNewMember.Icon = Icon
        If IsObject(Manager) Then
            Set objNewMember.Manager = Manager
        Else
            objNewMember.Manager = Manager
        End If
        objNewMember.Manager_SysopID = Manager_SysopID
        objNewMember.AssistanceManager_SysopID = AssistanceManager_SysopID
        If IsObject(Comment) Then
            Set objNewMember.Comment = Comment
        Else
            objNewMember.Comment = Comment
        End If
        objNewMember.MISCFee = MISCFee
        objNewMember.bTaxMode = bTaxMode
        objNewMember.cTaxCode = cTaxCode
        objNewMember.cTaxCountry = cTaxCountry
        objNewMember.cTaxExemptCode = cTaxExemptCode
        objNewMember.ftpHostName = ftpHostName
        objNewMember.ftpUsername = ftpUsername
        objNewMember.ftpPassword = ftpPassword
        objNewMember.ftpPort = ftpPort
        objNewMember.ftpBasePath = ftpBasePath
        objNewMember.ftpTotalMB = ftpTotalMB
        objNewMember.ftpCostPerMB = ftpCostPerMB
        objNewMember.ftpURLPath = ftpURLPath
        Set objNewMember.colResellers_FoneNum = New colResellers_FoneNum
        Set objNewMember.colResellers_EmailAddy = New colResellers_EmailAddy
        Set objNewMember.colResellers_AccountHeader = New colResellers_AccountHeader
        Set objNewMember.colResellers_InvoiceItems = New colResellers_InvoiceItems
        Set objNewMember.colResellers_Payments = New colResellers_Payments
        Set objNewMember.colResellers_SnailMail = New colResellers_SnailMail
        
        objNewMember.IDX = Me.Count + 1
        objNewMember.SESSION = SESSION
        
        
        objNewMember.FetchStatus = Val(FetchStatus)
        
        
        objNewMember.NumSales_Minimum = NumSales_Minimum
        objNewMember.NumSales_Maximum = NumSales_Maximum
        objNewMember.NumSales_CappingAt = NumSales_CappingAt
        objNewMember.NumSales_AdminFeePerSale = NumSales_AdminFeePerSale
        objNewMember.Cycle_IntervalType = Cycle_IntervalType
        objNewMember.Cycle_IntervalLength = Cycle_IntervalLength
        
        If Len(sKey) = 0 Then
            mCol.Add objNewMember
        Else
            mCol.Add objNewMember, sKey
        End If
    
    
        'return the object created
        Set Add = objNewMember
        Set objNewMember = Nothing
    Else
        Set Add = Me(Me.IsKeyInstance(sKey))
    End If
       

End Function
Public Function CommitReseller(Optional ByVal VirtualID As Long) As Long
End Function

Public Function PopulateResellers(ByRef oConn As ADODB.Connection, ByRef ActionTyp As enumFetchStatus, ByVal VirtualID As Long, ByVal SESSION As String, Optional pb As ProgressBar, Optional ByVal pbset As Boolean) As Long

    Me.FetchStatus = fs_LoadingData
    
    Static tmpSession As String
    
    Dim SQL As String
    
    Dim rsKewl As New ADODB.Recordset
    Dim rsPhone As ADODB.Recordset
    Dim rsAddy As ADODB.Recordset
    Dim rseMail As ADODB.Recordset
    Dim rsHeader As ADODB.Recordset
    Dim rsInvoice As ADODB.Recordset
    
    If Len(tmpSession) = 0 And Len(SESSION) > 0 Then
        tmpSession = SESSION
    ElseIf Len(tmpSession) > 0 And Len(SESSION) = 0 Then
        SESSION = tmpSession
    End If
    
    If pbset = True Then
        pb.Value = 0
        pb.Max = 10
    End If
    
    If ActionTyp = fs_LoadHeader Or ActionTyp = fs_LoadAll Or ActionTyp = fs_LoadMinimum Then
            
        Me.Clear
        
        SQL = "select AES_DECRYPT(ftpPassword,'" + odb.colSalts.ReturnSalt("ftp") + "') as fPWD, ftpFileDBMode, ftpGroupingFolder, ftpIEProxy, ftpPingAlive, ftpNumberofFolders, ftpNumberofFiles, RecID, AgencyID, VirtualID, Description, BriefDesc, Realm, CreationDate, ABN, Subscribed, SysopID, ACN, NoSub, CreatedBy_SysopID, NextCycle, PreviousCycle, Cycle_IntervalType, Cycle_IntervalLength, JoiningFee, LogoURL, Icon, Manager, Manager_SysopID, AssistanceManager_SysopID, Comment, MISCFee, bTaxMode, cTaxCode, cTaxCountry, cTaxExemptCode , ftpHostName, ftpUsername, ftpPassword, ftpPort, ftpBasePath, ftpTotalMB, ftpCostPerMB, ftpURLPath, NumSales_AdminFeePerSale, NumSales_CappingAt, NumSales_Minimum, NumSales_Maximum from virtualisp where RecID = '" & VirtualID & "'"
        bResult = MySQL.OpenTable(oConn, rsKewl, , SQL, adOpenStatic, adLockOptimistic)
                
        If rsKewl.State = adStateOpen Then
            If rsKewl.RecordCount > 0 Then
                
                With Me.Add("RECID" & rsKewl!RecID, rsKewl!RecID, rsKewl!AgencyID, rsKewl!VirtualID, IIf(IsNull(rsKewl!Description), "", rsKewl!Description), IIf(IsNull(rsKewl!BriefDesc), "", rsKewl!BriefDesc), IIf(IsNull(rsKewl!Realm), "", rsKewl!Realm), _
                            rsKewl!CreationDate, rsKewl!ABN, rsKewl!Subscribed, rsKewl!SysopID, IIf(IsNull(rsKewl!ACN), "", rsKewl!ACN), rsKewl!NoSub, rsKewl!CreatedBy_SysopID, IIf(IsNull(rsKewl!NextCycle), DateAdd("d", 12, sysnow), rsKewl!NextCycle), IIf(IsNull(rsKewl!PreviousCycle), sysnow, rsKewl!PreviousCycle), rsKewl!JoiningFee, _
                            IIf(IsNull(rsKewl!LogoURL), "", rsKewl!LogoURL), rsKewl!Icon, IIf(IsNull(rsKewl!Manager), "", rsKewl!Manager), rsKewl!Manager_SysopID, rsKewl!AssistanceManager_SysopID, IIf(IsNull(rsKewl!Comment), "", rsKewl!Comment), _
                            rsKewl!MISCFee, IIf(IsNull(rsKewl!bTaxMode), "", rsKewl!bTaxMode), IIf(IsNull(rsKewl!cTaxCode), "", rsKewl!cTaxCode), IIf(IsNull(rsKewl!cTaxCountry), "", rsKewl!cTaxCountry), IIf(IsNull(rsKewl!cTaxExemptCode), "", rsKewl!cTaxExemptCode), _
                            IIf(IsNull(rsKewl!ftpHostName), "", rsKewl!ftpHostName), IIf(IsNull(rsKewl!ftpUsername), "", rsKewl!ftpUsername), IIf(IsNull(rsKewl!fPWD), "", rsKewl!fPWD), IIf(IsNull(rsKewl!ftpPort), "", rsKewl!ftpPort), IIf(IsNull(rsKewl!ftpBasePath), "", rsKewl!ftpBasePath), _
                            rsKewl!ftpTotalMB, rsKewl!ftpCostPerMB, IIf(IsNull(rsKewl!ftpURLPath), "", rsKewl!ftpURLPath), fs_NoChanges, rsKewl!NumSales_Minimum, rsKewl!NumSales_Maximum, rsKewl!NumSales_CappingAt, _
                            rsKewl!NumSales_AdminFeePerSale, IIf(IsNull(rsKewl!Cycle_IntervalType), "", rsKewl!Cycle_IntervalType), rsKewl!Cycle_IntervalLength, SESSION, "RECID" & rsKewl!RecID)
                            
                            .ftpFileDBMode = rsKewl!ftpFileDBMode
                            .ftpGroupingFolder = IIf(IsNull(rsKewl!ftpGroupingFolder), "temp", rsKewl!ftpGroupingFolder)
                            .ftpIEProxy = rsKewl!ftpIEProxy
                            .ftpPingAlive = rsKewl!ftpPingAlive
                            .ftpNumberofFolders = rsKewl!ftpNumberofFolders
                            .ftpNumberofFiles = rsKewl!ftpNumberofFiles
                End With
                
            End If
        End If
        
        If pbset = True Then
            If pb.Value + 2 >= pb.Max Then
                pb.Max = pb.Max + 10
            End If
            PopulateResellers = PopulateResellers + 1: pb.Value = pb.Value + 1
        End If
        
        SQL = "select AES_DECRYPT(ftpPassword,'" + odb.colSalts.ReturnSalt("ftp") + "') as fPWD, ftpFileDBMode, ftpGroupingFolder, ftpIEProxy, ftpPingAlive, ftpNumberofFolders, ftpNumberofFiles, RecID, AgencyID, VirtualID, Description, BriefDesc, Realm, CreationDate, ABN, Subscribed, SysopID, ACN, NoSub, CreatedBy_SysopID, NextCycle, PreviousCycle, Cycle_IntervalType, Cycle_IntervalLength, JoiningFee, LogoURL, Icon, Manager, Manager_SysopID, AssistanceManager_SysopID, Comment, MISCFee, bTaxMode, cTaxCode, cTaxCountry, cTaxExemptCode , ftpHostName, ftpUsername, ftpPassword, ftpPort, ftpBasePath, ftpTotalMB, ftpCostPerMB, ftpURLPath, NumSales_AdminFeePerSale, NumSales_CappingAt, NumSales_Minimum, NumSales_Maximum from virtualisp where VirtualID = '" & VirtualID & "'"
        bResult = MySQL.OpenTable(oConn, rsKewl, , SQL, adOpenStatic, adLockOptimistic)
    
        If rsKewl.State = adStateOpen Then
            If rsKewl.RecordCount > 0 Then
                If pbset = True Then pb.Max = pb.Max + rsKewl.RecordCount
                While Not rsKewl.EOF
                    
                    With Me.Add("RECID" & rsKewl!RecID, rsKewl!RecID, rsKewl!AgencyID, rsKewl!VirtualID, IIf(IsNull(rsKewl!Description), "", rsKewl!Description), IIf(IsNull(rsKewl!BriefDesc), "", rsKewl!BriefDesc), IIf(IsNull(rsKewl!Realm), "", rsKewl!Realm), _
                                rsKewl!CreationDate, rsKewl!ABN, rsKewl!Subscribed, rsKewl!SysopID, IIf(IsNull(rsKewl!ACN), "", rsKewl!ACN), rsKewl!NoSub, rsKewl!CreatedBy_SysopID, IIf(IsNull(rsKewl!NextCycle), DateAdd("d", 12, sysnow), rsKewl!NextCycle), IIf(IsNull(rsKewl!PreviousCycle), sysnow, rsKewl!PreviousCycle), rsKewl!JoiningFee, _
                                IIf(IsNull(rsKewl!LogoURL), "", rsKewl!LogoURL), rsKewl!Icon, IIf(IsNull(rsKewl!Manager), "", rsKewl!Manager), rsKewl!Manager_SysopID, rsKewl!AssistanceManager_SysopID, IIf(IsNull(rsKewl!Comment), "", rsKewl!Comment), _
                                rsKewl!MISCFee, IIf(IsNull(rsKewl!bTaxMode), "", rsKewl!bTaxMode), IIf(IsNull(rsKewl!cTaxCode), "", rsKewl!cTaxCode), IIf(IsNull(rsKewl!cTaxCountry), "", rsKewl!cTaxCountry), IIf(IsNull(rsKewl!cTaxExemptCode), "", rsKewl!cTaxExemptCode), _
                                IIf(IsNull(rsKewl!ftpHostName), "", rsKewl!ftpHostName), IIf(IsNull(rsKewl!ftpUsername), "", rsKewl!ftpUsername), IIf(IsNull(rsKewl!fPWD), "", rsKewl!fPWD), IIf(IsNull(rsKewl!ftpPort), "", rsKewl!ftpPort), IIf(IsNull(rsKewl!ftpBasePath), "", rsKewl!ftpBasePath), _
                                rsKewl!ftpTotalMB, rsKewl!ftpCostPerMB, IIf(IsNull(rsKewl!ftpURLPath), "", rsKewl!ftpURLPath), fs_NoChanges, rsKewl!NumSales_Minimum, rsKewl!NumSales_Maximum, rsKewl!NumSales_CappingAt, _
                                rsKewl!NumSales_AdminFeePerSale, IIf(IsNull(rsKewl!Cycle_IntervalType), "", rsKewl!Cycle_IntervalType), rsKewl!Cycle_IntervalLength, SESSION, "RECID" & rsKewl!RecID)
                                
                            .ftpFileDBMode = rsKewl!ftpFileDBMode
                            .ftpGroupingFolder = IIf(IsNull(rsKewl!ftpGroupingFolder), "temp", rsKewl!ftpGroupingFolder)
                            .ftpIEProxy = rsKewl!ftpIEProxy
                            .ftpPingAlive = rsKewl!ftpPingAlive
                            .ftpNumberofFolders = rsKewl!ftpNumberofFolders
                            .ftpNumberofFiles = rsKewl!ftpNumberofFiles
                    End With
                    
                    
                    rsKewl.MoveNext
                    
                    If pbset = True Then
                        If pb.Value + 2 >= pb.Max Then
                            pb.Max = pb.Max + 10
                        End If
                        PopulateResellers = PopulateResellers + 1: pb.Value = pb.Value + 1
                    End If
                    
                Wend
            End If
        End If
        
        
        If ActionTyp = fs_LoadAll Or ActionTyp = fs_LoadMinimum Then
            Dim lk As Long
            
            For lk = Me.Count To 1 Step -1
                'loads the records
            
                SQL = "select count(*) as RecCount from visp_phonenumbers where visp_RecID = '" & Me(lk).RecID & "'"
                bResult = MySQL.OpenTable(oConn, rsPhone, , SQL)
            
                If rsPhone.State = adStateOpen Then
                    If Me(Me.FindIDX(Me(lk).RecID)).colResellers_FoneNum.Count < rsPhone!RecCount Then
                        SQL = "select * from visp_phonenumbers where visp_RecID = '" & Me(lk).RecID & "'"
                        bResult = MySQL.OpenTable(oConn, rsPhone, , SQL)
                        Me(Me.FindIDX(Me(lk).RecID)).colResellers_FoneNum.Clear
                        If pbset = True Then pb.Max = pb.Max + rsPhone.RecordCount
                        While Not rsPhone.EOF
                            With Me(Me.FindIDX(Me(lk).RecID)).colResellers_FoneNum.Add("RECID" & rsPhone!RecID, rsPhone!RecID, rsPhone!visp_RecID, rsPhone!FlagID, rsPhone!DateAdded, IIf(IsNull(rsPhone!PhoneNumber), "", rsPhone!PhoneNumber), _
                                                                                  IIf(IsNull(rsPhone!Extension), "", rsPhone!Extension), IIf(IsNull(rsPhone!ContactName), "", rsPhone!ContactName), rsPhone!Cancelled, rsPhone!Checked, _
                                                                                  IIf(IsNull(rsPhone!PhotoURL), "", rsPhone!PhotoURL), fs_NoChanges, SESSION, "RECID" & rsPhone!RecID)
                                .ShortNote = IIf(IsNull(rsPhone!ShortNote), "", rsPhone!ShortNote)
                            End With
                            If pbset = True Then
                                If pb.Value + 2 >= pb.Max Then
                                    pb.Max = pb.Max + 10
                                End If
                                PopulateResellers = PopulateResellers + 1: pb.Value = pb.Value + 1
                            End If
                            rsPhone.MoveNext
                        Wend
                    End If
                End If
                rsPhone.Close
                
                
                SQL = "select count(*) as RecCount from visp_emailaddresses where visp_RecID = '" & Me(lk).RecID & "'"
                bResult = MySQL.OpenTable(oConn, rseMail, , SQL)
                
                If rseMail.State = adStateOpen Then
                    If Me(Me.FindIDX(Me(lk).RecID)).colResellers_EmailAddy.Count < rseMail!RecCount Then
                        SQL = "select * from visp_emailaddresses where visp_RecID = '" & Me(lk).RecID & "'"
                        bResult = MySQL.OpenTable(oConn, rseMail, , SQL)
                        Me(Me.FindIDX(Me(lk).RecID)).colResellers_EmailAddy.Clear
                        If pbset = True Then pb.Max = pb.Max + rseMail.RecordCount
                        While Not rseMail.EOF
                            With Me(Me.FindIDX(Me(lk).RecID)).colResellers_EmailAddy.Add("RECID" & rseMail!RecID, rseMail!RecID, rseMail!visp_RecID, rseMail!FlagID, rseMail!DateAdded, IIf(IsNull(rseMail!EmailAddress), "", rseMail!EmailAddress), _
                                                                                    IIf(IsNull(rseMail!ContactName), "", rseMail!ContactName), rseMail!Cancelled, rseMail!Checked, IIf(IsNull(rseMail!PhotoURL), "", rseMail!PhotoURL), fs_NoChanges, _
                                                                                    SESSION, "RECID" & rseMail!RecID)
                            End With
                            If pbset = True Then
                                If pb.Value + 2 >= pb.Max Then
                                    pb.Max = pb.Max + 10
                                End If
                                PopulateResellers = PopulateResellers + 1: pb.Value = pb.Value + 1
                            End If
                            rseMail.MoveNext
                        Wend
                    End If
                End If
                rseMail.Close
                
                
                SQL = "select count(*) as RecCount from visp_addresses where visp_RecID = '" & Me(lk).RecID & "'"
                bResult = MySQL.OpenTable(oConn, rsAddy, , SQL)
               
                If rsAddy.State = adStateOpen Then
                    If Me(Me.FindIDX(Me(lk).RecID)).colResellers_SnailMail.Count < rsAddy!RecCount Then
                        SQL = "select * from visp_addresses where visp_RecID = '" & Me(lk).RecID & "'"
                        bResult = MySQL.OpenTable(oConn, rsAddy, , SQL)
                        Me(Me.FindIDX(Me(lk).RecID)).colResellers_SnailMail.Clear
                        If pbset = True Then pb.Max = pb.Max + rsAddy.RecordCount
                        While Not rsAddy.EOF
                            With Me(Me.FindIDX(Me(lk).RecID)).colResellers_SnailMail.Add("RECID" & rsAddy!RecID, rsAddy!RecID, rsAddy!visp_RecID, rsAddy!FlagID, rsAddy!DateAdded, IIf(IsNull(rsAddy!ContactName), "", rsAddy!ContactName), _
                                                                                    IIf(IsNull(rsAddy!Street1), "", rsAddy!Street1), IIf(IsNull(rsAddy!Street2), "", rsAddy!Street2), IIf(IsNull(rsAddy!Country), "", rsAddy!Country), _
                                                                                    IIf(IsNull(rsAddy!State), "", rsAddy!State), IIf(IsNull(rsAddy!Suburb), "", rsAddy!Suburb), IIf(IsNull(rsAddy!PostCode), "", rsAddy!PostCode), _
                                                                                    rsAddy!Cancelled, rsAddy!Checked, IIf(IsNull(rsAddy!PhotoURL), "", rsAddy!PhotoURL), fs_NoChanges, SESSION, "RECID" & rsAddy!RecID)
                            End With
                            If pbset = True Then
                                If pb.Value + 2 >= pb.Max Then
                                    pb.Max = pb.Max + 10
                                End If
                                PopulateResellers = PopulateResellers + 1: pb.Value = pb.Value + 1
                            End If
                            rsAddy.MoveNext
                        Wend
                    End If
                End If
                rsAddy.Close
                
                If ActionTyp = fs_LoadAll Then
                    
                    bResult = MySQL.OpenTable(oConn, rsHeader, , "select * from visp_accountheader where VirtualID = '" & Me(lk).RecID & "'")
                    If rsHeader.State = adStateOpen Then
                        If Me(Me.FindIDX(Me(lk).RecID)).colResellers_AccountHeader.Count < rsHeader.RecordCount Then
                            Me(Me.FindIDX(Me(lk).RecID)).colResellers_AccountHeader.Clear
                            If pbset = True Then pb.Max = pb.Max + rsHeader.RecordCount
                            While Not rsPhone.EOF
                                With Me(Me.FindIDX(Me(lk).RecID)).colResellers_AccountHeader.Add("RECID" & rsHeader!RecID, rsHeader!RecID, rsHeader!Me(lk).RecID, rsHeader!InvoiceID, rsHeader!StartDate, rsHeader!EndDate, rsHeader!TotalIncome, _
                                                                                            rsHeader!TotalCost, rsHeader!Paid, IIf(IsNull(rsHeader!ChequeNumber), "", rsHeader!ChequeNumber), rsHeader!Margin, rsHeader!visp_PaymentID, rsHeader!Income_GrossPayments, rsHeader!Income_AssessableGovernment, _
                                                                                            rsHeader!Income_IndustryPayments, rsHeader!Income_OtherBusiness, rsHeader!Income_TotalOverall, rsHeader!Expenses_Superannuation, rsHeader!Expenses_Commission, _
                                                                                            rsHeader!Expenses_CostofSale, rsHeader!Expenses_BadDebts, rsHeader!Expenses_Lease, rsHeader!Expenses_Rent, rsHeader!Expenses_TotalInterest, rsHeader!Expenses_TotalRoyalty, _
                                                                                            rsHeader!Expenses_Depreciation, rsHeader!Expenses_MotorVechicle, rsHeader!Expenses_RepairsMaintenance, rsHeader!Expenses_AllOther, rsHeader!Expenses_TotalOverall, _
                                                                                            rsHeader!Adjustments_Expenses, rsHeader!Adjustments_ReconciliationItems, rsHeader!Adjustments_Income, rsHeader!Adjustments_DroughtAllowance, rsHeader!sum_NetFromBusine, _
                                                                                            rsHeader!FlagID, fs_NoChanges, SESSION, "RECID" & rsHeader!RecID)
                                End With
                                If pbset = True Then
                                    If pb.Value + 2 >= pb.Max Then
                                        pb.Max = pb.Max + 10
                                    End If
                                    PopulateResellers = PopulateResellers + 1: pb.Value = pb.Value + 1
                                End If
                                rsHeader.MoveNext
                            Wend
                        End If
                    End If
                    rsHeader.Close
                
                    bResult = MySQL.OpenTable(oConn, rsInvoice, , "select * from visp_invoiceitems where VirtualID = '" & Me(lk).RecID & "'")
                    If rsInvoice.State = adStateOpen Then
                        If Me(Me.FindIDX(Me(lk).RecID)).colResellers_InvoiceItems.Count < rsInvoice.RecordCount Then
                            Me(Me.FindIDX(Me(lk).RecID)).colResellers_InvoiceItems.Clear
                            If pbset = True Then pb.Max = pb.Max + rsInvoice.RecordCount
                            While Not rsInvoice.EOF
                                With Me(Me.FindIDX(Me(lk).RecID)).colResellers_InvoiceItems.Add("RECID" & rsInvoice!RecID, rsInvoice!RecID, rsInvoice!InvoiceID, rsInvoice!AmountDue, rsInvoice!CostPrice, rsInvoice!GSTCharged, rsInvoice!PaymentDue, rsInvoice!FinalClosureDate, _
                                                                                                rsInvoice!AmountPaid, rsInvoice!GSTPaid, rsInvoice!PaidWhen, rsInvoice!TotalDue, rsInvoice!TotalPaid, rsInvoice!VirtualID, rsInvoice!SysopID, rsInvoice!FlagID, IIf(IsNull(rsInvoice!Description), "", rsInvoice!Description), fs_NoChanges, SESSION, "RECID" & rsInvoice!RecID)
                                End With
                                If pbset = True Then
                                    If pb.Value + 2 >= pb.Max Then
                                        pb.Max = pb.Max + 10
                                    End If
                                    PopulateResellers = PopulateResellers + 1: pb.Value = pb.Value + 1
                                End If
                                rsInvoice.MoveNext
                            Wend
                        End If
                    End If
                    rsInvoice.Close
                
                
                End If
                
                gSleep
                
            Next
        Else
            Me.FetchStatus = fs_Idle
            If pbset = True Then pb.Value = 0: Exit Function
        End If
    End If
    
    
    
    ' Will load all the email address for a visp.
    If ActionTyp = fs_LoadPhone Or ActionTyp = fs_LoadAllContactDetails Then
        
        SQL = "select count(*) as RecCount from visp_phonenumbers where visp_RecID = '" & VirtualID & "'"
        bResult = MySQL.OpenTable(oConn, rsPhone, , SQL)
        
        If rsPhone.State = adStateOpen Then
            If Me(Me.FindIDX(VirtualID)).colResellers_FoneNum.Count < rsPhone!RecCount Then
                SQL = "select * from visp_phonenumbers where visp_RecID = '" & VirtualID & "'"
                bResult = MySQL.OpenTable(oConn, rsPhone, , SQL)
                Me(Me.FindIDX(VirtualID)).colResellers_FoneNum.Clear
                If pbset = True Then pb.Max = pb.Max + rsPhone.RecordCount
                While Not rsPhone.EOF
                    With Me(Me.FindIDX(VirtualID)).colResellers_FoneNum.Add("RECID" & rsPhone!RecID, rsPhone!RecID, rsPhone!visp_RecID, rsPhone!FlagID, rsPhone!DateAdded, IIf(IsNull(rsPhone!PhoneNumber), "", rsPhone!PhoneNumber), _
                                                                          IIf(IsNull(rsPhone!Extension), "", rsPhone!Extension), IIf(IsNull(rsPhone!ContactName), "", rsPhone!ContactName), rsPhone!Cancelled, rsPhone!Checked, _
                                                                          IIf(IsNull(rsPhone!PhotoURL), "", rsPhone!PhotoURL), fs_NoChanges, SESSION, "RECID" & rsPhone!RecID)
                        .ShortNote = IIf(IsNull(rsPhone!ShortNote), "", rsPhone!ShortNote)
                    End With
                    If pbset = True Then
                        If pb.Value + 2 >= pb.Max Then
                            pb.Max = pb.Max + 10
                        End If
                        PopulateResellers = PopulateResellers + 1: pb.Value = pb.Value + 1
                    End If
                    rsPhone.MoveNext
                Wend
            End If
        End If
        rsPhone.Close
    End If
    
    ' Will load the address information for a individual visp
    If ActionTyp = fs_LoadEmail Or ActionTyp = fs_LoadAllContactDetails Then
        
        SQL = "select count(*) as RecCount from visp_emailaddresses where visp_RecID = '" & VirtualID & "'"
        bResult = MySQL.OpenTable(oConn, rseMail, , SQL)
        
        If rseMail.State = adStateOpen Then
            If Me(Me.FindIDX(VirtualID)).colResellers_EmailAddy.Count < rseMail!RecCount Then
                SQL = "select * from visp_emailaddresses where visp_RecID = '" & VirtualID & "'"
                bResult = MySQL.OpenTable(oConn, rseMail, , SQL)
                Me(Me.FindIDX(VirtualID)).colResellers_EmailAddy.Clear
                If pbset = True Then pb.Max = pb.Max + rseMail.RecordCount
                While Not rseMail.EOF
                    With Me(Me.FindIDX(VirtualID)).colResellers_EmailAddy.Add("RECID" & rseMail!RecID, rseMail!RecID, rseMail!visp_RecID, rseMail!FlagID, rseMail!DateAdded, IIf(IsNull(rseMail!EmailAddress), "", rseMail!EmailAddress), _
                                                                            IIf(IsNull(rseMail!ContactName), "", rseMail!ContactName), rseMail!Cancelled, rseMail!Checked, IIf(IsNull(rseMail!PhotoURL), "", rseMail!PhotoURL), fs_NoChanges, _
                                                                            SESSION, "RECID" & rseMail!RecID)
                    End With
                    If pbset = True Then
                        If pb.Value + 2 >= pb.Max Then
                            pb.Max = pb.Max + 10
                        End If
                        PopulateResellers = PopulateResellers + 1: pb.Value = pb.Value + 1
                    End If
                    rseMail.MoveNext
                Wend
            End If
        End If
        rseMail.Close
    End If
    
    
    ' Will load the phone details of the Virtual ISP
    If ActionTyp = fs_LoadAddress Or ActionTyp = fs_LoadAllContactDetails Then
        SQL = "select count(*) as RecCount from visp_addresses where visp_RecID = '" & VirtualID & "'"
         bResult = MySQL.OpenTable(oConn, rsAddy, , SQL)
        
         If rsAddy.State = adStateOpen Then
             If Me(Me.FindIDX(VirtualID)).colResellers_SnailMail.Count < rsAddy!RecCount Then
                 SQL = "select * from visp_addresses where visp_RecID = '" & VirtualID & "'"
                 bResult = MySQL.OpenTable(oConn, rsAddy, , SQL)
                 Me(Me.FindIDX(VirtualID)).colResellers_SnailMail.Clear
                 If pbset = True Then pb.Max = pb.Max + rsAddy.RecordCount
                 While Not rsAddy.EOF
                     With Me(Me.FindIDX(VirtualID)).colResellers_SnailMail.Add("RECID" & rsAddy!RecID, rsAddy!RecID, rsAddy!visp_RecID, rsAddy!FlagID, rsAddy!DateAdded, IIf(IsNull(rsAddy!ContactName), "", rsAddy!ContactName), _
                                                                             IIf(IsNull(rsAddy!Street1), "", rsAddy!Street1), IIf(IsNull(rsAddy!Street2), "", rsAddy!Street2), IIf(IsNull(rsAddy!Country), "", rsAddy!Country), _
                                                                             IIf(IsNull(rsAddy!State), "", rsAddy!State), IIf(IsNull(rsAddy!Suburb), "", rsAddy!Suburb), IIf(IsNull(rsAddy!PostCode), "", rsAddy!PostCode), _
                                                                             rsAddy!Cancelled, rsAddy!Checked, IIf(IsNull(rsAddy!PhotoURL), "", rsAddy!PhotoURL), fs_NoChanges, SESSION, "RECID" & rsAddy!RecID)
                     End With
                     If pbset = True Then
                         If pb.Value + 2 >= pb.Max Then
                             pb.Max = pb.Max + 10
                         End If
                         PopulateResellers = PopulateResellers + 1: pb.Value = pb.Value + 1
                     End If
                     rsAddy.MoveNext
                 Wend
             End If
         End If
         rsAddy.Close
    End If
    
    
    ' Will load the invoice/tax header details of the Virtual ISP
    If ActionTyp = fs_LoadHeader Then
        bResult = MySQL.OpenTable(oConn, rsHeader, , "select * from visp_accountheader where VirtualID = '" & VirtualID & "'")
        If rsHeader.State = adStateOpen Then
            If Me(Me.FindIDX(VirtualID)).colResellers_AccountHeader.Count < rsHeader.RecordCount Then
                Me(Me.FindIDX(VirtualID)).colResellers_AccountHeader.Clear
                If pbset = True Then pb.Max = pb.Max + rsHeader.RecordCount
                While Not rsPhone.EOF
                    With Me(Me.FindIDX(VirtualID)).colResellers_AccountHeader.Add("RECID" & rsHeader!RecID, rsHeader!RecID, rsHeader!VirtualID, rsHeader!InvoiceID, rsHeader!StartDate, rsHeader!EndDate, rsHeader!TotalIncome, _
                                                                                rsHeader!TotalCost, rsHeader!Paid, IIf(IsNull(rsHeader!ChequeNumber), "", rsHeader!ChequeNumber), rsHeader!Margin, rsHeader!visp_PaymentID, rsHeader!Income_GrossPayments, rsHeader!Income_AssessableGovernment, _
                                                                                rsHeader!Income_IndustryPayments, rsHeader!Income_OtherBusiness, rsHeader!Income_TotalOverall, rsHeader!Expenses_Superannuation, rsHeader!Expenses_Commission, _
                                                                                rsHeader!Expenses_CostofSale, rsHeader!Expenses_BadDebts, rsHeader!Expenses_Lease, rsHeader!Expenses_Rent, rsHeader!Expenses_TotalInterest, rsHeader!Expenses_TotalRoyalty, _
                                                                                rsHeader!Expenses_Depreciation, rsHeader!Expenses_MotorVechicle, rsHeader!Expenses_RepairsMaintenance, rsHeader!Expenses_AllOther, rsHeader!Expenses_TotalOverall, _
                                                                                rsHeader!Adjustments_Expenses, rsHeader!Adjustments_ReconciliationItems, rsHeader!Adjustments_Income, rsHeader!Adjustments_DroughtAllowance, rsHeader!sum_NetFromBusine, _
                                                                                rsHeader!FlagID, fs_NoChanges, SESSION, "RECID" & rsHeader!RecID)
                    End With
                    rsHeader.MoveNext
                    If pbset = True Then
                        If pb.Value + 2 >= pb.Max Then
                            pb.Max = pb.Max + 10
                        End If
                        PopulateResellers = PopulateResellers + 1: pb.Value = pb.Value + 1
                    End If
                Wend
            End If
        End If
        rsHeader.Close
    End If
    
    
    
    ' Will load the invoice/tax header details of the Virtual ISP
    If ActionTyp = fs_LoadInvoice Then
        bResult = MySQL.OpenTable(oConn, rsInvoice, , "select * from visp_invoiceitems where VirtualID = '" & Me(lk).RecID & "'")
        If rsInvoice.State = adStateOpen Then
            If Me(Me.FindIDX(Me(lk).RecID)).colResellers_InvoiceItems.Count < rsInvoice.RecordCount Then
                Me(Me.FindIDX(Me(lk).RecID)).colResellers_InvoiceItems.Clear
                If pbset = True Then pb.Max = pb.Max + rsInvoice.RecordCount
                While Not rsInvoice.EOF
                    With Me(Me.FindIDX(Me(lk).RecID)).colResellers_InvoiceItems.Add("RECID" & rsInvoice!RecID, rsInvoice!RecID, rsInvoice!InvoiceID, rsInvoice!AmountDue, rsInvoice!CostPrice, rsInvoice!GSTCharged, rsInvoice!PaymentDue, rsInvoice!FinalClosureDate, _
                                                                                    rsInvoice!AmountPaid, rsInvoice!GSTPaid, rsInvoice!PaidWhen, rsInvoice!TotalDue, rsInvoice!TotalPaid, rsInvoice!VirtualID, rsInvoice!SysopID, rsInvoice!FlagID, IIf(IsNull(rsInvoice!Description), "", rsInvoice!Description), fs_NoChanges, SESSION, "RECID" & rsInvoice!RecID)
                    End With
                    If pbset = True Then
                        If pb.Value + 2 >= pb.Max Then
                            pb.Max = pb.Max + 10
                        End If
                        PopulateResellers = PopulateResellers + 1: pb.Value = pb.Value + 1
                    End If
                    rsInvoice.MoveNext
                Wend
            End If
        End If
        rsInvoice.Close
    End If
    
    Me.FetchStatus = fs_Idle
    If pbset = True Then pb.Value = 0: Exit Function

    
    
'    ' Loads all contact details, this is email, phone and addresses
'    If ActionTyp = fs_LoadAllContactDetails Then
'
'        If Me.Count > 0 Then
'
'            SQL = "select count(*) as RecCount from visp_phonenumbers where visp_RecID = '" & VirtualID & "'"
'            bResult = MySQL.OpenTable(oConn, rsPhone, , SQL)
'
'            If rsPhone.State = adStateOpen Then
'                If Me(Me.FindIDX(VirtualID)).colResellers_FoneNum.Count < rsPhone!RecCount Then
'                    SQL = "select * from visp_phonenumbers where visp_RecID = '" & VirtualID & "'"
'                    bResult = MySQL.OpenTable(oConn, rsPhone, , SQL)
'                    Me(Me.FindIDX(VirtualID)).colResellers_FoneNum.Clear
'                    If pbset = True Then pb.Max = pb.Max + rsPhone.RecordCount
'                    While Not rsPhone.EOF
'                        With Me(Me.FindIDX(VirtualID)).colResellers_FoneNum.Add("RECID" & rsPhone!RecID, rsPhone!RecID, rsPhone!visp_RecID, rsPhone!FlagID, rsPhone!DateAdded, IIf(IsNull(rsPhone!PhoneNumber), "", rsPhone!PhoneNumber), _
'                                                                              IIf(IsNull(rsPhone!Extension), "", rsPhone!Extension), IIf(IsNull(rsPhone!ContactName), "", rsPhone!ContactName), rsPhone!Cancelled, rsPhone!Checked, _
'                                                                              IIf(IsNull(rsPhone!PhotoURL), "", rsPhone!PhotoURL), fs_NoChanges, SESSION, "RECID" & rsPhone!RecID)
'                            .ShortNote = IIf(IsNull(rsPhone!ShortNote), "", rsPhone!ShortNote)
'                        End With
'                        If pbset = True Then
'                            If pb.Value + 2 >= pb.Max Then
'                                pb.Max = pb.Max + 10
'                            End If
'                            PopulateResellers = PopulateResellers + 1: pb.Value = pb.Value + 1
'                        End If
'                        rsPhone.MoveNext
'                    Wend
'                End If
'            End If
'            rsPhone.Close
'
'
'            SQL = "select count(*) as RecCount from visp_emailaddresses where visp_RecID = '" & VirtualID & "'"
'            bResult = MySQL.OpenTable(oConn, rsEmail, , SQL)
'
'            If rsEmail.State = adStateOpen Then
'                If Me(Me.FindIDX(VirtualID)).colResellers_EmailAddy.Count < rsEmail!RecCount Then
'                    SQL = "select * from visp_emailaddresses where visp_RecID = '" & VirtualID & "'"
'                    bResult = MySQL.OpenTable(oConn, rsEmail, , SQL)
'                    Me(Me.FindIDX(VirtualID)).colResellers_EmailAddy.Clear
'                    If pbset = True Then pb.Max = pb.Max + rsEmail.RecordCount
'                    While Not rsEmail.EOF
'                        With Me(Me.FindIDX(VirtualID)).colResellers_EmailAddy.Add("RECID" & rsEmail!RecID, rsEmail!RecID, rsEmail!visp_RecID, rsEmail!FlagID, rsEmail!DateAdded, IIf(IsNull(rsEmail!EmailAddress), "", rsEmail!EmailAddress), _
'                                                                                IIf(IsNull(rsEmail!ContactName), "", rsEmail!ContactName), rsEmail!Cancelled, rsEmail!Checked, IIf(IsNull(rsEmail!PhotoURL), "", rsEmail!PhotoURL), fs_NoChanges, _
'                                                                                SESSION, "RECID" & rsEmail!RecID)
'                        End With
'                        If pbset = True Then
'                            If pb.Value + 2 >= pb.Max Then
'                                pb.Max = pb.Max + 10
'                            End If
'                            PopulateResellers = PopulateResellers + 1: pb.Value = pb.Value + 1
'                        End If
'                        rsEmail.MoveNext
'                    Wend
'                End If
'            End If
'            rsEmail.Close
'
'
'            SQL = "select count(*) as RecCount from visp_addresses where visp_RecID = '" & VirtualID & "'"
'            bResult = MySQL.OpenTable(oConn, rsAddy, , SQL)
'
'            If rsAddy.State = adStateOpen Then
'                If Me(Me.FindIDX(VirtualID)).colResellers_SnailMail.Count < rsAddy!RecCount Then
'                    SQL = "select * from visp_addresses where visp_RecID = '" & VirtualID & "'"
'                    bResult = MySQL.OpenTable(oConn, rsAddy, , SQL)
'                    Me(Me.FindIDX(VirtualID)).colResellers_SnailMail.Clear
'                    If pbset = True Then pb.Max = pb.Max + rsAddy.RecordCount
'                    While Not rsAddy.EOF
'                        With Me(Me.FindIDX(VirtualID)).colResellers_SnailMail.Add("RECID" & rsAddy!RecID, rsAddy!RecID, rsAddy!visp_RecID, rsAddy!FlagID, rsAddy!DateAdded, IIf(IsNull(rsAddy!ContactName), "", rsAddy!ContactName), _
'                                                                                IIf(IsNull(rsAddy!Street1), "", rsAddy!Street1), IIf(IsNull(rsAddy!Street2), "", rsAddy!Street2), IIf(IsNull(rsAddy!Country), "", rsAddy!Country), _
'                                                                                IIf(IsNull(rsAddy!State), "", rsAddy!State), IIf(IsNull(rsAddy!Suburb), "", rsAddy!Suburb), IIf(IsNull(rsAddy!PostCode), "", rsAddy!PostCode), _
'                                                                                rsAddy!Cancelled, rsAddy!Checked, IIf(IsNull(rsAddy!PhotoURL), "", rsAddy!PhotoURL), fs_NoChanges, SESSION, "RECID" & rsAddy!RecID)
'                        End With
'                        If pbset = True Then
'                            If pb.Value + 2 >= pb.Max Then
'                                pb.Max = pb.Max + 10
'                            End If
'                            PopulateResellers = PopulateResellers + 1: pb.Value = pb.Value + 1
'                        End If
'                        rsAddy.MoveNext
'                    Wend
'                End If
'            End If
'            rsAddy.Close
'
'         Me.FetchStatus = fs_Idle
'        If pbset = True Then pb.Value = 0: Exit Function
'    End If
    
End Function

Public Function FindIDX(Optional VirtualID As Long) As Long

    Dim fg As Long
    Static LastResult As Long
    Static LastVirtualID As Long
    Static Initialised As Boolean
    
    If LastVirtualID <> VirtualID Or Initialised = False Then
        For fg = Me.Count To 1 Step -1
            If Me(fg).RecID = VirtualID Then
                FindIDX = fg
                Exit For
            End If
        Next
        LastResult = fg
        LastVirtualID = VirtualID
        Initialised = True
    Else
        FindIDX = LastResult
    End If
    
End Function


Public Property Get Item(vntIndexKey As Variant) As clsReseller
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

