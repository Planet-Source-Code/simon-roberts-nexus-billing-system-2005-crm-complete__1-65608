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
Public Function ReturnCatch(ByVal VirtualID As Long, Optional ByVal CatchType As String = "default") As String

    On Error Resume Next
    Dim lk As Long
    
    If Me(Me.FindIDX(VirtualID)).colReseller_emailcatch.Count > 0 Then
        
        For lk = 1 To Me(Me.FindIDX(VirtualID)).colReseller_emailcatch.Count
            If LCase(Me(Me.FindIDX(VirtualID)).colReseller_emailcatch(lk).CatchType) = LCase(CatchType) Then
                ReturnCatch = Me(Me.FindIDX(VirtualID)).colReseller_emailcatch(lk).EmailAddy
                Exit For
            End If
        Next lk
        
    
    Else
    
        ReturnCatch = "nexusarchives@ep.net.au"
        
    End If
    
    If Len(ReturnCatch) = 0 Then ReturnCatch = "nexusarchives@ep.net.au"
    

End Function

Public Sub AddObject(ByVal nReseller As clsReseller)
    
    mCol.Add nReseller, nReseller.Key
        
End Sub

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
    If Not mvarCurrentReseller Is Nothing Then
        Me.CurrentVirtualID = mvarCurrentReseller.RecID
    End If
    
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
        Set objNewMember.colReseller_emailcatch = New colReseller_emailcatch
        
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
Public Function CommitCurrentReseller(Optional ByRef bCreatNew As Boolean = False) As Long

    Dim SQL As String
    Dim vRecID As Long
    Dim tRecID As Long
    
    Select Case bCreatNew
    Case False
        
        vRecID = Me.CurrentReseller.RecID
            
    Case True
    
        On Error Resume Next
       
        Do
            vRecID = MySQL.GetTMPRecID("virtualisp", directConn, "RecID")
            MySQL.Execute directConn, "INSERT INTO virtualisp (RecID) VALUES ('" & vRecID & "')"
        Loop Until Err.Number = 0
        
        Me.CurrentVirtualID = vRecID
        Me.CurrentReseller.RecID = vRecID
        Me.CurrentReseller.CreatedBy_SysopID = Login.lSysopID
        Me.CurrentReseller.VirtualID = Login.lVirtualID
        Me.CurrentReseller.AgencyID = Login.lAgencyID
        
    End Select

        ' ******* 16-10-2004 ******  Abbiibbiit Abbbittit Thats all fokes
        ' Virtual/Reseller Platform has full collection and class support
        ' framework.
        ' 1:47pm Simon Roberts
        ' ***************************************************************
        
        SQL = "update sysops set VirtualID = '" & vRecID & "' where RecID = '" & Me.CurrentReseller.SysopID & "'"
        MySQL.Execute directConn, SQL
        
        SQL = "update virtualisp set " + _
               "ABN = '" & MySQL.ESC(Me.CurrentReseller.ABN) & "'" & _
               ",ACN = '" & MySQL.ESC(Me.CurrentReseller.ACN) & "'" & _
                ",AgencyID = '" & Me.CurrentReseller.AgencyID & "'" & _
                ",AssistanceManager_SysopID = '" & Me.CurrentReseller.AssistanceManager_SysopID & "'" & _
                ",BriefDesc = '" & MySQL.ESC(CStr(Me.CurrentReseller.BriefDesc)) & "'" & _
                ",bTaxMode = '" & Me.CurrentReseller.bTaxMode & "'" & _
                ",CreatedBy_SysopID = '" & Me.CurrentReseller.CreatedBy_SysopID & "'" & _
                ",cTaxCode = '" & MySQL.ESC(Me.CurrentReseller.cTaxCode) & "'" & _
                ",cTaxCountry = '" & MySQL.ESC(Me.CurrentReseller.cTaxCountry) & "'" & _
                ",Cycle_IntervalLength = '" & Me.CurrentReseller.Cycle_IntervalLength & "'" & _
                ",Cycle_IntervalType = '" & Me.CurrentReseller.Cycle_IntervalType & "'" & _
                ",Description = '" & MySQL.ESC(Me.CurrentReseller.Description) & "'" & _
                ",ftpBasePath = '" & MySQL.ESC(Me.CurrentReseller.ftpBasePath) & "'" & _
                ",ftpCostPerMB = '" & Me.CurrentReseller.ftpCostPerMB & "'" & _
                ",ftpFileDBMode = '" & Me.CurrentReseller.ftpFileDBMode & "'" & _
                ",ftpGroupingFolder = '" & MySQL.ESC(Me.CurrentReseller.ftpGroupingFolder) & "'" & _
                ",ftpHostName = '" & MySQL.ESC(Me.CurrentReseller.ftpHostName) & "'" & _
                ",ftpIEProxy = '" & Me.CurrentReseller.ftpIEProxy & "'" & _
                ",ftpNumberofFiles = '" & Me.CurrentReseller.ftpNumberofFiles & "'" & _
                ",ftpNumberofFolders = '" & Me.CurrentReseller.ftpNumberofFolders & "'" & _
                ",ftpPassword = AES_ENCRYPT('" & MySQL.ESC(Me.CurrentReseller.ftpPassword) & "','" & odb.colSalts.ReturnSalt("ftp") & "')" & _
                ",ftpPingAlive = '" & Me.CurrentReseller.ftpPingAlive & "'" & _
                ",ftpPort = '" & MySQL.ESC(Me.CurrentReseller.ftpPort) & "'"
        SQL = SQL & ",ftpProxy = '" & MySQL.ESC(Me.CurrentReseller.ftpProxy) & "'" & _
                ",ftpTotalMB = '" & Me.CurrentReseller.ftpTotalMB & "'" & _
                ",ftpURLPath = '" & MySQL.ESC(Me.CurrentReseller.ftpURLPath) & "'" & _
                ",ftpTotalMB = '" & Me.CurrentReseller.ftpTotalMB & "'" & _
                ",ftpUsername = '" & Me.CurrentReseller.ftpUsername & "'" & _
                ",Icon = '" & Me.CurrentReseller.Icon & "'" & _
                ",JoiningFee = '" & Me.CurrentReseller.JoiningFee & "'" & _
                ",LogoURL = '" & MySQL.ESC(Me.CurrentReseller.LogoURL) & "'" & _
                ",Manager_SysopID = '" & Me.CurrentReseller.Manager_SysopID & "'" & _
                ",MISCFee = '" & Me.CurrentReseller.MISCFee & "'" & _
                ",NextCycle = '" & Format(Me.CurrentReseller.NextCycle, "yyyy-mm-dd ttttt") & "'" & _
                ",NoSub = '" & Me.CurrentReseller.NoSub & "'" & _
                ",NumSales_AdminFeePerSale = '" & Me.CurrentReseller.NumSales_AdminFeePerSale & "'" & _
                ",NumSales_CappingAt = '" & Me.CurrentReseller.NumSales_CappingAt & "'" & _
                ",NumSales_Maximum = '" & Me.CurrentReseller.NumSales_Maximum & "'" & _
                ",NumSales_Minimum = '" & Me.CurrentReseller.NumSales_Minimum & "'" & _
                ",PreviousCycle = '" & Format(Me.CurrentReseller.PreviousCycle, "yyyy-mm-dd ttttt") & "'" & _
                ",Realm = '" & MySQL.ESC(Me.CurrentReseller.Realm) & "'" & _
                ",RecID = '" & vRecID & "'" & _
                ",Subscribed = '" & Me.CurrentReseller.Subscribed & "'" & _
                ",SysopID = '" & Me.CurrentReseller.SysopID & "'" & _
                ",VirtualID = '" & Me.CurrentReseller.VirtualID & "'"
        SQL = SQL & " where RecID = " & vRecID
    
        MySQL.Execute directConn, SQL
    
    bCreatNew = False
    
    Dim bDoIt As Boolean
    
    If Me.CurrentReseller.colResellers_AccountHeader.Count > 0 Then
        
        
        Dim objAccHeader As colResellers_AccountHeader
        Set objAccHeader = Me.CurrentReseller.colResellers_AccountHeader
        
        Dim objAccItem As clsResellers_AccountHeader
        
        For Each objAccItem In objAccHeader
        
            If objAccItem.FetchStatus = fs_NewLine_Insert Then
                
                Do
                    tRecID = MySQL.GetTMPRecID("visp_accountheader", directConn, "RecID")
                    directConn.Execute "Insert into visp_accountheader (RecID) Values ('" & tRecID & "')"
                Loop Until Err.Number = 0
                bDoIt = True
                objAccItem.RecID = tRecID
                objAccItem.VirtualID = vRecID
            ElseIf objAccItem.FetchStatus = fs_Edited Then
            
                tRecID = objAccItem.RecID
                bDoIt = True
                
            End If
        
            If bDoIt = True Then
                SQL = "update visp_accountheader set " & _
                      "Adjustments_DroughtAllowance = '" & objAccItem.Adjustments_DroughtAllowance & "'" & _
                      ",Adjustments_Expenses = '" & objAccItem.Adjustments_Expenses & "'" & _
                      ",Adjustments_Income = '" & objAccItem.Adjustments_Income & "'" & _
                      ",Adjustments_ReconciliationItems = '" & objAccItem.Adjustments_ReconciliationItems & "'" & _
                      ",ChequeNumber = '" & MySQL.ESC(objAccItem.ChequeNumber) & "'" & _
                      ",EndDate = '" & objAccItem.EndDate & "'" & _
                      ",Expenses_AllOther = '" & objAccItem.Expenses_AllOther & "'" & _
                      ",Expenses_BadDebts = '" & objAccItem.Expenses_BadDebts & "'" & _
                      ",Expenses_Commission = '" & objAccItem.Expenses_Commission & "'" & _
                      ",Expenses_CostofSale = '" & objAccItem.Expenses_CostofSale & "'" & _
                      ",Expenses_Depreciation = '" & objAccItem.Expenses_Depreciation & "'" & _
                      ",Expenses_Lease = '" & objAccItem.Expenses_Lease & "'" & _
                      ",Expenses_MotorVechicle = '" & objAccItem.Expenses_MotorVechicle & "'" & _
                      ",Expenses_Rent = '" & objAccItem.Expenses_Rent & "'" & _
                      ",Expenses_RepairsMaintenance = '" & objAccItem.Expenses_RepairsMaintenance & "'" & _
                      ",Expenses_Superannuation = '" & objAccItem.Expenses_Superannuation & "'" & _
                      ",Expenses_TotalInterest = '" & objAccItem.Expenses_TotalInterest & "'" & _
                      ",Expenses_TotalOverall = '" & objAccItem.Expenses_TotalOverall & "'" & _
                      ",Expenses_TotalRoyalty = '" & objAccItem.Expenses_TotalRoyalty & "'" & _
                      ",FlagID = '" & objAccItem.FlagID & "'" & _
                      ",Income_AssessableGovernment = '" & objAccItem.Income_AssessableGovernment & "'" & _
                      ",Income_GrossPayments = '" & objAccItem.Income_GrossPayments & "'" & _
                      ",Income_IndustryPayments = '" & objAccItem.Income_IndustryPayments & "'"
                SQL = SQL & ",Income_OtherBusiness = '" & objAccItem.Income_OtherBusiness & "'" & _
                      ",Income_TotalOverall = '" & objAccItem.Income_TotalOverall & "'" & _
                      ",InvoiceID = '" & objAccItem.InvoiceID & "'" & _
                      ",Margin = '" & objAccItem.Margin & "'" & _
                      ",Paid = '" & objAccItem.Paid & "'" & _
                      ",RecID = '" & tRecID & "'" & _
                      ",StartDate = '" & objAccItem.StartDate & "'" & _
                      ",sum_NetFromBusiness = '" & objAccItem.sum_NetFromBusiness & "'" & _
                      ",TotalCost = '" & objAccItem.TotalCost & "'" & _
                      ",TotalIncome = '" & objAccItem.TotalIncome & "'" & _
                      ",VirtualID = '" & objAccItem.VirtualID & "'" & _
                      ",visp_PaymentID = '" & objAccItem.visp_PaymentID & "'"
                      
                SQL = SQL & " where RecID = " & tRecID
                
                MySQL.Execute directConn, SQL
                
            End If
            objAccItem.FetchStatus = fs_NoChanges
            bDoIt = False
        Next
    
    End If
    
    
    If Me.CurrentReseller.colResellers_EmailAddy.Count > 0 Then
        
        Dim objEmladdy As colResellers_EmailAddy
        Set objEmladdy = Me.CurrentReseller.colResellers_EmailAddy
        
        Dim objEmlItem As clsResellers_EmailAddy
        
        For Each objEmlItem In objEmladdy
        
            If objEmlItem.FetchStatus = fs_NewLine_Insert Then
                
                Do
                    tRecID = MySQL.GetTMPRecID("visp_emailaddresses", directConn, "RecID")
                    directConn.Execute "Insert into visp_emailaddresses (RecID) Values ('" & tRecID & "')"
                Loop Until Err.Number = 0
                bDoIt = True
                objEmlItem.RecID = tRecID
                objEmlItem.visp_RecID = vRecID
            ElseIf objEmlItem.FetchStatus = fs_Edited Then
            
                tRecID = objEmlItem.RecID
                bDoIt = True
                
            End If
        
            If bDoIt = True Then
                SQL = "update visp_emailaddresses set " & _
                      "Cancelled = '" & objEmlItem.Cancelled & "'" & _
                      ",Checked = '" & objEmlItem.Checked & "'" & _
                      ",ContactName = '" & MySQL.ESC(objEmlItem.ContactName) & "'" & _
                      ",EmailAddress = '" & MySQL.ESC(objEmlItem.EmailAddress) & "'" & _
                      ",FlagID = '" & objEmlItem.FlagID & "'" & _
                      ",PhotoURL = '" & MySQL.ESC(objEmlItem.PhotoURL) & "'" & _
                      ",RecID = '" & tRecID & "'" & _
                      ",visp_RecID  = '" & vRecID & "'"

                SQL = SQL & " where RecID = " & tRecID
                
                MySQL.Execute directConn, SQL
                
            End If
            bDoIt = False
            objEmlItem.FetchStatus = fs_NoChanges
        Next
    
    End If
    
    
    If Me.CurrentReseller.colResellers_FoneNum.Count > 0 Then
        
        Dim objFoneaddy As colResellers_FoneNum
        Set objFoneaddy = Me.CurrentReseller.colResellers_FoneNum
        
        Dim objFoneItem As clsResellers_FoneNum
        
        For Each objFoneItem In objFoneaddy
        
            If objFoneItem.FetchStatus = fs_NewLine_Insert Then
                
                Do
                    tRecID = MySQL.GetTMPRecID("visp_phonenumbers", directConn, "RecID")
                    directConn.Execute "Insert into visp_phonenumbers (RecID) Values ('" & tRecID & "')"
                Loop Until Err.Number = 0
                bDoIt = True
                objFoneItem.RecID = tRecID
                objFoneItem.visp_RecID = vRecID
                
            ElseIf objFoneItem.FetchStatus = fs_Edited Then
            
                tRecID = objFoneItem.RecID
                bDoIt = True
                
            End If
        
            If bDoIt = True Then
                SQL = "update visp_phonenumbers set " & _
                      "Cancelled = '" & objFoneItem.Cancelled & "'" & _
                      ",Checked = '" & objFoneItem.Checked & "'" & _
                      ",ContactName = '" & MySQL.ESC(objFoneItem.ContactName) & "'" & _
                      ",Extension = '" & MySQL.ESC(objFoneItem.Extension) & "'" & _
                      ",FlagID = '" & objFoneItem.FlagID & "'" & _
                      ",PhotoURL = '" & MySQL.ESC(objFoneItem.PhotoURL) & "'" & _
                      ",PhoneNumber = '" & MySQL.ESC(objFoneItem.PhoneNumber) & "'" & _
                      ",ShortNote = '" & MySQL.ESC(objFoneItem.ShortNote) & "'" & _
                      ",RecID = '" & tRecID & "'" & _
                      ",visp_RecID  = '" & vRecID & "'"

                SQL = SQL & " where RecID = " & tRecID
                
                MySQL.Execute directConn, SQL
                
            End If
            objFoneItem.FetchStatus = fs_NoChanges
            bDoIt = False
        Next
    
    End If
    
    
    If Me.CurrentReseller.colResellers_InvoiceItems.Count > 0 Then
        
        Dim objInvaddy As colResellers_InvoiceItems
        Set objInvaddy = Me.CurrentReseller.colResellers_InvoiceItems
        
        Dim objInvItem As clsResellers_InvoiceItems
        
        For Each objInvItem In objInvaddy
        
            If objInvItem.FetchStatus = fs_NewLine_Insert Then
                
                Do
                    tRecID = MySQL.GetTMPRecID("visp_invoiceitems", directConn, "RecID")
                    directConn.Execute "Insert into visp_invoiceitems (RecID) Values ('" & tRecID & "')"
                Loop Until Err.Number = 0
                bDoIt = True
                
                objInvItem.RecID = tRecID
                objInvItem.VirtualID = vRecID
                
            ElseIf objInvItem.FetchStatus = fs_Edited Then
            
                tRecID = objInvItem.RecID
                bDoIt = True
                
            End If
        
            If bDoIt = True Then
                SQL = "update visp_invoiceitems set " & _
                      "AmountDue = '" & objInvItem.AmountDue & "'" & _
                      ",AmountPaid = '" & objInvItem.AmountPaid & "'" & _
                      ",CostPrice = '" & objInvItem.CostPrice & "'" & _
                      ",Description = '" & MySQL.ESC(objInvItem.Description) & "'" & _
                      ",FinalClosureDate = '" & objInvItem.FinalClosureDate & "'" & _
                      ",FlagID = '" & objInvItem.FlagID & "'" & _
                      ",GSTCharged = '" & objInvItem.GSTCharged & "'" & _
                      ",GSTPaid = '" & objInvItem.GSTPaid & "'" & _
                      ",InvoiceID = '" & objInvItem.InvoiceID & "'" & _
                      ",PaidWhen = '" & Format(objInvItem.PaidWhen, "yyyy-mm-dd ttttt") & "'" & _
                      ",PaymentDue = '" & Format(objInvItem.PaymentDue, "yyyy-mm-dd ttttt") & "'" & _
                      ",RecID = '" & tRecID & "'" & _
                      ",SysopID = '" & objInvItem.SysopID & "'" & _
                      ",TotalDue = '" & objInvItem.TotalDue & "'" & _
                      ",TotalPaid = '" & objInvItem.TotalPaid & "'" & _
                      ",VirtualID = '" & vRecID & "'"
                       
                SQL = SQL & " where RecID = " & tRecID
                MySQL.Execute directConn, SQL
            End If
            bDoIt = False
            objInvItem.FetchStatus = fs_NoChanges
        Next
    End If
    
    
    If Me.CurrentReseller.colResellers_Payments.Count > 0 Then
        
        Dim objPaytaddy As colResellers_Payments
        Set objPaytaddy = Me.CurrentReseller.colResellers_Payments
        
        Dim objPaytItem As clsResellers_Payments
        
        For Each objPaytItem In objPaytaddy
        
            If objPaytItem.FetchStatus = fs_NewLine_Insert Then
                
                Do
                    tRecID = MySQL.GetTMPRecID("visp_paymenthistory", directConn, "RecID")
                    directConn.Execute "Insert into visp_paymenthistory (RecID) Values ('" & tRecID & "')"
                Loop Until Err.Number = 0
                bDoIt = True
                
                
                objPaytItem.VirtualID = vRecID
                
            ElseIf objPaytItem.FetchStatus = fs_Edited Then
            
                
            End If
        
            If bDoIt = True Then
                SQL = "update visp_paymenthistory set " & _
                      "AmountDue = '" & objPaytItem.AmountPaid & "'" & _
                      ",DateBanked = '" & Format(objPaytItem.DateBanked, "yyyy-mm-dd ttttt") & "'" & _
                      ",DateKeyed = '" & Format(objPaytItem.DateKeyed, "yyyy-mm-dd ttttt") & "'" & _
                      ",FlagID = '" & objPaytItem.FlagID & "'" & _
                      ",InvoiceID = '" & objPaytItem.InvoiceID & "'" & _
                      ",RemittanceSent = '" & objPaytItem.RemittanceSent & "'" & _
                      ",SysopID = '" & objPaytItem.SysopID & "'" & _
                      ",VirtualID = '" & vRecID & "'"
                       
                SQL = SQL & " where RecID = " & tRecID
                MySQL.Execute directConn, SQL
            End If
            bDoIt = False
            objPaytItem.FetchStatus = fs_NoChanges
        Next
    End If
    
    If Me.CurrentReseller.colResellers_SnailMail.Count > 0 Then
        
        Dim ih As Long
        For ih = 1 To Me.CurrentReseller.colResellers_SnailMail.Count
        
            If Me.CurrentReseller.colResellers_SnailMail.Item(ih).FetchStatus = fs_NewLine_Insert Then
                
                Do
                    tRecID = MySQL.GetTMPRecID("visp_addresses", directConn, "RecID")
                    directConn.Execute "Insert into visp_addresses(RecID) Values ('" & tRecID & "')"
                Loop Until Err.Number = 0
                bDoIt = True
                
                Me.CurrentReseller.colResellers_SnailMail.Item(ih).RecID = tRecID
                Me.CurrentReseller.colResellers_SnailMail.Item(ih).visp_RecID = vRecID
                
            ElseIf Me.CurrentReseller.colResellers_SnailMail.Item(ih).FetchStatus = fs_Edited Then
            
                tRecID = Me.CurrentReseller.colResellers_SnailMail.Item(ih).RecID
                bDoIt = True
                
            End If
        
            If bDoIt = True Then
                SQL = "update visp_addresses set " & _
                      "Cancelled = '" & Me.CurrentReseller.colResellers_SnailMail.Item(ih).Cancelled & "'" & _
                      ",Checked = '" & Me.CurrentReseller.colResellers_SnailMail.Item(ih).Checked & "'" & _
                      ",ContactName = '" & MySQL.ESC(Me.CurrentReseller.colResellers_SnailMail.Item(ih).ContactName) & "'" & _
                      ",Country = '" & MySQL.ESC(Me.CurrentReseller.colResellers_SnailMail.Item(ih).Country) & "'" & _
                      ",FlagID = '" & Me.CurrentReseller.colResellers_SnailMail.Item(ih).FlagID & "'" & _
                      ",PhotoURL = '" & MySQL.ESC(Me.CurrentReseller.colResellers_SnailMail.Item(ih).PhotoURL) & "'" & _
                      ",RecID = '" & tRecID & "'" & _
                      ",State = '" & MySQL.ESC(Me.CurrentReseller.colResellers_SnailMail.Item(ih).State) & "'" & _
                      ",Street1 = '" & MySQL.ESC(Me.CurrentReseller.colResellers_SnailMail.Item(ih).Street1) & "'" & _
                      ",Street2 = '" & MySQL.ESC(Me.CurrentReseller.colResellers_SnailMail.Item(ih).Street2) & "'" & _
                      ",Suburb = '" & MySQL.ESC(Me.CurrentReseller.colResellers_SnailMail.Item(ih).Suburb) & "'" & _
                      ",visp_RecID = '" & vRecID & "'"

                       
                SQL = SQL & " where RecID = " & tRecID
                MySQL.Execute directConn, SQL
            End If
            Me.CurrentReseller.colResellers_SnailMail.Item(ih).FetchStatus = fs_NoChanges
            bDoIt = False
        Next
    End If
End Function

Public Function PopulateResellers(ByRef oConn As adodb.Connection, ByRef ActionTyp As enumFetchStatus, ByVal VirtualID As Long, ByVal SESSION As String, Optional pb As ProgressBar, Optional ByVal pbset As Boolean) As Long

    Me.FetchStatus = fs_LoadingData
    
    Static tmpSession As String
    
    Dim SQL As String
    
    Dim rsKewl As New adodb.Recordset
    Dim rsPhone As adodb.Recordset
    Dim rsAddy As adodb.Recordset
    Dim rseMail As adodb.Recordset
    Dim rsHeader As adodb.Recordset
    Dim rsInvoice As adodb.Recordset
    
    If Len(tmpSession) = 0 And Len(SESSION) > 0 Then
        tmpSession = SESSION
    ElseIf Len(tmpSession) > 0 And Len(SESSION) = 0 Then
        SESSION = tmpSession
    End If
    
    If pbset = True And Not pb Is Nothing Then
        
        If Not pb Is Nothing Then
            pb.Value = 0
            pb.Max = 10
        End If
    End If
    
    If ActionTyp = fs_LoadHeader Or ActionTyp = fs_LoadAll Or ActionTyp = fs_LoadMinimum Then
            
        If VirtualID = 0 Then Me.Clear
        
        SQL = "select distinct AES_DECRYPT(ftpPassword,'" + odb.colSalts.ReturnSalt("ftp") + "') as fPWD, ftpFileDBMode, ftpGroupingFolder, ftpIEProxy, ftpPingAlive, ftpNumberofFolders, ftpNumberofFiles, RecID, AgencyID, VirtualID, Description, BriefDesc, Realm, CreationDate, ABN, Subscribed, SysopID, ACN, NoSub, CreatedBy_SysopID, NextCycle, PreviousCycle, Cycle_IntervalType, Cycle_IntervalLength, JoiningFee, LogoURL, Icon, Manager, Manager_SysopID, AssistanceManager_SysopID, Comment, MISCFee, bTaxMode, cTaxCode, cTaxCountry, cTaxExemptCode , ftpHostName, ftpUsername, ftpPassword, ftpPort, ftpBasePath, ftpTotalMB, ftpCostPerMB, ftpURLPath, NumSales_AdminFeePerSale, NumSales_CappingAt, NumSales_Minimum, NumSales_Maximum from virtualisp where RecID = '" & VirtualID & "'"
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
        
        If SESSION = "VISPDetails" Then Exit Function
        
        If pbset = True And Not pb Is Nothing Then
            
            If pb.Value + 2 >= pb.Max Then
                pb.Max = pb.Max + 10
            End If
            
            PopulateResellers = PopulateResellers + 1: pb.Value = pb.Value + 1
        End If
        
        SQL = "select AES_DECRYPT(ftpPassword,'" + odb.colSalts.ReturnSalt("ftp") + "') as fPWD, ftpFileDBMode, ftpGroupingFolder, ftpIEProxy, ftpPingAlive, ftpNumberofFolders, ftpNumberofFiles, RecID, AgencyID, VirtualID, Description, BriefDesc, Realm, CreationDate, ABN, Subscribed, SysopID, ACN, NoSub, CreatedBy_SysopID, NextCycle, PreviousCycle, Cycle_IntervalType, Cycle_IntervalLength, JoiningFee, LogoURL, Icon, Manager, Manager_SysopID, AssistanceManager_SysopID, Comment, MISCFee, bTaxMode, cTaxCode, cTaxCountry, cTaxExemptCode , ftpHostName, ftpUsername, ftpPassword, ftpPort, ftpBasePath, ftpTotalMB, ftpCostPerMB, ftpURLPath, NumSales_AdminFeePerSale, NumSales_CappingAt, NumSales_Minimum, NumSales_Maximum from virtualisp where VirtualID = '" & VirtualID & "'"
        bResult = MySQL.OpenTable(oConn, rsKewl, , SQL, adOpenStatic, adLockOptimistic)
    
        If rsKewl.State = adStateOpen Then
            If rsKewl.RecordCount > 0 Then
                If pbset = True And Not pb Is Nothing Then pb.Max = pb.Max + rsKewl.RecordCount
                While Not rsKewl.EOF
                    
                    With Me.Add("RECID" & rsKewl!RecID, rsKewl!RecID, rsKewl!AgencyID, rsKewl!VirtualID, IIf(IsNull(rsKewl!Description), "", rsKewl!Description), IIf(IsNull(rsKewl!BriefDesc), "", rsKewl!BriefDesc), IIf(IsNull(rsKewl!Realm), "", rsKewl!Realm), _
                                rsKewl!CreationDate, IIf(IsNull(rsKewl!ABN), "", rsKewl!ABN), rsKewl!Subscribed, rsKewl!SysopID, IIf(IsNull(rsKewl!ACN), "", rsKewl!ACN), rsKewl!NoSub, rsKewl!CreatedBy_SysopID, IIf(IsNull(rsKewl!NextCycle), DateAdd("d", 12, NextCycle), DateAdd("d", 12, sysnow)), IIf(IsNull(rsKewl!PreviousCycle), sysnow, rsKewl!PreviousCycle), rsKewl!JoiningFee, _
                                IIf(IsNull(rsKewl!LogoURL), "", rsKewl!LogoURL), IIf(IsNull(rsKewl!Icon), Rnd * 130, rsKewl!Icon), IIf(IsNull(rsKewl!Manager), "", rsKewl!Manager), rsKewl!Manager_SysopID, rsKewl!AssistanceManager_SysopID, IIf(IsNull(rsKewl!Comment), "", rsKewl!Comment), _
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
                    
                    If pbset = True And Not pb Is Nothing Then
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
            
                    If Me(Me.FindIDX(Me(lk).RecID)).colReseller_emailcatch.Count = 0 Then
                        SQL = "select visp_emailcatch.visp_RecID IN " & INClause & " as bfound, visp_emailcatch.* from visp_emailcatch"
                        bResult = MySQL.OpenTable(oConn, rsPhone, , SQL)
                        'Me(Me.FindIDX(Me(lk).RecID)).colReseller_emailcatch.Clear
                        If pbset = True And Not pb Is Nothing Then pb.Max = pb.Max + rsPhone.RecordCount
                        While Not rsPhone.EOF
                            If rsPhone!bFound = 1 Then
                                With Me(Me.FindIDX(Me(lk).RecID)).colReseller_emailcatch.Add("RECID" & rsPhone!RecID, rsPhone!RecID, rsPhone!VirtualID, IIf(IsNull(rsPhone!CatchType), "default", rsPhone!CatchType), IIf(IsNull(rsPhone!CatchEmailAddress), "nexusarchives@ep.net.au", rsPhone!CatchEmailAddress), "RECID" & rsPhone!RecID)
                                
                                End With
                            End If
                            If pbset = True And Not pb Is Nothing Then
                                If pb.Value + 2 >= pb.Max Then
                                    pb.Max = pb.Max + 10
                                End If
                                PopulateResellers = PopulateResellers + 1: pb.Value = pb.Value + 1
                            End If
                            rsPhone.MoveNext
                        Wend
                    End If
                
                rsPhone.Close
            Next
            
            
            For lk = Me.Count To 1 Step -1
                'loads the records
            
                SQL = "select count(*) as RecCount from visp_phonenumbers where visp_RecID = '" & Me(lk).RecID & "'"
                bResult = MySQL.OpenTable(oConn, rsPhone, , SQL)
            
                If rsPhone.State = adStateOpen Then
                    If Me(Me.FindIDX(Me(lk).RecID)).colResellers_FoneNum.Count < rsPhone!RecCount Then
                        SQL = "select * from visp_phonenumbers where visp_RecID = '" & Me(lk).RecID & "'"
                        bResult = MySQL.OpenTable(oConn, rsPhone, , SQL)
                        Me(Me.FindIDX(Me(lk).RecID)).colResellers_FoneNum.Clear
                        If pbset = True And Not pb Is Nothing Then pb.Max = pb.Max + rsPhone.RecordCount
                        While Not rsPhone.EOF
                            With Me(Me.FindIDX(Me(lk).RecID)).colResellers_FoneNum.Add("RECID" & rsPhone!RecID, rsPhone!RecID, rsPhone!visp_RecID, rsPhone!FlagID, rsPhone!DateAdded, IIf(IsNull(rsPhone!PhoneNumber), "", rsPhone!PhoneNumber), _
                                                                                  IIf(IsNull(rsPhone!Extension), "", rsPhone!Extension), IIf(IsNull(rsPhone!ContactName), "", rsPhone!ContactName), rsPhone!Cancelled, rsPhone!Checked, _
                                                                                  IIf(IsNull(rsPhone!PhotoURL), "", rsPhone!PhotoURL), fs_NoChanges, SESSION, "RECID" & rsPhone!RecID)
                                .ShortNote = IIf(IsNull(rsPhone!ShortNote), "", rsPhone!ShortNote)
                            End With
                            If pbset = True And Not pb Is Nothing Then
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
                        If pbset = True And Not pb Is Nothing Then pb.Max = pb.Max + rseMail.RecordCount
                        While Not rseMail.EOF
                            With Me(Me.FindIDX(Me(lk).RecID)).colResellers_EmailAddy.Add("RECID" & rseMail!RecID, rseMail!RecID, rseMail!visp_RecID, rseMail!FlagID, rseMail!DateAdded, IIf(IsNull(rseMail!EmailAddress), "", rseMail!EmailAddress), _
                                                                                    IIf(IsNull(rseMail!ContactName), "", rseMail!ContactName), rseMail!Cancelled, rseMail!Checked, IIf(IsNull(rseMail!PhotoURL), "", rseMail!PhotoURL), fs_NoChanges, _
                                                                                    SESSION, "RECID" & rseMail!RecID)
                            End With
                            If pbset = True And Not pb Is Nothing Then
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
                        If pbset = True And Not pb Is Nothing Then pb.Max = pb.Max + rsAddy.RecordCount
                        While Not rsAddy.EOF
                            With Me(Me.FindIDX(Me(lk).RecID)).colResellers_SnailMail.Add("RECID" & rsAddy!RecID, rsAddy!RecID, rsAddy!visp_RecID, rsAddy!FlagID, rsAddy!DateAdded, IIf(IsNull(rsAddy!ContactName), "", rsAddy!ContactName), _
                                                                                    IIf(IsNull(rsAddy!Street1), "", rsAddy!Street1), IIf(IsNull(rsAddy!Street2), "", rsAddy!Street2), IIf(IsNull(rsAddy!Country), "", rsAddy!Country), _
                                                                                    IIf(IsNull(rsAddy!State), "", rsAddy!State), IIf(IsNull(rsAddy!Suburb), "", rsAddy!Suburb), IIf(IsNull(rsAddy!PostCode), "", rsAddy!PostCode), _
                                                                                    rsAddy!Cancelled, rsAddy!Checked, IIf(IsNull(rsAddy!PhotoURL), "", rsAddy!PhotoURL), fs_NoChanges, SESSION, "RECID" & rsAddy!RecID)
                            End With
                            If pbset = True And Not pb Is Nothing Then
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
                            If pbset = True And Not pb Is Nothing Then pb.Max = pb.Max + rsHeader.RecordCount
                            While Not rsPhone.EOF
                                With Me(Me.FindIDX(Me(lk).RecID)).colResellers_AccountHeader.Add("RECID" & rsHeader!RecID, rsHeader!RecID, rsHeader!Me(lk).RecID, rsHeader!InvoiceID, rsHeader!StartDate, rsHeader!EndDate, rsHeader!TotalIncome, _
                                                                                            rsHeader!TotalCost, rsHeader!Paid, IIf(IsNull(rsHeader!ChequeNumber), "", rsHeader!ChequeNumber), rsHeader!Margin, rsHeader!visp_PaymentID, rsHeader!Income_GrossPayments, rsHeader!Income_AssessableGovernment, _
                                                                                            rsHeader!Income_IndustryPayments, rsHeader!Income_OtherBusiness, rsHeader!Income_TotalOverall, rsHeader!Expenses_Superannuation, rsHeader!Expenses_Commission, _
                                                                                            rsHeader!Expenses_CostofSale, rsHeader!Expenses_BadDebts, rsHeader!Expenses_Lease, rsHeader!Expenses_Rent, rsHeader!Expenses_TotalInterest, rsHeader!Expenses_TotalRoyalty, _
                                                                                            rsHeader!Expenses_Depreciation, rsHeader!Expenses_MotorVechicle, rsHeader!Expenses_RepairsMaintenance, rsHeader!Expenses_AllOther, rsHeader!Expenses_TotalOverall, _
                                                                                            rsHeader!Adjustments_Expenses, rsHeader!Adjustments_ReconciliationItems, rsHeader!Adjustments_Income, rsHeader!Adjustments_DroughtAllowance, rsHeader!sum_NetFromBusine, _
                                                                                            rsHeader!FlagID, fs_NoChanges, SESSION, "RECID" & rsHeader!RecID)
                                End With
                                If pbset = True And Not pb Is Nothing Then
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
                            If pbset = True And Not pb Is Nothing Then pb.Max = pb.Max + rsInvoice.RecordCount
                            While Not rsInvoice.EOF
                                With Me(Me.FindIDX(Me(lk).RecID)).colResellers_InvoiceItems.Add("RECID" & rsInvoice!RecID, rsInvoice!RecID, rsInvoice!InvoiceID, rsInvoice!AmountDue, rsInvoice!CostPrice, rsInvoice!GSTCharged, rsInvoice!PaymentDue, rsInvoice!FinalClosureDate, _
                                                                                                rsInvoice!AmountPaid, rsInvoice!GSTPaid, rsInvoice!PaidWhen, rsInvoice!TotalDue, rsInvoice!TotalPaid, rsInvoice!VirtualID, rsInvoice!SysopID, rsInvoice!FlagID, IIf(IsNull(rsInvoice!Description), "", rsInvoice!Description), fs_NoChanges, SESSION, "RECID" & rsInvoice!RecID)
                                End With
                                If pbset = True And Not pb Is Nothing Then
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
            If pbset = True And Not pb Is Nothing Then pb.Value = 0: Exit Function
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
                If pbset = True And Not pb Is Nothing Then pb.Max = pb.Max + rsPhone.RecordCount
                While Not rsPhone.EOF
                    With Me(Me.FindIDX(VirtualID)).colResellers_FoneNum.Add("RECID" & rsPhone!RecID, rsPhone!RecID, rsPhone!visp_RecID, rsPhone!FlagID, rsPhone!DateAdded, IIf(IsNull(rsPhone!PhoneNumber), "", rsPhone!PhoneNumber), _
                                                                          IIf(IsNull(rsPhone!Extension), "", rsPhone!Extension), IIf(IsNull(rsPhone!ContactName), "", rsPhone!ContactName), rsPhone!Cancelled, rsPhone!Checked, _
                                                                          IIf(IsNull(rsPhone!PhotoURL), "", rsPhone!PhotoURL), fs_NoChanges, SESSION, "RECID" & rsPhone!RecID)
                        .ShortNote = IIf(IsNull(rsPhone!ShortNote), "", rsPhone!ShortNote)
                    End With
                    If pbset = True And Not pb Is Nothing Then
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
                If pbset = True And Not pb Is Nothing Then pb.Max = pb.Max + rseMail.RecordCount
                While Not rseMail.EOF
                    With Me(Me.FindIDX(VirtualID)).colResellers_EmailAddy.Add("RECID" & rseMail!RecID, rseMail!RecID, rseMail!visp_RecID, rseMail!FlagID, rseMail!DateAdded, IIf(IsNull(rseMail!EmailAddress), "", rseMail!EmailAddress), _
                                                                            IIf(IsNull(rseMail!ContactName), "", rseMail!ContactName), rseMail!Cancelled, rseMail!Checked, IIf(IsNull(rseMail!PhotoURL), "", rseMail!PhotoURL), fs_NoChanges, _
                                                                            SESSION, "RECID" & rseMail!RecID)
                    End With
                    If pbset = True And Not pb Is Nothing Then
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
                 If pbset = True And Not pb Is Nothing Then pb.Max = pb.Max + rsAddy.RecordCount
                 While Not rsAddy.EOF
                     With Me(Me.FindIDX(VirtualID)).colResellers_SnailMail.Add("RECID" & rsAddy!RecID, rsAddy!RecID, rsAddy!visp_RecID, rsAddy!FlagID, rsAddy!DateAdded, IIf(IsNull(rsAddy!ContactName), "", rsAddy!ContactName), _
                                                                             IIf(IsNull(rsAddy!Street1), "", rsAddy!Street1), IIf(IsNull(rsAddy!Street2), "", rsAddy!Street2), IIf(IsNull(rsAddy!Country), "", rsAddy!Country), _
                                                                             IIf(IsNull(rsAddy!State), "", rsAddy!State), IIf(IsNull(rsAddy!Suburb), "", rsAddy!Suburb), IIf(IsNull(rsAddy!PostCode), "", rsAddy!PostCode), _
                                                                             rsAddy!Cancelled, rsAddy!Checked, IIf(IsNull(rsAddy!PhotoURL), "", rsAddy!PhotoURL), fs_NoChanges, SESSION, "RECID" & rsAddy!RecID)
                     End With
                     If pbset = True And Not pb Is Nothing Then
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
                If pbset = True And Not pb Is Nothing Then pb.Max = pb.Max + rsHeader.RecordCount
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
                    If pbset = True And Not pb Is Nothing Then
                        If pb.Value + 2 >= pb.Max Then
                            pb.Max = pb.Max + 10
                        End If
                        PopulateResellers = PopulateResellers + 1: pb.Value = pb.Value + 1
                    End If
                Wend
            End If
            rsHeader.Close
        End If
    End If
    
    
    
    ' Will load the invoice/tax header details of the Virtual ISP
    If ActionTyp = fs_LoadInvoice Then
        bResult = MySQL.OpenTable(oConn, rsInvoice, , "select * from visp_invoiceitems where VirtualID = '" & Me(lk).RecID & "'")
        If rsInvoice.State = adStateOpen Then
            If Me(Me.FindIDX(Me(lk).RecID)).colResellers_InvoiceItems.Count < rsInvoice.RecordCount Then
                Me(Me.FindIDX(Me(lk).RecID)).colResellers_InvoiceItems.Clear
                If pbset = True And Not pb Is Nothing Then pb.Max = pb.Max + rsInvoice.RecordCount
                While Not rsInvoice.EOF
                    With Me(Me.FindIDX(Me(lk).RecID)).colResellers_InvoiceItems.Add("RECID" & rsInvoice!RecID, rsInvoice!RecID, rsInvoice!InvoiceID, rsInvoice!AmountDue, rsInvoice!CostPrice, rsInvoice!GSTCharged, rsInvoice!PaymentDue, rsInvoice!FinalClosureDate, _
                                                                                    rsInvoice!AmountPaid, rsInvoice!GSTPaid, rsInvoice!PaidWhen, rsInvoice!TotalDue, rsInvoice!TotalPaid, rsInvoice!VirtualID, rsInvoice!SysopID, rsInvoice!FlagID, IIf(IsNull(rsInvoice!Description), "", rsInvoice!Description), fs_NoChanges, SESSION, "RECID" & rsInvoice!RecID)
                    End With
                    If pbset = True And Not pb Is Nothing Then
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
    If pbset = True And Not pb Is Nothing Then pb.Value = 0: Exit Function

    
    
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
'                    If pbset = True and Not pb Is Nothing Then pb.Max = pb.Max + rsPhone.RecordCount
'                    While Not rsPhone.EOF
'                        With Me(Me.FindIDX(VirtualID)).colResellers_FoneNum.Add("RECID" & rsPhone!RecID, rsPhone!RecID, rsPhone!visp_RecID, rsPhone!FlagID, rsPhone!DateAdded, IIf(IsNull(rsPhone!PhoneNumber), "", rsPhone!PhoneNumber), _
'                                                                              IIf(IsNull(rsPhone!Extension), "", rsPhone!Extension), IIf(IsNull(rsPhone!ContactName), "", rsPhone!ContactName), rsPhone!Cancelled, rsPhone!Checked, _
'                                                                              IIf(IsNull(rsPhone!PhotoURL), "", rsPhone!PhotoURL), fs_NoChanges, SESSION, "RECID" & rsPhone!RecID)
'                            .ShortNote = IIf(IsNull(rsPhone!ShortNote), "", rsPhone!ShortNote)
'                        End With
'                        If pbset = True and Not pb Is Nothing Then
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
'                    If pbset = True and Not pb Is Nothing Then pb.Max = pb.Max + rsEmail.RecordCount
'                    While Not rsEmail.EOF
'                        With Me(Me.FindIDX(VirtualID)).colResellers_EmailAddy.Add("RECID" & rsEmail!RecID, rsEmail!RecID, rsEmail!visp_RecID, rsEmail!FlagID, rsEmail!DateAdded, IIf(IsNull(rsEmail!EmailAddress), "", rsEmail!EmailAddress), _
'                                                                                IIf(IsNull(rsEmail!ContactName), "", rsEmail!ContactName), rsEmail!Cancelled, rsEmail!Checked, IIf(IsNull(rsEmail!PhotoURL), "", rsEmail!PhotoURL), fs_NoChanges, _
'                                                                                SESSION, "RECID" & rsEmail!RecID)
'                        End With
'                        If pbset = True and Not pb Is Nothing Then
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
'                    If pbset = True and Not pb Is Nothing Then pb.Max = pb.Max + rsAddy.RecordCount
'                    While Not rsAddy.EOF
'                        With Me(Me.FindIDX(VirtualID)).colResellers_SnailMail.Add("RECID" & rsAddy!RecID, rsAddy!RecID, rsAddy!visp_RecID, rsAddy!FlagID, rsAddy!DateAdded, IIf(IsNull(rsAddy!ContactName), "", rsAddy!ContactName), _
'                                                                                IIf(IsNull(rsAddy!Street1), "", rsAddy!Street1), IIf(IsNull(rsAddy!Street2), "", rsAddy!Street2), IIf(IsNull(rsAddy!Country), "", rsAddy!Country), _
'                                                                                IIf(IsNull(rsAddy!State), "", rsAddy!State), IIf(IsNull(rsAddy!Suburb), "", rsAddy!Suburb), IIf(IsNull(rsAddy!PostCode), "", rsAddy!PostCode), _
'                                                                                rsAddy!Cancelled, rsAddy!Checked, IIf(IsNull(rsAddy!PhotoURL), "", rsAddy!PhotoURL), fs_NoChanges, SESSION, "RECID" & rsAddy!RecID)
'                        End With
'                        If pbset = True and Not pb Is Nothing Then
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
'        If pbset = True and Not pb Is Nothing Then pb.Value = 0: Exit Function
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
 On Error Resume Next
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
  
  If Err.Number <> 0 Then Set Item = Nothing
  
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

