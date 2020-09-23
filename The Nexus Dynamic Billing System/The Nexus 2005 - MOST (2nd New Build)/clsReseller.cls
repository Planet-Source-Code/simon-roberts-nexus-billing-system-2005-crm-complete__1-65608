VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReseller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"colReseller_emailcatch"
Attribute VB_Ext_KEY = "Member1" ,"colResellers_AccountHeader"
Attribute VB_Ext_KEY = "Member2" ,"colResellers_EmailAddy"
Attribute VB_Ext_KEY = "Member3" ,"colResellers_FoneNum"
Attribute VB_Ext_KEY = "Member4" ,"colResellers_InvoiceItems"
Attribute VB_Ext_KEY = "Member5" ,"colResellers_Payments"
Attribute VB_Ext_KEY = "Member6" ,"colResellers_SnailMail"
Public Key As String

Enum enumPRIVFetchStatus
    NoChanges = 1
    Edited_Update = 2
    NewLine_Insert = 3
    DeleteRecord = 4
    LoadingData = 5
    Saving = 6
    Deleting = 7
    Idle = 8
    CreateNewViSP = 9
    AccountHeader = 10
    Addresses = 11
    EmailAddresses = 12
    PhoneNumbers = 13
    InvoiceItems = 14
    PaymentHistory = 15
    LoadHeader = 16
    LoadInvoice = 17
    LoadAllContactDetails = 18
    LoadPaymentHistory = 19
    LoadAll = 20
    LoadMinimum = 21
    LoadEmail = 22
    LoadPhone = 23
    LoadAddress = 24
End Enum

'local variable(s) to hold property value(s)
Private mvarRecID As Long 'local copy
Private mvarAgencyID As Long 'local copy
Private mvarVirtualID As Long 'local copy
Private mvarDescription As String 'local copy
Private mvarBriefDesc As Variant 'local copy
Private mvarRealm As String 'local copy
Private mvarCreationDate As Date 'local copy
Private mvarABN As String 'local copy
Private mvarSubscribed As Long 'local copy
Private mvarSysopID As Long 'local copy
Private mvarACN As String 'local copy
Private mvarNoSub As Long 'local copy
Private mvarCreatedBy_SysopID As Long 'local copy
Private mvarNextCycle As Date 'local copy
Private mvarPreviousCycle As Date 'local copy
Private mvarJoiningFee As Single 'local copy
Private mvarLogoURL As String 'local copy
Private mvarIcon As Integer 'local copy
Private mvarManager As Variant 'local copy
Private mvarManager_SysopID As Long 'local copy
Private mvarAssistanceManager_SysopID As Long 'local copy
Private mvarComment As Variant 'local copy
Private mvarMISCFee As Single 'local copy
Private mvarbTaxMode As Integer 'local copy
Private mvarcTaxCode As String 'local copy
Private mvarcTaxCountry As String 'local copy
Private mvarcTaxExemptCode As String 'local copy
Private mvarftpHostName As String 'local copy
Private mvarftpProxy As String 'local copy
Private mvarftpUsername As String 'local copy
Private mvarftpPassword As String 'local copy
Private mvarftpPort As String 'local copy
Private mvarftpBasePath As String 'local copy
Private mvarftpURLPath As String 'local copy
Private mvarftpTotalMB As Long 'local copy
Private mvarftpCostPerMB As Single 'local copy

Private mvarcolResellers_FoneNum As colResellers_FoneNum
Private mvarcolResellers_EmailAddy As colResellers_EmailAddy
Private mvarcolResellers_AccountHeader As colResellers_AccountHeader
Private mvarcolResellers_InvoiceItems As colResellers_InvoiceItems
Private mvarIDX As Long 'local copy
Private mvarSESSION As String 'local copy
Private mvarcolResellers_Payments As colResellers_Payments

Private mvarStatus As Byte 'local copy
Private mvarNumSales_Minimum As Integer 'local copy
Private mvarNumSales_Maximum As Integer 'local copy
Private mvarNumSales_CappingAt As Integer 'local copy
Private mvarNumSales_AdminFeePerSale As Single 'local copy
Private mvarCycle_IntervalType As String 'local copy
Private mvarCycle_IntervalLength As Integer 'local copy
Private mvarcolResellers_SnailMail As colResellers_SnailMail 'local cop
Private mvarFetchStatus As Byte 'local copy
'local variable(s) to hold property value(s)
Private mvarftpFileDBMode As Byte 'local copy
Private mvarftpPingAlive As Byte 'local copy
Private mvarftpIEProxy As Byte 'local copy
Private mvarftpNumberofFiles As Long 'local copy
Private mvarftpNumberofFolders As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarftpGroupingFolder As String 'local copy
Private mvarcolReseller_emailcatch As colReseller_emailcatch


Public Property Get colReseller_emailcatch() As colReseller_emailcatch
    If mvarcolReseller_emailcatch Is Nothing Then
        Set mvarcolReseller_emailcatch = New colReseller_emailcatch
    End If


    Set colReseller_emailcatch = mvarcolReseller_emailcatch
End Property


Public Property Set colReseller_emailcatch(vData As colReseller_emailcatch)
    Set mvarcolReseller_emailcatch = vData
End Property

Public Function GetNetworkXML(Optional ByRef XMLType As Byte = 1) As String

    Dim XML As String
    
    Select Case XMLType
    Case 1
    
        XML = XML & "<Hostname>" & mvarftpHostName & "</Hostname>"
        XML = XML & "<Proxy>" & mvarftpProxy & "</Proxy>"
        XML = XML & "<Username>" & mvarftpUsername & "</Username>"
        XML = XML & "<Password>" & mvarftpPassword & "</Password>"
        XML = XML & "<Port>" & mvarftpPort & "</Port>"
        XML = XML & "<BasePath>" & mvarftpBasePath & "</BasePath>"
        XML = XML & "<nTotalMB>" & mvarftpTotalMB & "</nTotalMB>"
        XML = XML & "<nCostMB>" & mvarftpCostPerMB & "</nCostMB>"
        XML = XML & "<nDBMode>" & mvarftpFileDBMode & "</nDBMode>"
        XML = XML & "<nPing>" & mvarftpPingAlive & "</nPing>"
        XML = XML & "<nIEProxy>" & mvarftpIEProxy & "</nIEProxy>"
        XML = XML & "<nFiles>" & mvarftpNumberofFiles & "</nFiles>"
        XML = XML & "<nFolders>" & mvarftpNumberofFolders & "</nFolders>"
        XML = XML & "<GroupingFolder>" & mvarftpGroupingFolder & "</GroupingFolder>"
        XML = XML & "<URLPath>" & mvarftpURLPath & "</URLPath>"
        
        'DragDrop = 2
        'LocalMachine = 4
        'LANMachine = 5
        'FTP_RemotleyHosted = 100
        'UNC_RemotelyHosted = 111
        'HTTP_RemotelyHosted = 150
        'xxx_OriginUnknown = 666
        
        Select Case mvarftpFileDBMode
        Case 0
            XML = XML & "<DefaultOrigin>111</DefaultOrigin>"
        Case 1
            XML = XML & "<DefaultOrigin>100</DefaultOrigin>"
        Case 2
            XML = XML & "<DefaultOrigin>4</DefaultOrigin>"
        Case 3
            XML = XML & "<DefaultOrigin>5</DefaultOrigin>"
        Case 4
            XML = XML & "<DefaultOrigin>150</DefaultOrigin>"
        Case Else
            XML = XML & "<DefaultOrigin>666</DefaultOrigin>"
        End Select
        
        GetNetworkXML = XML
        
    End Select
    
End Function

Public Function ParseNetworkXML(NetXML As String)

    mvarftpHostName = XMLVal(NetXML, "Hostname")
    mvarftpProxy = XMLVal(NetXML, "Proxy")
    mvarftpUsername = XMLVal(NetXML, "Username")
    mvarftpPassword = XMLVal(NetXML, "Password")
    mvarftpPort = XMLVal(NetXML, "Port")
    mvarftpBasePath = XMLVal(NetXML, "BasePath")
    mvarftpTotalMB = Val(XMLVal(NetXML, "nTotalMB"))
    mvarftpCostPerMB = Val(XMLVal(NetXML, "nCostMB"))
    mvarftpFileDBMode = Val(XMLVal(NetXML, "nDBMode"))
    mvarftpPingAlive = Val(XMLVal(NetXML, "nPing"))
    mvarftpIEProxy = Val(XMLVal(NetXML, "nIEProxy"))
    mvarftpNumberofFiles = Val(XMLVal(NetXML, "nFiles"))
    mvarftpNumberofFolders = Val(XMLVal(NetXML, "nFolders"))
    mvarftpGroupingFolder = XMLVal(NetXML, "GroupingFolder")
    mvarftpURLPath = XMLVal(NetXML, "URLPath")
        
End Function
Public Property Let ftpGroupingFolder(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ftpGroupingFolder = 5
    mvarftpGroupingFolder = vData
End Property


Public Property Get ftpGroupingFolder() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ftpGroupingFolder
    ftpGroupingFolder = mvarftpGroupingFolder
End Property



Public Property Let ftpNumberofFolders(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ftpNumberofFolders = 5
    mvarftpNumberofFolders = vData
End Property


Public Property Get ftpNumberofFolders() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ftpNumberofFolders
    ftpNumberofFolders = mvarftpNumberofFolders
End Property



Public Property Let ftpNumberofFiles(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ftpNumberofFiles = 5
    mvarftpNumberofFiles = vData
End Property


Public Property Get ftpNumberofFiles() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ftpNumberofFiles
    ftpNumberofFiles = mvarftpNumberofFiles
End Property



Public Property Let ftpIEProxy(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ftpIEProxy = 5
    mvarftpIEProxy = vData
End Property


Public Property Get ftpIEProxy() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ftpIEProxy
    ftpIEProxy = mvarftpIEProxy
End Property



Public Property Let ftpPingAlive(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ftpPingAlive = 5
    mvarftpPingAlive = vData
End Property


Public Property Get ftpPingAlive() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ftpPingAlive
    ftpPingAlive = mvarftpPingAlive
End Property



Public Property Let ftpFileDBMode(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ftpFileDBMode = 5
    mvarftpFileDBMode = vData
End Property


Public Property Get ftpFileDBMode() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ftpFileDBMode
    ftpFileDBMode = mvarftpFileDBMode
End Property



Public Property Let FetchStatus(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FetchStatus = 5
    mvarFetchStatus = vData
End Property


Public Property Get FetchStatus() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FetchStatus
    FetchStatus = mvarFetchStatus
End Property




Public Property Get colResellers_SnailMail() As colResellers_SnailMail
    If mvarcolResellers_SnailMail Is Nothing Then
        Set mvarcolResellers_SnailMail = New colResellers_SnailMail
    End If


    Set colResellers_SnailMail = mvarcolResellers_SnailMail
End Property


Public Property Set colResellers_SnailMail(vData As colResellers_SnailMail)
    Set mvarcolResellers_SnailMail = vData
End Property

Public Property Let Cycle_IntervalLength(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Cycle_IntervalLength = 5
    mvarCycle_IntervalLength = vData
End Property


Public Property Get Cycle_IntervalLength() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Cycle_IntervalLength
    Cycle_IntervalLength = mvarCycle_IntervalLength
End Property



Public Property Let Cycle_IntervalType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Cycle_IntervalType = 5
    mvarCycle_IntervalType = vData
End Property


Public Property Get Cycle_IntervalType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Cycle_IntervalType
    Cycle_IntervalType = mvarCycle_IntervalType
End Property



Public Property Let NumSales_AdminFeePerSale(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NumSales_AdminFeePerSale = 5
    mvarNumSales_AdminFeePerSale = vData
End Property


Public Property Get NumSales_AdminFeePerSale() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NumSales_AdminFeePerSale
    NumSales_AdminFeePerSale = mvarNumSales_AdminFeePerSale
End Property



Public Property Let NumSales_CappingAt(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NumSales_CappingAt = 5
    mvarNumSales_CappingAt = vData
End Property


Public Property Get NumSales_CappingAt() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NumSales_CappingAt
    NumSales_CappingAt = mvarNumSales_CappingAt
End Property



Public Property Let NumSales_Maximum(ByVal vData As Integer)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.NumSales_Maximum = Form1
     mvarNumSales_Maximum = vData
End Property


Public Property Get NumSales_Maximum() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NumSales_Maximum
     NumSales_Maximum = mvarNumSales_Maximum
End Property



Public Property Let NumSales_Minimum(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NumSales_Minimum = 5
    mvarNumSales_Minimum = vData
End Property


Public Property Get NumSales_Minimum() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NumSales_Minimum
    NumSales_Minimum = mvarNumSales_Minimum
End Property



Public Property Let Status(ByVal vData As Byte)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Status = Form1
     mvarStatus = vData
End Property


Public Property Get Status() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Status
     Status = mvarStatus
End Property




Public Function CommitSubClass(ByVal Class2Commit As enumPRIVFetchStatus, Optional ByVal VirtualID As Long) As Long
End Function

Public Function PopulateSubClass(Class2Get As enumPRIVFetchStatus, ByVal VirtualID As Long) As Long
End Function



Public Property Get colResellers_Payments() As colResellers_Payments
    If mvarcolResellers_Payments Is Nothing Then
        Set mvarcolResellers_Payments = New colResellers_Payments
    End If


    Set colResellers_Payments = mvarcolResellers_Payments
End Property


Public Property Set colResellers_Payments(vData As colResellers_Payments)
    Set mvarcolResellers_Payments = vData
End Property


Public Property Let SESSION(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SESSION = 5
    mvarSESSION = vData
End Property


Public Property Get SESSION() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SESSION
    SESSION = mvarSESSION
End Property



Public Property Let IDX(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IDX = 5
    mvarIDX = vData
End Property


Public Property Get IDX() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IDX
    IDX = mvarIDX
End Property





Public Property Get colResellers_InvoiceItems() As colResellers_InvoiceItems
    If mvarcolResellers_InvoiceItems Is Nothing Then
        Set mvarcolResellers_InvoiceItems = New colResellers_InvoiceItems
    End If


    Set colResellers_InvoiceItems = mvarcolResellers_InvoiceItems
End Property


Public Property Set colResellers_InvoiceItems(vData As colResellers_InvoiceItems)
    Set mvarcolResellers_InvoiceItems = vData
End Property



Public Property Get colResellers_AccountHeader() As colResellers_AccountHeader
    If mvarcolResellers_AccountHeader Is Nothing Then
        Set mvarcolResellers_AccountHeader = New colResellers_AccountHeader
    End If


    Set colResellers_AccountHeader = mvarcolResellers_AccountHeader
End Property


Public Property Set colResellers_AccountHeader(vData As colResellers_AccountHeader)
    Set mvarcolResellers_AccountHeader = vData
End Property



Public Property Get colResellers_EmailAddy() As colResellers_EmailAddy
    If mvarcolResellers_EmailAddy Is Nothing Then
        Set mvarcolResellers_EmailAddy = New colResellers_EmailAddy
    End If


    Set colResellers_EmailAddy = mvarcolResellers_EmailAddy
End Property


Public Property Set colResellers_EmailAddy(vData As colResellers_EmailAddy)
    Set mvarcolResellers_EmailAddy = vData
End Property



Public Property Get colResellers_FoneNum() As colResellers_FoneNum
    If mvarcolResellers_FoneNum Is Nothing Then
        Set mvarcolResellers_FoneNum = New colResellers_FoneNum
    End If


    Set colResellers_FoneNum = mvarcolResellers_FoneNum
End Property


Public Property Set colResellers_FoneNum(vData As colResellers_FoneNum)
    Set mvarcolResellers_FoneNum = vData
End Property
Private Sub Class_Terminate()
  Set mvarcolReseller_emailcatch = Nothing
  Set mvarcolResellers_SnailMail = Nothing
  Set mvarcolResellers_SnailMail = Nothing
  Set mvarcolResellers_Payments = Nothing
  Set mvarcolResellers_InvoiceItems = Nothing
  Set mvarcolResellers_AccountHeader = Nothing
  Set mvarcolResellers_EmailAddy = Nothing
    Set mvarcolResellers_FoneNum = Nothing
End Sub



Public Property Let ftpURLPath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ftpURLPath = 5
    mvarftpURLPath = vData
End Property


Public Property Get ftpURLPath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ftpURLPath
    ftpURLPath = mvarftpURLPath
End Property



Public Property Let ftpCostPerMB(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ftpCostPerMB = 5
    mvarftpCostPerMB = vData
End Property


Public Property Get ftpCostPerMB() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ftpCostPerMB
    ftpCostPerMB = mvarftpCostPerMB
End Property



Public Property Let ftpTotalMB(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ftpTotalMB = 5
    mvarftpTotalMB = vData
End Property


Public Property Get ftpTotalMB() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ftpTotalMB
    ftpTotalMB = mvarftpTotalMB
End Property




Public Property Let ftpBasePath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ftpBasePath = 5
    mvarftpBasePath = vData
End Property


Public Property Get ftpBasePath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ftpBasePath
    ftpBasePath = mvarftpBasePath
End Property



Public Property Let ftpPort(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ftpPort = 5
    mvarftpPort = vData
End Property


Public Property Get ftpPort() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ftpPort
    ftpPort = mvarftpPort
End Property



Public Property Let ftpPassword(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ftpPassword = 5
    mvarftpPassword = vData
End Property


Public Property Get ftpPassword() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ftpPassword
    ftpPassword = mvarftpPassword
End Property



Public Property Let ftpUsername(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ftpUsername = 5
    mvarftpUsername = vData
End Property


Public Property Get ftpUsername() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ftpUsername
    ftpUsername = mvarftpUsername
End Property

Public Property Let ftpProxy(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ftpProxy = 5
    mvarftpProxy = vData
End Property


Public Property Get ftpProxy() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ftpProxy
    ftpProxy = mvarftpProxy
End Property


Public Property Let ftpHostName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ftpHostName = 5
    mvarftpHostName = vData
End Property


Public Property Get ftpHostName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ftpHostName
    ftpHostName = mvarftpHostName
End Property



Public Property Let cTaxExemptCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.cTaxExemptCode = 5
    mvarcTaxExemptCode = vData
End Property


Public Property Get cTaxExemptCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.cTaxExemptCode
    cTaxExemptCode = mvarcTaxExemptCode
End Property



Public Property Let cTaxCountry(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.cTaxCountry = 5
    mvarcTaxCountry = vData
End Property


Public Property Get cTaxCountry() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.cTaxCountry
    cTaxCountry = mvarcTaxCountry
End Property



Public Property Let cTaxCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.cTaxCode = 5
    mvarcTaxCode = vData
End Property


Public Property Get cTaxCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.cTaxCode
    cTaxCode = mvarcTaxCode
End Property



Public Property Let bTaxMode(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bTaxMode = 5
    mvarbTaxMode = vData
End Property


Public Property Get bTaxMode() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bTaxMode
    bTaxMode = mvarbTaxMode
End Property



Public Property Let MISCFee(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MISCFee = 5
    mvarMISCFee = vData
End Property


Public Property Get MISCFee() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MISCFee
    MISCFee = mvarMISCFee
End Property



Public Property Let Comment(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Comment = 5
    mvarComment = vData
End Property


Public Property Set Comment(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Comment = Form1
    Set mvarComment = vData
End Property


Public Property Get Comment() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Comment
    If IsObject(mvarComment) Then
        Set Comment = mvarComment
    Else
        Comment = mvarComment
    End If
End Property



Public Property Let AssistanceManager_SysopID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AssistanceManager_SysopID = 5
    mvarAssistanceManager_SysopID = vData
End Property


Public Property Get AssistanceManager_SysopID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AssistanceManager_SysopID
    AssistanceManager_SysopID = mvarAssistanceManager_SysopID
End Property



Public Property Let Manager_SysopID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Manager_SysopID = 5
    mvarManager_SysopID = vData
End Property


Public Property Get Manager_SysopID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Manager_SysopID
    Manager_SysopID = mvarManager_SysopID
End Property



Public Property Let Manager(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Manager = 5
    mvarManager = vData
End Property


Public Property Set Manager(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Manager = Form1
    Set mvarManager = vData
End Property


Public Property Get Manager() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Manager
    If IsObject(mvarManager) Then
        Set Manager = mvarManager
    Else
        Manager = mvarManager
    End If
End Property



Public Property Let Icon(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Icon = 5
    mvarIcon = vData
End Property


Public Property Get Icon() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Icon
    Icon = mvarIcon
End Property



Public Property Let LogoURL(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LogoURL = 5
    mvarLogoURL = vData
End Property


Public Property Get LogoURL() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LogoURL
    LogoURL = mvarLogoURL
End Property



Public Property Let JoiningFee(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.JoiningFee = 5
    mvarJoiningFee = vData
End Property


Public Property Get JoiningFee() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.JoiningFee
    JoiningFee = mvarJoiningFee
End Property



Public Property Let PreviousCycle(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PreviousCycle = 5
    mvarPreviousCycle = vData
End Property


Public Property Get PreviousCycle() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PreviousCycle
    PreviousCycle = mvarPreviousCycle
End Property



Public Property Let NextCycle(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NextCycle = 5
    mvarNextCycle = vData
End Property


Public Property Get NextCycle() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NextCycle
    NextCycle = mvarNextCycle
End Property



Public Property Let CreatedBy_SysopID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CreatedBy_SysopID = 5
    mvarCreatedBy_SysopID = vData
End Property


Public Property Get CreatedBy_SysopID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CreatedBy_SysopID
    CreatedBy_SysopID = mvarCreatedBy_SysopID
End Property



Public Property Let NoSub(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NoSub = 5
    mvarNoSub = vData
End Property


Public Property Get NoSub() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NoSub
    NoSub = mvarNoSub
End Property



Public Property Let ACN(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ACN = 5
    mvarACN = vData
End Property


Public Property Get ACN() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ACN
    ACN = mvarACN
End Property



Public Property Let SysopID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SysopID = 5
    mvarSysopID = vData
End Property


Public Property Get SysopID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SysopID
    SysopID = mvarSysopID
End Property



Public Property Let Subscribed(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Subscribed = 5
    mvarSubscribed = vData
End Property


Public Property Get Subscribed() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Subscribed
    Subscribed = mvarSubscribed
End Property



Public Property Let ABN(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ABN = 5
    mvarABN = vData
End Property


Public Property Get ABN() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ABN
    ABN = mvarABN
End Property



Public Property Let CreationDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CreationDate = 5
    mvarCreationDate = vData
End Property


Public Property Get CreationDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CreationDate
    CreationDate = mvarCreationDate
End Property



Public Property Let Realm(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Realm = 5
    mvarRealm = vData
End Property


Public Property Get Realm() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Realm
    Realm = mvarRealm
End Property



Public Property Let BriefDesc(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BriefDesc = 5
    mvarBriefDesc = vData
End Property


Public Property Set BriefDesc(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.BriefDesc = Form1
    Set mvarBriefDesc = vData
End Property


Public Property Get BriefDesc() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BriefDesc
    If IsObject(mvarBriefDesc) Then
        Set BriefDesc = mvarBriefDesc
    Else
        BriefDesc = mvarBriefDesc
    End If
End Property



Public Property Let Description(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Description = 5
    mvarDescription = vData
End Property


Public Property Get Description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    Description = mvarDescription
End Property



Public Property Let VirtualID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VirtualID = 5
    mvarVirtualID = vData
End Property


Public Property Get VirtualID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VirtualID
    VirtualID = mvarVirtualID
End Property



Public Property Let AgencyID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AgencyID = 5
    mvarAgencyID = vData
End Property


Public Property Get AgencyID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AgencyID
    AgencyID = mvarAgencyID
End Property



Public Property Let RecID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RecID = 5
    mvarRecID = vData
End Property


Public Property Get RecID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RecID
    RecID = mvarRecID
End Property



