VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVirtualISP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"colVirtualISP"
'local variable(s) to hold property value(s)
Private mvarVirtualID As Long 'local copy
Private mvarPrimarySysopID As Long 'local copy
Private mvarftpHostName As String 'local copy
Private mvarftpPort As Integer 'local copy
Private mvarftpPassword As String 'local copy
Private mvarftpUsername As String 'local copy
Private mvarftpBasePath As String 'local copy
Private mvarftpISConnected As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarRecID As Long 'local copy
Private mvarAgencyID As Long 'local copy
Private mvarDescription As String 'local copy
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
Private mvarManager As String 'local copy
Private mvarMISCFee As Single 'local copy
Private mvarbTaxMode As Byte 'local copy
Private mvarcTaxCode As String 'local copy
Private mvarcTaxCountry As String 'local copy
Private mvarcTaxExemptCode As String 'local copy
Private mvarcolVirtualISP As colVirtualISP
'local variable(s) to hold property value(s)
Private mvarftpURLPath As String 'local copy
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





Public Property Get colVirtualISP() As colVirtualISP
    If mvarcolVirtualISP Is Nothing Then
        Set mvarcolVirtualISP = New colVirtualISP
    End If


    Set colVirtualISP = mvarcolVirtualISP
End Property


Public Property Set colVirtualISP(vData As colVirtualISP)
    Set mvarcolVirtualISP = vData
End Property
Private Sub Class_Terminate()
    Set mvarcolVirtualISP = Nothing
End Sub



Public Function PopulateMe(ByVal VirtualID As Long, oConn As ADODB.Connection) As Boolean

    Dim rsload As ADODB.Recordset
    
    If MySQL.OpenTable(oConn, rsload, , "select *, AES_DECRYPT(ftpPassword,'" + odb.colSalts.ReturnSalt("ftp") + "') as fPWD from virtualisp where RecID = '" & VirtualID & "'") = True Then
        If rsload.State = adStateOpen Then
            If rsload.RecordCount > 0 Then
                PopulateMe = True
                Me.ABN = NullStr(rsload!ABN)
                Me.ACN = NullStr(rsload!ACN)
                Me.AgencyID = NullStr(rsload!AgencyID, nNumber)
                Me.bTaxMode = NullStr(rsload!bTaxMode)
                Me.CreatedBy_SysopID = NullStr(rsload!CreatedBy_SysopID, nNumber)
                Me.CreationDate = NullStr(rsload!CreationDate)
                Me.cTaxCode = NullStr(rsload!cTaxCode)
                Me.cTaxCountry = NullStr(rsload!cTaxCountry)
                Me.cTaxExemptCode = NullStr(rsload!cTaxExemptCode)
                Me.Description = NullStr(rsload!Description)
                Me.ftpBasePath = NullStr(rsload!ftpBasePath)
                Me.ftpHostName = NullStr(rsload!ftpHostName)
                Me.ftpISConnected = False
                Me.ftpPassword = NullStr(rsload!fPWD)
                Me.ftpPort = NullStr(rsload!ftpPort, nNumber)
                Me.ftpUsername = NullStr(rsload!ftpUsername)
                Me.Icon = NullStr(rsload!Icon, nNumber)
                Me.JoiningFee = NullStr(rsload!JoiningFee, nNumber)
                Me.LogoURL = NullStr(rsload!LogoURL)
                Me.Manager = NullStr(rsload!Manager)
                Me.MISCFee = NullStr(rsload!MISCFee, nNumber)
                Me.NextCycle = IIf(IsNull(rsload!NextCycle), sysnow, rsload!NextCycle)
                Me.PreviousCycle = IIf(IsNull(rsload!PreviousCycle), sysnow, rsload!PreviousCycle)
                Me.NoSub = NullStr(rsload!NoSub, nNumber)
                Me.SysopID = NullStr(rsload!SysopID)
                Me.Realm = NullStr(rsload!Realm)
                Me.RecID = NullStr(rsload!RecID, nNumber)
                Me.Subscribed = NullStr(rsload!Subscribed, nNumber)
                Me.VirtualID = NullStr(rsload!VirtualID, nNumber)
                Me.ftpURLPath = NullStr(rsload!ftpURLPath)
            Else
                
                Me.ftpBasePath = "/$FileDB$/"
                Me.ftpHostName = Crypt("ë0ë.¢Õë.¢ëè.ë≤", False, "None")
                Me.ftpISConnected = False
                Me.ftpPassword = Crypt("¶’±Ÿ¶≠ïœØØª≈¡¥ØØ", False, "None")
                Me.ftpPort = 21
                Me.ftpUsername = "daemon"
                Me.ftpURLPath = "http://202.172.123.25/$FileDB$/"
                
            End If
        End If
    End If

End Function

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



Public Property Let bTaxMode(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bTaxMode = 5
    mvarbTaxMode = vData
End Property


Public Property Get bTaxMode() As Byte
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



Public Property Let Manager(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Manager = 5
    mvarManager = vData
End Property


Public Property Get Manager() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Manager
    Manager = mvarManager
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



Public Property Let ftpISConnected(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ftpISConnected = 5
    mvarftpISConnected = vData
End Property


Public Property Get ftpISConnected() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ftpISConnected
    ftpISConnected = mvarftpISConnected
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



Public Property Let ftpPort(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ftpPort = 5
    mvarftpPort = vData
End Property


Public Property Get ftpPort() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ftpPort
    ftpPort = mvarftpPort
End Property



Public Property Let ftpHostName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ftpHostname = 5
    mvarftpHostName = vData
End Property


Public Property Get ftpHostName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ftpHostname
    ftpHostName = mvarftpHostName
End Property



Public Property Let PrimarySysopID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrimarySysopID = 5
    mvarPrimarySysopID = vData
End Property


Public Property Get PrimarySysopID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrimarySysopID
    PrimarySysopID = mvarPrimarySysopID
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



