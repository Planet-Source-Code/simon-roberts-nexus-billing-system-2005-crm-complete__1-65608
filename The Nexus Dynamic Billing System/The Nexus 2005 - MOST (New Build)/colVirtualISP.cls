VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colVirtualISP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"clsVirtualISP"
Attribute VB_Ext_KEY = "Member0" ,"clsVirtualISP"
'local variable to hold collection
Private mCol As Collection

Public Function Add(VirtualID As Long, PrimarySysopID As Long, ftpHostName As String, ftpPort As Integer, ftpPassword As String, ftpUsername As String, ftpBasePath As String, ftpISConnected As Boolean, RecID As Long, AgencyID As Long, Description As String, Realm As String, CreationDate As Date, ABN As String, Subscribed As Long, SysopID As Long, ACN As String, NoSub As Long, CreatedBy_SysopID As Long, NextCycle As Date, PreviousCycle As Date, JoiningFee As Single, LogoURL As String, Icon As Integer, Manager As String, MISCFee As Single, bTaxMode As Byte, cTaxCode As String, cTaxCountry As String, cTaxExemptCode As String, colVirtualISP As colVirtualISP, Optional sKey As String) As clsVirtualISP
    'create a new object
    Dim objNewMember As clsVirtualISP
    Set objNewMember = New clsVirtualISP


    'set the properties passed into the method
    objNewMember.VirtualID = VirtualID
    objNewMember.PrimarySysopID = PrimarySysopID
    objNewMember.ftpHostName = ftpHostName
    objNewMember.ftpPort = ftpPort
    objNewMember.ftpPassword = ftpPassword
    objNewMember.ftpUsername = ftpUsername
    objNewMember.ftpBasePath = ftpBasePath
    objNewMember.ftpISConnected = ftpISConnected
    objNewMember.RecID = RecID
    objNewMember.AgencyID = AgencyID
    objNewMember.Description = Description
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
    objNewMember.Manager = Manager
    objNewMember.MISCFee = MISCFee
    objNewMember.bTaxMode = bTaxMode
    objNewMember.cTaxCode = cTaxCode
    objNewMember.cTaxCountry = cTaxCountry
    objNewMember.cTaxExemptCode = cTaxExemptCode
    Set objNewMember.colVirtualISP = colVirtualISP
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As clsVirtualISP
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

