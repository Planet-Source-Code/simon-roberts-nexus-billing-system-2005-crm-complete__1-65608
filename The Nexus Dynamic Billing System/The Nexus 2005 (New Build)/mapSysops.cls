VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mapSysops"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"clsSysops"
Attribute VB_Ext_KEY = "Member0" ,"clsSysops"
'local variable to hold collection
Private mCol As Collection
Public Sub Clear()
End Sub


Public Function Add(Key As String, RecID As Double, Username As String, PublicKey As String, Description As String, SecurityLevel As Byte, Checked As Boolean, VirtualID As Double, AgencyID As Double, bMaintain As Boolean, bVISP As Boolean, bCreateSysop As Boolean, bPrimary As Boolean, bAgency As Boolean, bTemplates As Boolean, Firstname As String, Surname As String, NextCycle As Date, PreviousCycle As Date, bRecievables As Boolean, bInvoice As Boolean, bExpenditure As Boolean, bHoldings As Boolean, bComm As Boolean, bRefund As Boolean, bAddCust As Boolean, bOwnership As Boolean, bAccSettings As Boolean, bVendors As Boolean, Optional sKey As String) As clsSysops
    'create a new object
    Dim objNewMember As clsSysops
    Set objNewMember = New clsSysops


    'set the properties passed into the method
    objNewMember.Key = Key
    objNewMember.RecID = RecID
    objNewMember.Username = Username
    objNewMember.PublicKey = PublicKey
    objNewMember.Description = Description
    objNewMember.SecurityLevel = SecurityLevel
    objNewMember.Checked = Checked
    objNewMember.VirtualID = VirtualID
    objNewMember.AgencyID = AgencyID
    objNewMember.bMaintain = bMaintain
    objNewMember.bVISP = bVISP
    objNewMember.bCreateSysop = bCreateSysop
    objNewMember.bPrimary = bPrimary
    objNewMember.bAgency = bAgency
    objNewMember.bTemplates = bTemplates
    objNewMember.Firstname = Firstname
    objNewMember.Surname = Surname
    objNewMember.NextCycle = NextCycle
    objNewMember.PreviousCycle = PreviousCycle
    objNewMember.bRecievables = bRecievables
    objNewMember.bInvoice = bInvoice
    objNewMember.bExpenditure = bExpenditure
    objNewMember.bHoldings = bHoldings
    objNewMember.bComm = bComm
    objNewMember.bRefund = bRefund
    objNewMember.bAddCust = bAddCust
    objNewMember.bOwnership = bOwnership
    objNewMember.bAccSettings = bAccSettings
    objNewMember.bVendors = bVendors
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As clsSysops
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

