VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colFilesGrouping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"clsFilesGrouping"
Attribute VB_Ext_KEY = "Member0" ,"clsFilesGrouping"
'local variable to hold collection
Private mCol As Collection
Public Sub Clear(Optional VirtualID As Long = 0, Optional ClientID As Long = 0, Optional SESSION As String)
End Sub

Public Function FindIDX(sKey As String) As Long
End Function


Public Function Add(Key As String, FileAllocation As enumFileAllocation, GroupingFolder As String, IDX As Long, SESSION As String, SubFolder As String, VirtualID As Long, ClientID As Long, VendorID As Long, RefundID As Long, SysopID As Long, CommsID As Long, SerializedFolder As String, LastQuery As Object, Optional sKey As String) As clsFilesGrouping
    'create a new object
    Dim objNewMember As clsFilesGrouping
    Set objNewMember = New clsFilesGrouping


    'set the properties passed into the method
    objNewMember.Key = Key
    If IsObject(FileAllocation) Then
        Set objNewMember.FileAllocation = FileAllocation
    Else
        objNewMember.FileAllocation = FileAllocation
    End If
    objNewMember.GroupingFolder = GroupingFolder
    objNewMember.IDX = IDX
    objNewMember.SESSION = SESSION
    objNewMember.SubFolder = SubFolder
    objNewMember.VirtualID = VirtualID
    objNewMember.ClientID = ClientID
    objNewMember.VendorID = VendorID
    objNewMember.RefundID = RefundID
    objNewMember.SysopID = SysopID
    objNewMember.CommsID = CommsID
    objNewMember.SerializedFolder = SerializedFolder
    If IsObject(LastQuery) Then
        Set objNewMember.LastQuery = LastQuery
    Else
        objNewMember.LastQuery = LastQuery
    End If
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As clsFilesGrouping
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

