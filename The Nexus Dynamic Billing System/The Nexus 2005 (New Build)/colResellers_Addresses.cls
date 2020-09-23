VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colResellers_Addresses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"clsResellers_Addresses"
Attribute VB_Ext_KEY = "Member0" ,"clsResellers_Addresses"
'local variable to hold collection
Private mCol As Collection
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
Public Function Add(Key As String, RecID As Long, visp_RecID As Long, FlagID As Integer, DateAdded As Date, ContactName As String, Street1 As String, Street2 As String, Country As String, State As String, Suburb As String, PostCode As String, Cancelled As Integer, Checked As Integer, PhotoURL As String, Status As enumLineStatus, SESSION As String, Optional sKey As String) As clsResellers_Addresses
    'create a new object
    Dim objNewMember As clsResellers_Addresses
    Set objNewMember = New clsResellers_Addresses


    'set the properties passed into the method
    objNewMember.Key = Key
    objNewMember.RecID = RecID
    objNewMember.visp_RecID = visp_RecID
    objNewMember.IDX = Me.Count + 1
    objNewMember.SESSION = SESSION
    objNewMember.FlagID = FlagID
    objNewMember.DateAdded = DateAdded
    objNewMember.ContactName = ContactName
    objNewMember.Cancelled = Cancelled
    objNewMember.Checked = Checked
    objNewMember.PhotoURL = PhotoURL
    If IsObject(Status) Then
        Set objNewMember.Status = Status
    Else
        objNewMember.Status = Status
    End If
    objNewMember.Street1 = Street1
    objNewMember.Street2 = Street2
    objNewMember.Country = Country
    objNewMember.State = State
    objNewMember.Suburb = Suburb
    objNewMember.PostCode = PostCode
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As clsResellers_Addresses
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

