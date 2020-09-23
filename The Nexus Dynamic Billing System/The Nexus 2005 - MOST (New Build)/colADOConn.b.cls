VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colADOConn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"clsADOConn"
Attribute VB_Ext_KEY = "Member0" ,"clsADOConn"
'local variable to hold collection
Private mCol As Collection
Public Function FindConn(ByVal dbName As String) As Integer

    If Me.Count > 0 Then
        Dim lx As Long
        For lx = 1 To Me.Count
            If Me(lx).dbName = dbName Then
                FindConn = lx
                Exit For
                Exit Function
            End If
        Next
    End If
    
End Function


Public Function Clear()


    If Me.Count > 0 Then
        Dim lx As Integer
        For lx = Me.Count To 1 Step -1
            Me.Remove lx
        Next lx
    End If
    
End Function


Public Function Add(Key As String, oConn As adodb.Connection, Server As String, dbName As String, Optional sKey As String) As clsADOConn
    'create a new object
    Dim objNewMember As clsADOConn
    Set objNewMember = New clsADOConn


    'set the properties passed into the method
    objNewMember.Key = Key
    
    objNewMember.oConn = oConn
    objNewMember.Server = Server
    objNewMember.dbName = dbName
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As clsADOConn
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

