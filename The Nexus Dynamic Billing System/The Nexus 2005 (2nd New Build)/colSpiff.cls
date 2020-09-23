VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colSpiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"clsBonuses"
Attribute VB_Ext_KEY = "Member0" ,"clsBonuses"
'local variable to hold collection
Private mCol As Collection
Public Sub Clear(VirtualID As Long)

    Dim ix As Double
    Dim objNewMember As clsBonuses
    
    For ix = Me.Count To 1 Step -1
        If VirtualID = -1 Then
            Me.Remove ix
        Else
            Set objNewMember = Me(ix)
            If objNewMember.VirtualID = VirtualID Then
            
            Else
                Me.Remove ix
            End If
        End If
    Next
    

End Sub


Public Function Add(Key As String, enumSource As spiffSource, UnitType As Integer, Min As Integer, Max As Long, UnitsAwarded As Integer, VirtualID As Long, ptRecID As Long, Optional sKey As String) As clsBonuses
    'create a new object
    Dim objNewMember As clsBonuses
    Set objNewMember = New clsBonuses


    'set the properties passed into the method
    objNewMember.Key = Key
    objNewMember.enumSource = enumSource
    objNewMember.UnitType = UnitType
    objNewMember.Min = Min
    objNewMember.Max = Max
    objNewMember.UnitsAwarded = UnitsAwarded
    objNewMember.VirtualID = VirtualID
    objNewMember.ptRecID = ptRecID
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As clsBonuses
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrueving the number of elements in the
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

