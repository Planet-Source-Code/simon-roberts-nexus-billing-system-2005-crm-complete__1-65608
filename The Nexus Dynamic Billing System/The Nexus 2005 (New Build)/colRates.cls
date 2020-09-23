VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"clsRates"
Attribute VB_Ext_KEY = "Member0" ,"clsRates"
'local variable to hold collection
Private mCol As Collection
Public Sub Clear()

    Dim ix As Long
    If Me.Count > 0 Then
        For ix = Me.Count To 1 Step -1
            Me.Remove ix
        Next
    End If
    
End Sub

Public Function RangeExists(Min As Long, Max As Long) As Boolean
End Function


Public Function Add(Key As String, Optional ClassID As Double, Optional ptRecID As Double, Optional Min As Integer, Optional Max As Integer, Optional Rate As Currency, Optional Percentile As Single, Optional sKey As String) As clsRates
    'create a new object
    Dim objNewMember As clsRates
    Set objNewMember = New clsRates


    'set the properties passed into the method
    objNewMember.Key = Key
    objNewMember.ClassID = ClassID
    objNewMember.ptRecID = ptRecID
    objNewMember.Min = Min
    objNewMember.Max = Max
    objNewMember.Rate = Rate
    objNewMember.Percentile = Percentile
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As clsRates
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

