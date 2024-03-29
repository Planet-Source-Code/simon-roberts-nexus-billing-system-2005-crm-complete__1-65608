VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colSalts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"clsSalts"
Attribute VB_Ext_KEY = "Member0" ,"clsSalts"
'local variable to hold collection
Private mCol As Collection
Public Function ReturnSalt(ByVal sName As String, Optional ByRef Index As Long) As String
    
    If Me.Count > 0 Then
        If Index = 0 Then
            Dim iCnt As Long
            For iCnt = 1 To Me.Count
                If LCase(Me(iCnt).sName) = LCase(sName) Then
                    Index = iCnt
                    Exit For
                End If
            Next
            If Index = 0 Then
                ReturnSalt = "p0st41"
            Else
                ReturnSalt = Me(Index).sSalt
            End If
        Else
            ReturnSalt = Me(Index).sSalt
        End If
    Else
        ReturnSalt = "p0st41"
    End If
    
End Function

Public Function Clear()

    If Me.Count > 0 Then
        Dim kLoop As Long
        For kLoop = Me.Count To 1 Step -1
            Me.Remove kLoop
        Next
    End If
    
    gSleep
    
End Function

Public Function Add(Key As String, sName As String, sSalt As String, Optional sKey As String) As clsSalts
    'create a new object
    Dim objNewMember As clsSalts
    Set objNewMember = New clsSalts


    'set the properties passed into the method
    objNewMember.Key = Key
    objNewMember.sName = sName
    objNewMember.sSalt = sSalt
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As clsSalts
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

