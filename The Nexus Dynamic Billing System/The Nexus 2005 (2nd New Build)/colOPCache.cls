VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colOPCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"clsOPCache"
Attribute VB_Ext_KEY = "Member0" ,"clsOPCache"
'local variable to hold collection
Private mCol As Collection
Public Function FindBySQL(ByVal vSQL As String) As String

    If Me.Count > 0 Then
    
        Dim dx As Long
        For dx = Me.Count To 1 Step -1
        
            If Me(dx).SQLStatement = vSQL Then
                FindBySQL = Me(dx).Key
                Me(dx).LastSort = Now
                Exit For
            End If
        
        Next dx
            
        For dx = Me.Count To 1 Step -1
        
            If DateDiff("n", Me(dx).LastSort, sysnow) >= 120 Then
                If Me(dx).Key <> FindBySQL Then
                    Me.Remove dx
                    gSleep
                End If
            End If
        
        Next dx
            
        
    End If
End Function


Public Function Add(Key As String, SQLStatement As Variant, dbName As String, fldName As String, tblName As String, bNumeric As Boolean, sValue As String, dValue As Double, Optional sKey As String) As clsOPCache
    'create a new object
    Dim objNewMember As clsOPCache
    Set objNewMember = New clsOPCache


    'set the properties passed into the method
    objNewMember.Key = Key
    
    objNewMember.LastSort = Now
    If IsObject(SQLStatement) Then
        Set objNewMember.SQLStatement = SQLStatement
    Else
        objNewMember.SQLStatement = SQLStatement
    End If
    objNewMember.dbName = dbName
    objNewMember.fldName = fldName
    objNewMember.tblName = tblName
    objNewMember.bNumeric = bNumeric
    objNewMember.dValue = dValue
    objNewMember.sValue = sValue
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As clsOPCache
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

