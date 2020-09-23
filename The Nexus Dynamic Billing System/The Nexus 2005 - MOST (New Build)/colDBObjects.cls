VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colDBObjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"clsDBObjects"
Attribute VB_Ext_KEY = "Member0" ,"clsDBObjects"
'local variable to hold collection
Private mCol As Collection
'local variable(s) to hold property value(s)
Private mvardbTables As Integer 'local copy
Private mvardbFields As Long 'local copy
Public Function FindField(ByVal tblName As String, ByVal fldName As String) As Long

    If Me.Count > 0 Then
        Dim lx As Long
        For lx = Me.Count To 1 Step -1
            If LCase(Me(lx).Tablename) = LCase(tblName) And LCase(Me(lx).FieldName) = LCase(fldName) Then
                FindField = lx
                Exit For
            End If
        Next
    End If
    
End Function

Public Function FindTable(ByVal tblName As String) As Long

    If Me.Count > 0 Then
        Dim lx As Long
        For lx = Me.Count To 1 Step -1
            If LCase(Me(lx).Tablename) = LCase(tblName) Then
                FindTable = lx
                Exit For
            End If
        Next
    End If
    
End Function

Public Property Let dbFields(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dbFields = 5
    mvardbFields = vData
End Property


Public Property Get dbFields() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dbFields
    dbFields = mvardbFields
End Property



Public Property Let dbTables(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dbTables = 5
    mvardbTables = vData
End Property


Public Property Get dbTables() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dbTables
    dbTables = mvardbTables
End Property




Public Function CheckValue(Value As Variant, refField As Long, ByRef refReturn As Variant) As Integer

    On Error Resume Next
    
    Select Case Me(refField).TypeID
    Case adEmpty '0 This data type indicates that no value was specified (DBTYPE_EMPTY).
        CheckValue = True
    Case adSmallInt Or adInteger Or adSingle Or adDouble Or adNumeric Or adBoolean Or adTinyInt Or adUnsignedTinyInt Or adUnsignedSmallInt Or adUnsignedInt Or adBigInt Or adUnsignedBigInt Or adDecimal Or adNumeric
        If IsNumeric(Value) = True Then CheckValue = True
    Case adCurrency '6 A data type indicates a currency value (DBTYPE_CY). Currency is a fixed-point number with 4 digits to the right of the decimal point. It is stored in an 8-byte signed integer scaled by 10,000. This data type is not supported by the OLE DB Provider.
        If CCur(Value) <> 0 Or Value = 0 Then CheckValue = True
    Case adDate, adDBDate, adDBTime, adDBTimeStamp  '7 This data type indicates a date value stored as a Double, the whole part of which is the number of days since December 30, 1899, and the fractional part of which is the fraction of a day. This data type is not supported by the OLE DB Provider.
        If IsDate(Value) = True Then
            CheckValue = True
            Select Case Me(refField).TypeID
            Case adDate
                Select Case Me(refField).DefinedSize
                Case 6
                    refReturn = Format(CDate(Value), "YYYY-MM-DD")
                Case Else
                    refReturn = Format(CDate(Value), "YYYY-MM-DD Hh:Mm:Ss")
                End Select
                CheckValue = 1
            Case adDBTime
                refReturn = Format(CDate(Value), "Long Time")
                CheckValue = 1
            Case adDBDate
                refReturn = Format(CDate(Value), "YYYY-MM-DD")
                CheckValue = 1
            End Select
        End If
    Case aRegDBsTR '8 This data type indicates a null-terminated Unicode character string (DBTYPE_BSTR). This data type is not supported by the OLE DB Provider.
        CheckValue = True
    Case adIDispatch '9 This data type indicates a pointer to an IDispatch interface on an OLE object (DBTYPE_IDISPATCH). This data type is not supported by the OLE DB Provider.
        CheckValue = True
    Case adError '10 This data type indicates a 32-bit error code (DBTYPE_ERROR). This data type is not supported by the OLE DB Provider.
        CheckValue = True
    Case adVariant '12 This data type indicates an Automation variant (DBTYPE_VARIANT). This data type is not supported by the OLE DB Provider.
        CheckValue = True
    Case adIUnknown '13 This data type indicates a pointer to an IUnknown interface on an OLE object (DBTYPE_IUNKNOWN). This data type is not supported by the OLE DB Provider.
        CheckValue = True
    Case adGUID '72 This data type indicates a globally unique identifier or GUID (DBTYPE_GUID). This data type is not supported by the OLE DB Provider.
        CheckValue = True
    Case adBinary '128 This data type indicates fixed-length binary data (DBTYPE_BYTES).
        CheckValue = True
    Case adChar, adWChar, adVarChar, adLongVarChar, adVarWChar, adLongVarWChar
        If Value <> "" Then CheckValue = True
    Case adVarBinary '204 This data type indicates variable-length binary data (DBTYPE_BYTES).
        CheckValue = True
    Case adLongVarBinary '205 This data type indicates a long binary value.
        CheckValue = True
    Case Else
        CheckValue = True
    End Select

End Function

Public Function Clear()

    If Me.Count > 0 Then
        Dim ix As Long
        For ix = Me.Count To 1 Step -1
        
            Me.Remove ix
        
        Next
        
    End If

End Function


Public Function Add(Key As String, SecLevel As Byte, bNull As Boolean, bID As Boolean, Extra As String, vDefault As String, IndexKey As String, dbName As String, Tablename As String, FieldName As String, oTBtvIndex As Long, DefinedSize As Long, NumericScale As Long, Prescision As Long, Status As Long, TypeID As Long, TypeString As String, Attributes As Long, Optional sKey As String) As clsDBObjects
    'create a new object
    Dim objNewMember As clsDBObjects
    Set objNewMember = New clsDBObjects


    'set the properties passed into the method
    objNewMember.Key = Key
    objNewMember.SecLevel = SecLevel
    objNewMember.bNull = bNull
    objNewMember.bID = bID
    objNewMember.Extra = Extra
    objNewMember.vDefault = vDefault
    objNewMember.IndexKey = IndexKey
    objNewMember.Tablename = Tablename
    objNewMember.FieldName = FieldName
    objNewMember.oTBtvIndex = oTBtvIndex
    objNewMember.DefinedSize = DefinedSize
    objNewMember.NumericScale = NumericScale
    objNewMember.Prescision = Prescision
    objNewMember.Status = Status
    objNewMember.TypeID = TypeID
    objNewMember.TypeString = TypeString
    objNewMember.Attributes = Attributes
    objNewMember.dbName = dbName
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

    If Len(FieldName) > 0 Then
        Me.dbFields = Me.dbFields + 1
    End If
    
End Function

Public Property Get Item(vntIndexKey As Variant) As clsDBObjects
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

