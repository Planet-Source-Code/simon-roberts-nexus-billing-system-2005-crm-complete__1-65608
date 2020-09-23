VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDBObjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvarTableName As String 'local copy
Private mvarFieldName As String 'local copy
Private mvaroTBtvIndex As Long 'local copy
Private mvarDefinedSize As Long 'local copy
Private mvarNumericScale As Long 'local copy
Private mvarPrescision As Long 'local copy
Private mvarStatus As Long 'local copy
Private mvarTypeID As Long 'local copy
Private mvarTypeString As String 'local copy
Private mvarAttributes As Long 'local copy
'local variable(s) to hold property value(s)
Private mvardbName As String 'local copy
'local variable(s) to hold property value(s)
Private mvarbNULL As Boolean 'local copy
Private mvarIndexKey As String 'local copy
Private mvarvDefault As String 'local copy
Private mvarExtra As String 'local copy
Private mvarbID As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarSecLevel As Byte 'local copy
Public Property Let SecLevel(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SecLevel = 5
    mvarSecLevel = vData
End Property


Public Property Get SecLevel() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SecLevel
    SecLevel = mvarSecLevel
End Property



Public Property Let bID(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bID = 5
    mvarbID = vData
End Property


Public Property Get bID() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bID
    bID = mvarbID
End Property



Public Property Let Extra(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Extra = 5
    mvarExtra = vData
End Property


Public Property Get Extra() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Extra
    Extra = mvarExtra
End Property



Public Property Let vDefault(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.vDefault = 5
    mvarvDefault = vData
End Property


Public Property Get vDefault() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.vDefault
    vDefault = mvarvDefault
End Property



Public Property Let IndexKey(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IndexKey = 5
    mvarIndexKey = vData
End Property


Public Property Get IndexKey() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IndexKey
    IndexKey = mvarIndexKey
End Property



Public Property Let bNull(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bNULL = 5
    mvarbNULL = vData
End Property


Public Property Get bNull() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bNULL
    bNull = mvarbNULL
End Property



Public Property Let dbName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dbName = 5
    mvardbName = vData
End Property


Public Property Get dbName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dbName
    dbName = mvardbName
End Property



Public Property Let Attributes(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Attributes = 5
    mvarAttributes = vData
End Property


Public Property Get Attributes() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Attributes
    Attributes = mvarAttributes
End Property



Public Property Let TypeString(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TypeString = 5
    mvarTypeString = vData
End Property


Public Property Get TypeString() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TypeString
    TypeString = mvarTypeString
End Property



Public Property Let TypeID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TypeID = 5
    mvarTypeID = vData
End Property


Public Property Get TypeID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TypeID
    TypeID = mvarTypeID
End Property



Public Property Let Status(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Status = 5
    mvarStatus = vData
End Property


Public Property Get Status() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Status
    Status = mvarStatus
End Property



Public Property Let Prescision(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Prescision = 5
    mvarPrescision = vData
End Property


Public Property Get Prescision() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Prescision
    Prescision = mvarPrescision
End Property



Public Property Let NumericScale(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NumericScale = 5
    mvarNumericScale = vData
End Property


Public Property Get NumericScale() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NumericScale
    NumericScale = mvarNumericScale
End Property



Public Property Let DefinedSize(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DefinedSize = 5
    mvarDefinedSize = vData
End Property


Public Property Get DefinedSize() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DefinedSize
    DefinedSize = mvarDefinedSize
End Property



Public Property Let oTBtvIndex(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.oTBtvIndex = 5
    mvaroTBtvIndex = vData
End Property


Public Property Get oTBtvIndex() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.oTBtvIndex
    oTBtvIndex = mvaroTBtvIndex
End Property



Public Property Let FieldName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FieldName = 5
    mvarFieldName = vData
End Property


Public Property Get FieldName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FieldName
    FieldName = mvarFieldName
End Property



Public Property Let Tablename(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TableName = 5
    mvarTableName = vData
End Property


Public Property Get Tablename() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TableName
    Tablename = mvarTableName
End Property



