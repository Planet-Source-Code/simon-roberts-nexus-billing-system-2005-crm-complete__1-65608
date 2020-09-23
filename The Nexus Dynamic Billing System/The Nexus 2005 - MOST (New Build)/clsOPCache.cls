VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOPCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

Private mvarSQLStatement As Variant 'local copy
Private mvardbName As String 'local copy
Private mvarfldName As String 'local copy
Private mvartblName As String 'local copy
Private mvarbNumeric As Boolean 'local copy
Private mvarvResult As Variant 'local copy
'local variable(s) to hold property value(s)
Private mvarLastSort As Date 'local copy
'local variable(s) to hold property value(s)
Private mvarsValue As String 'local copy
Private mvardValue As Double 'local copy
Public Property Let dValue(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValue = 5
    mvardValue = vData
End Property


Public Property Get dValue() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValue
    dValue = mvardValue
End Property



Public Property Let sValue(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sValue = 5
    mvarsValue = vData
End Property


Public Property Get sValue() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sValue
    sValue = mvarsValue
End Property



Public Property Let LastSort(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastSort = 5
    mvarLastSort = vData
End Property


Public Property Get LastSort() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastSort
    LastSort = mvarLastSort
End Property



Public Property Let vResult(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.vResult = 5
    mvarvResult = vData
End Property


Public Property Set vResult(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.vResult = Form1
    Set mvarvResult = vData
End Property


Public Property Get vResult() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.vResult
    If IsObject(mvarvResult) Then
        Set vResult = mvarvResult
    Else
        vResult = mvarvResult
    End If
End Property



Public Property Let bNumeric(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bNumeric = 5
    mvarbNumeric = vData
End Property


Public Property Get bNumeric() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bNumeric
    bNumeric = mvarbNumeric
End Property



Public Property Let tblName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.tblName = 5
    mvartblName = vData
End Property


Public Property Get tblName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.tblName
    tblName = mvartblName
End Property



Public Property Let fldName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fldName = 5
    mvarfldName = vData
End Property


Public Property Get fldName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fldName
    fldName = mvarfldName
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



Public Property Let SQLStatement(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SQLStatement = 5
    mvarSQLStatement = vData
End Property


Public Property Set SQLStatement(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.SQLStatement = Form1
    Set mvarSQLStatement = vData
End Property


Public Property Get SQLStatement() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SQLStatement
    If IsObject(mvarSQLStatement) Then
        Set SQLStatement = mvarSQLStatement
    Else
        SQLStatement = mvarSQLStatement
    End If
End Property



