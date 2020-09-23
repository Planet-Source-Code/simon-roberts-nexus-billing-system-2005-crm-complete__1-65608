VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChannels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvarDescription As String 'local copygf
Private mvarRecID As Double 'local copy
Private mvarSubofRecID As Double 'local copy
Private mvarGetSQL As Variant 'local copy
Private mvarCountSQL As Variant 'local copy
Private mvarIcon As Byte 'local copy
Public Property Let Icon(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Icon = 5
    mvarIcon = vData
End Property


Public Property Get Icon() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Icon
    Icon = mvarIcon
End Property



Public Property Let CountSQL(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CountSQL = 5
    mvarCountSQL = vData
End Property


Public Property Set CountSQL(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.CountSQL = Form1
    Set mvarCountSQL = vData
End Property


Public Property Get CountSQL() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CountSQL
    If IsObject(mvarCountSQL) Then
        Set CountSQL = mvarCountSQL
    Else
        CountSQL = mvarCountSQL
    End If
End Property



Public Property Let GetSQL(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GetSQL = 5
    mvarGetSQL = vData
End Property


Public Property Set GetSQL(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.GetSQL = Form1
    Set mvarGetSQL = vData
End Property


Public Property Get GetSQL() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GetSQL
    If IsObject(mvarGetSQL) Then
        Set GetSQL = mvarGetSQL
    Else
        GetSQL = mvarGetSQL
    End If
End Property



Public Property Let SubofRecID(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SubofRecID = 5
    mvarSubofRecID = vData
End Property


Public Property Get SubofRecID() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SubofRecID
    SubofRecID = mvarSubofRecID
End Property



Public Property Let RecID(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RecID = 5
    mvarRecID = vData
End Property


Public Property Get RecID() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RecID
    RecID = mvarRecID
End Property



Public Property Let Description(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Description = 5
    mvarDescription = vData
End Property


Public Property Get Description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    Description = mvarDescription
End Property



