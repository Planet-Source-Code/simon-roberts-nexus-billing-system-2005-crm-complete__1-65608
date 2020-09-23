VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHoldingNodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

Private mvarDescription As String 'local copy
Private mvarselectStatement As String 'local copy
Private mvarCountStatement As String 'local copy
Private mvarIconNo As Byte 'local copy
Private mvarVirtualID As Variant 'local copy
Private mvarRecID As Double 'local copy
Private mvarSubofRecID As Double 'local copy
'local variable(s) to hold property value(s)
Private mvarCatNo As String 'local copy
'local variable(s) to hold property value(s)
Private mvarTag As String 'local copy
Public Property Let Tag(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Tag = 5
    mvarTag = vData
End Property


Public Property Get Tag() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Tag
    Tag = mvarTag
End Property



Public Property Let CatNo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CatNo = 5
    mvarCatNo = vData
End Property


Public Property Get CatNo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CatNo
    CatNo = mvarCatNo
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



Public Property Let VirtualID(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VirtualID = 5
    mvarVirtualID = vData
End Property


Public Property Set VirtualID(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.VirtualID = Form1
    Set mvarVirtualID = vData
End Property


Public Property Get VirtualID() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VirtualID
    If IsObject(mvarVirtualID) Then
        Set VirtualID = mvarVirtualID
    Else
        VirtualID = mvarVirtualID
    End If
End Property



Public Property Let IconNo(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IconNo = 5
    mvarIconNo = vData
End Property


Public Property Get IconNo() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IconNo
    IconNo = mvarIconNo
End Property



Public Property Let CountStatement(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CountStatement = 5
    mvarCountStatement = vData
End Property


Public Property Get CountStatement() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CountStatement
    CountStatement = mvarCountStatement
End Property



Public Property Let selectStatement(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.selectStatement = 5
    mvarselectStatement = vData
End Property


Public Property Get selectStatement() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.selectStatement
    selectStatement = mvarselectStatement
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



