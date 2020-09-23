VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvariAction As Integer 'local copy
Private mvarxNode As Node 'local copy
Private mvarSQLStatement As String 'local copy
Private mvarSQLCount As String 'local copy
Private mvarIndex As Long 'local copy
Public Property Let Index(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Index = 5
    mvarIndex = vData
End Property


Public Property Get Index() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Index
    Index = mvarIndex
End Property



Public Property Let SQLCount(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SQLCount = 5
    mvarSQLCount = vData
End Property


Public Property Get SQLCount() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SQLCount
    SQLCount = mvarSQLCount
End Property



Public Property Let SQLStatement(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SQLStatement = 5
    mvarSQLStatement = vData
End Property


Public Property Get SQLStatement() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SQLStatement
    SQLStatement = mvarSQLStatement
End Property



Public Property Set xNode(ByVal vData As Node)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.xNode = Form1
    Set mvarxNode = vData
End Property


Public Property Get xNode() As Node
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.xNode
    Set xNode = mvarxNode
End Property



Public Property Let iAction(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iAction = 5
    mvariAction = vData
End Property


Public Property Get iAction() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iAction
    iAction = mvariAction
End Property



