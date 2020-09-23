VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsADOConn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvaroConn As ADODB.Connection 'local copy
Private mvarServer As String 'local copy
Private mvardbName As String 'local copy
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



Public Property Let Server(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Server = 5
    mvarServer = vData
End Property


Public Property Get Server() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Server
    Server = mvarServer
End Property



Public Property Let oConn(ByVal vData As ADODB.Connection)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.oConn = 5
    Set mvaroConn = vData
End Property


Public Property Get oConn() As ADODB.Connection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.oConn
    Set oConn = mvaroConn
End Property



