VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oFlagsExt"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"colFlagsMAP"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public oDB As New clsDB


'local variable(s) to hold property value(s)
Private mvarConnectionString As String 'local copy
Private mvarCommandTimeout As Integer 'local copy
Private mvaroConn As ADODB.Connection  'local copy
Private mvarcolFlagsMAP As colFlagsMAP
Private mvarFlagID As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarucMain As UserControl  'local copy
'local variable(s) to hold property value(s)
Private mvarVirtualID As Long 'local copy
Private mvarSysopID As Long 'local copy
Public Property Let SysopID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SysopID = 5
    mvarSysopID = vData
End Property


Public Property Get SysopID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SysopID
    SysopID = mvarSysopID
End Property



Public Property Let VirtualID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VirtualID = 5
    mvarVirtualID = vData
End Property


Public Property Get VirtualID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VirtualID
    VirtualID = mvarVirtualID
End Property



Public Property Let ucMain(ByVal vData As UserControl)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ucMain = 5
    mvarucMain = vData
End Property


Public Property Set ucMain(ByVal vData As UserControl)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ucMain = Form1
    Set mvarucMain = vData
End Property


Public Property Get ucMain() As UserControl
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ucMain
    If IsObject(UserControl) Then
        Set ucMain = UserControl
    Else
        ucMain = UserControl
    End If
End Property



Public Function Hook() As Boolean

    If Me.ConnectionString = "" Or Me.CommandTimeout = "" Then
        Hook = False
    End If
    
    
    If Me.oConn.State = adStateOpen Then
        Me.oConn.Close
    End If
    
    On Error Resume Next
    
    mvaroConn.Open Me.ConnectionString
    
    If mvaroConn.State = adStateOpen Then
        Hook = True
        
        Me.colFlagsMAP.Clear
        
        
    
    End If
    
        
End Function

Public Property Let FlagID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FlagID = 5
    mvarFlagID = vData
End Property


Public Property Get FlagID() As Long
Attribute FlagID.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FlagID
    FlagID = mvarFlagID
End Property





Public Property Get colFlagsMAP() As colFlagsMAP
    If mvarcolFlagsMAP Is Nothing Then
        Set mvarcolFlagsMAP = New colFlagsMAP
    End If


    Set colFlagsMAP = mvarcolFlagsMAP
End Property


Public Property Set colFlagsMAP(vData As colFlagsMAP)
    Set mvarcolFlagsMAP = vData
End Property
Private Sub Class_Terminate()
    Set mvarcolFlagsMAP = Nothing
End Sub



Public Property Let oConn(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.oConn = 5
    mvaroConn = vData
End Property


Public Property Set oConn(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.oConn = Form1
    Set mvaroConn = vData
End Property


Public Property Get oConn() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.oConn
    If IsObject(mvaroConn) Then
        Set oConn = mvaroConn
    Else
        oConn = mvaroConn
    End If
End Property



Public Property Let CommandTimeout(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CommandTimeout = 5
    mvarCommandTimeout = vData
End Property


Public Property Get CommandTimeout() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CommandTimeout
    CommandTimeout = mvarCommandTimeout
End Property



Public Property Let ConnectionString(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ConnectionString = 5
    mvarConnectionString = vData
End Property


Public Property Get ConnectionString() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ConnectionString
    ConnectionString = mvarConnectionString
End Property



