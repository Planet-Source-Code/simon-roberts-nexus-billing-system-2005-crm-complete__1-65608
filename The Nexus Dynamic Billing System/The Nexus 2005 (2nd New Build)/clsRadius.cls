VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRadius"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvarAcct_Session_ID As String 'local copy
Private mvarUniqueSessionID As String 'local copy
Private mvarUsername As String 'local copy
Private mvarStartTime As Date 'local copy
Private mvarStopTime As Date 'local copy
Private mvarAliveTime As Date 'local copy
Private mvarbytesInput As Single 'local copy
Private mvarbytesOuput As Single 'local copy
Public Property Let bytesOuput(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bytesOuput = 5
    mvarbytesOuput = vData
End Property


Public Property Get bytesOuput() As Single
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bytesOuput
    bytesOuput = mvarbytesOuput
End Property



Public Property Let bytesInput(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bytesInput = 5
    mvarbytesInput = vData
End Property


Public Property Get bytesInput() As Single
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bytesInput
    bytesInput = mvarbytesInput
End Property



Public Property Let AliveTime(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AliveTime = 5
    mvarAliveTime = vData
End Property


Public Property Get AliveTime() As Date
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AliveTime
    AliveTime = mvarAliveTime
End Property



Public Property Let StopTime(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StopTime = 5
    mvarStopTime = vData
End Property


Public Property Get StopTime() As Date
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StopTime
    StopTime = mvarStopTime
End Property



Public Property Let StartTime(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StartTime = 5
    mvarStartTime = vData
End Property


Public Property Get StartTime() As Date
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StartTime
    StartTime = mvarStartTime
End Property



Public Property Let Username(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Username = 5
    mvarUsername = vData
End Property


Public Property Get Username() As String
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Username
    Username = mvarUsername
End Property



Public Property Let UniqueSessionID(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UniqueSessionID = 5
    mvarUniqueSessionID = vData
End Property


Public Property Get UniqueSessionID() As String
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UniqueSessionID
    UniqueSessionID = mvarUniqueSessionID
End Property



Public Property Let Acct_Session_ID(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Acct_Session_Id = 5
    mvarAcct_Session_ID = vData
End Property


Public Property Get Acct_Session_ID() As String
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Acct_Session_Id
    Acct_Session_ID = mvarAcct_Session_ID
End Property



