VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_subRadius"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
'local variable(s) to hold property value(s)
Private mvarKey As String 'local copy
Private mvarRecID As Long 'local copy
Private mvarUsername As String 'local copy
Private mvarPassword As String 'local copy
Private mvarSessionsAllowed As Integer 'local copy
Private mvarAutoActivateFlag As Boolean 'local copy
Private mvarActivate As Date 'local copy
Private mvarDeactivate As Date 'local copy
Private mvarSessionTimeout As Integer 'local copy
Private mvarIdleTimeout As Integer 'local copy
Private mvaracci_RecID As Long 'local copy
Private mvarServiceType As String 'local copy
Private mvarFramedProtocol As String 'local copy
Private mvarChecked As Boolean 'local copy
Private mvarsfStartTime As Date 'local copy
Private mvarsfStopTime As Date 'local copy
Private mvarsfAliveTime As Date 'local copy
Private mvarsfCycle_Upload As Double 'local copy
Private mvarsfCycle_Download As Double 'local copy
Private mvarsfCycle_Mins As Long 'local copy
Private mvarAcct_Session_ID As String 'local copy
Private mvarptRecID As Long 'local copy
Private mvarVirtualID As Long 'local copy
Private mvarPrimaryDNS As String 'local copy
Private mvarSecondaryDNS As String 'local copy
Private mvarDateCreated As Date 'local copy
'local variable(s) to hold property value(s)
Private mvarIDX As Long 'local copy
Public Property Let IDX(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IDX = 5
    mvarIDX = vData
End Property


Public Property Get IDX() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IDX
    IDX = mvarIDX
End Property



Public Property Let DateCreated(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateCreated = 5
    mvarDateCreated = vData
End Property


Public Property Get DateCreated() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateCreated
    DateCreated = mvarDateCreated
End Property



Public Property Let SecondaryDNS(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SecondaryDNS = 5
    mvarSecondaryDNS = vData
End Property


Public Property Get SecondaryDNS() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SecondaryDNS
    SecondaryDNS = mvarSecondaryDNS
End Property



Public Property Let PrimaryDNS(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrimaryDNS = 5
    mvarPrimaryDNS = vData
End Property


Public Property Get PrimaryDNS() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrimaryDNS
    PrimaryDNS = mvarPrimaryDNS
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



Public Property Let ptRecID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ptRecID = 5
    mvarptRecID = vData
End Property


Public Property Get ptRecID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ptRecID
    ptRecID = mvarptRecID
End Property



Public Property Let Acct_Session_ID(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Acct_Session_ID = 5
    mvarAcct_Session_ID = vData
End Property


Public Property Get Acct_Session_ID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Acct_Session_ID
    Acct_Session_ID = mvarAcct_Session_ID
End Property



Public Property Let sfCycle_Mins(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sfCycle_Mins = 5
    mvarsfCycle_Mins = vData
End Property


Public Property Get sfCycle_Mins() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sfCycle_Mins
    sfCycle_Mins = mvarsfCycle_Mins
End Property



Public Property Let sfCycle_Download(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sfCycle_Download = 5
    mvarsfCycle_Download = vData
End Property


Public Property Get sfCycle_Download() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sfCycle_Download
    sfCycle_Download = mvarsfCycle_Download
End Property



Public Property Let sfCycle_Upload(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sfCycle_Upload = 5
    mvarsfCycle_Upload = vData
End Property


Public Property Get sfCycle_Upload() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sfCycle_Upload
    sfCycle_Upload = mvarsfCycle_Upload
End Property



Public Property Let sfAliveTime(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sfAliveTime = 5
    mvarsfAliveTime = vData
End Property


Public Property Get sfAliveTime() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sfAliveTime
    sfAliveTime = mvarsfAliveTime
End Property



Public Property Let sfStopTime(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sfStopTime = 5
    mvarsfStopTime = vData
End Property


Public Property Get sfStopTime() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sfStopTime
    sfStopTime = mvarsfStopTime
End Property



Public Property Let sfStartTime(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sfStartTime = 5
    mvarsfStartTime = vData
End Property


Public Property Get sfStartTime() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sfStartTime
    sfStartTime = mvarsfStartTime
End Property



Public Property Let Checked(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Checked = 5
    mvarChecked = vData
End Property


Public Property Get Checked() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Checked
    Checked = mvarChecked
End Property



Public Property Let FramedProtocol(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FramedProtocol = 5
    mvarFramedProtocol = vData
End Property


Public Property Get FramedProtocol() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FramedProtocol
    FramedProtocol = mvarFramedProtocol
End Property



Public Property Let ServiceType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ServiceType = 5
    mvarServiceType = vData
End Property


Public Property Get ServiceType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ServiceType
    ServiceType = mvarServiceType
End Property



Public Property Let acci_RecID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.acci_RecID = 5
    mvaracci_RecID = vData
End Property


Public Property Get acci_RecID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.acci_RecID
    acci_RecID = mvaracci_RecID
End Property



Public Property Let IdleTimeout(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IdleTimeout = 5
    mvarIdleTimeout = vData
End Property


Public Property Get IdleTimeout() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IdleTimeout
    IdleTimeout = mvarIdleTimeout
End Property



Public Property Let SessionTimeout(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SessionTimeout = 5
    mvarSessionTimeout = vData
End Property


Public Property Get SessionTimeout() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SessionTimeout
    SessionTimeout = mvarSessionTimeout
End Property



Public Property Let Deactivate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Deactivate = 5
    mvarDeactivate = vData
End Property


Public Property Get Deactivate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Deactivate
    Deactivate = mvarDeactivate
End Property



Public Property Let Activate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Activate = 5
    mvarActivate = vData
End Property


Public Property Get Activate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Activate
    Activate = mvarActivate
End Property



Public Property Let AutoActivateFlag(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AutoActivateFlag = 5
    mvarAutoActivateFlag = vData
End Property


Public Property Get AutoActivateFlag() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AutoActivateFlag
    AutoActivateFlag = mvarAutoActivateFlag
End Property



Public Property Let SessionsAllowed(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SessionsAllowed = 5
    mvarSessionsAllowed = vData
End Property


Public Property Get SessionsAllowed() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SessionsAllowed
    SessionsAllowed = mvarSessionsAllowed
End Property



Public Property Let Password(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Password = 5
    mvarPassword = vData
End Property


Public Property Get Password() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Password
    Password = mvarPassword
End Property



Public Property Let Username(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Username = 5
    mvarUsername = vData
End Property


Public Property Get Username() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Username
    Username = mvarUsername
End Property



Public Property Let RecID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RecID = 5
    mvarRecID = vData
End Property


Public Property Get RecID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RecID
    RecID = mvarRecID
End Property



Public Property Let Key(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Key = 5
    mvarKey = vData
End Property


Public Property Get Key() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Key
    Key = mvarKey
End Property



