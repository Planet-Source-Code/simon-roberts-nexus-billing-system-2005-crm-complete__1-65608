VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPackageNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarDescription As String 'local copy
Private mvarNumOf As Integer 'local copy
Private mvarPeriodFee As Single 'local copy
Private mvarJoiningFee As Single 'local copy
Private mvarPerMBBlock As Single 'local copy
Private mvarPerHour As Single 'local copy
Private mvarcycType As String 'local copy
Private mvarcycInterval As Integer 'local copy
Private mvarsvrCode As String 'local copy
Private mvarBillNow As Boolean 'local copy
Private mvarMBQuota As Long 'local copy
Private mvarcJoiningFee As Single 'local copy
Private mvarcPeriodFee As Single 'local copy
Private mvarcPerHour As Single 'local copy
Private mvarcPerMBBlock As Single 'local copy
Private mvarSessionTimeout As Integer 'local copy
Private mvarIdleTimeout As Long 'local copy
Private mvarSessionsAllowed As Integer 'local copy
Private mvarServiceID As Long 'local copy
Private mvarListedOnRadius As Boolean 'local copy
Private mvarptRecID As Double 'local copy
Private mvarKey As String 'local copy
'local variable(s) to hold property value(s)
Private mvarbContract As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarDomainIDX As Variant 'local copy
Private mvarRadiusIDX As Long 'local copy
Public Property Let RadiusIDX(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RadiusIDX = 5
    mvarRadiusIDX = vData
End Property


Public Property Get RadiusIDX() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RadiusIDX
    RadiusIDX = mvarRadiusIDX
End Property



Public Property Let DomainIDX(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DomainIDX = 5
    mvarDomainIDX = vData
End Property


Public Property Set DomainIDX(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.DomainIDX = Form1
    Set mvarDomainIDX = vData
End Property


Public Property Get DomainIDX() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DomainIDX
    If IsObject(mvarDomainIDX) Then
        Set DomainIDX = mvarDomainIDX
    Else
        DomainIDX = mvarDomainIDX
    End If
End Property



Public Property Let bContract(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bContract = 5
    mvarbContract = vData
End Property


Public Property Get bContract() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bContract
    bContract = mvarbContract
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



Public Property Let ptRecID(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ptRecID = 5
    mvarptRecID = vData
End Property


Public Property Get ptRecID() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ptRecID
    ptRecID = mvarptRecID
End Property



Public Property Let ListedOnRadius(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ListedOnRadius = 5
    mvarListedOnRadius = vData
End Property


Public Property Get ListedOnRadius() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ListedOnRadius
    ListedOnRadius = mvarListedOnRadius
End Property



Public Property Let ServiceID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ServiceID = 5
    mvarServiceID = vData
End Property


Public Property Get ServiceID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ServiceID
    ServiceID = mvarServiceID
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



Public Property Let IdleTimeout(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IdleTimeout = 5
    mvarIdleTimeout = vData
End Property


Public Property Get IdleTimeout() As Long
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



Public Property Let cPerMBBlock(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.cPerMBBlock = 5
    mvarcPerMBBlock = vData
End Property


Public Property Get cPerMBBlock() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.cPerMBBlock
    cPerMBBlock = mvarcPerMBBlock
End Property



Public Property Let cPerHour(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.cPerHour = 5
    mvarcPerHour = vData
End Property


Public Property Get cPerHour() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.cPerHour
    cPerHour = mvarcPerHour
End Property



Public Property Let cPeriodFee(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.cPeriodFee = 5
    mvarcPeriodFee = vData
End Property


Public Property Get cPeriodFee() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.cPeriodFee
    cPeriodFee = mvarcPeriodFee
End Property



Public Property Let cJoiningFee(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.cJoiningFee = 5
    mvarcJoiningFee = vData
End Property


Public Property Get cJoiningFee() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.cJoiningFee
    cJoiningFee = mvarcJoiningFee
End Property



Public Property Let MBQuota(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MBQuota = 5
    mvarMBQuota = vData
End Property


Public Property Get MBQuota() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MBQuota
    MBQuota = mvarMBQuota
End Property



Public Property Let BillNow(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BillNow = 5
    mvarBillNow = vData
End Property


Public Property Get BillNow() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BillNow
    BillNow = mvarBillNow
End Property



Public Property Let svrCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.svrCode = 5
    mvarsvrCode = vData
End Property


Public Property Get svrCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.svrCode
    svrCode = mvarsvrCode
End Property



Public Property Let cycInterval(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.cycInterval = 5
    mvarcycInterval = vData
End Property


Public Property Get cycInterval() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.cycInterval
    cycInterval = mvarcycInterval
End Property



Public Property Let cycType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.cycType = 5
    mvarcycType = vData
End Property


Public Property Get cycType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.cycType
    cycType = mvarcycType
End Property



Public Property Let PerHour(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PerHour = 5
    mvarPerHour = vData
End Property


Public Property Get PerHour() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PerHour
    PerHour = mvarPerHour
End Property



Public Property Let PerMBBlock(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PerMBBlock = 5
    mvarPerMBBlock = vData
End Property


Public Property Get PerMBBlock() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PerMBBlock
    PerMBBlock = mvarPerMBBlock
End Property



Public Property Let JoiningFee(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.JoiningFee = 5
    mvarJoiningFee = vData
End Property


Public Property Get JoiningFee() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.JoiningFee
    JoiningFee = mvarJoiningFee
End Property



Public Property Let PeriodFee(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PeriodFee = 5
    mvarPeriodFee = vData
End Property


Public Property Get PeriodFee() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PeriodFee
    PeriodFee = mvarPeriodFee
End Property



Public Property Let NumOf(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NumOf = 5
    mvarNumOf = vData
End Property


Public Property Get NumOf() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NumOf
    NumOf = mvarNumOf
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



