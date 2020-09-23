VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colPackages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsPackageNode"
Attribute VB_Ext_KEY = "Member0" ,"clsPackageNode"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable to hold collection
Private mCol As Collection
Private mvarttlExTax As Currency 'local copy
Private mvarttlTax As Currency 'local copy
Private mvarttlCost As Currency 'local copy
Private mvarttlMargin As Currency 'local copy
'local variable(s) to hold property value(s)
'local variable(s) to hold property value(s)
Private mvarContractID As Long 'local copy
Private mvarIntervalLength As Long 'local copy
Private mvarIntervalType As String 'local copy
'local variable(s) to hold property value(s)
Private mvarActivated As Date 'local copy
Public Property Let Activated(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Activated = 5
    mvarActivated = vData
End Property


Public Property Get Activated() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Activated
    Activated = mvarActivated
End Property



Public Property Let IntervalType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IntervalType = 5
    mvarIntervalType = vData
End Property


Public Property Get IntervalType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IntervalType
    IntervalType = mvarIntervalType
End Property



Public Property Let IntervalLength(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IntervalLength = 5
    mvarIntervalLength = vData
End Property


Public Property Get IntervalLength() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IntervalLength
    IntervalLength = mvarIntervalLength
End Property



Public Property Let ContractID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ContractID = 5
    mvarContractID = vData
End Property


Public Property Get ContractID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ContractID
    ContractID = mvarContractID
End Property





Public Sub ClearContract()

    If Me.Count > 0 Then
        Dim ix As Long
        For ix = Me.Count To 1 Step -1
            If Me.Item(ix).bContract = True Then Me.Remove ix
        Next
    End If
    
End Sub

Public Property Let ttlMargin(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ttlMargin = 5
    mvarttlMargin = vData
End Property


Public Property Get ttlMargin() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ttlMargin
    ttlMargin = mvarttlMargin
End Property



Public Property Let ttlCost(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ttlCost = 5
    mvarttlCost = vData
End Property


Public Property Get ttlCost() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ttlCost
    ttlCost = mvarttlCost
End Property



Public Property Let ttlTax(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ttlTax = 5
    mvarttlTax = vData
End Property


Public Property Get ttlTax() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ttlTax
    ttlTax = mvarttlTax
End Property



Public Property Let ttlExTax(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ttlExTax = 5
    mvarttlExTax = vData
End Property


Public Property Get ttlExTax() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ttlExTax
    ttlExTax = mvarttlExTax
End Property



Public Sub calTotals(Optional TaxRate As Single = 0.1)

    On Error Resume Next
    
    Dim cExTax As Currency
    Dim cTax As Currency
    Dim cCost As Currency
    Dim cMargin As Currency
    
    Dim ix As Long
    
    For ix = Me.Count To 1 Step -1
        cExTax = cExTax + Me(ix).PeriodFee '+ Me(ix).JoiningFee
        cCost = cCost + Me(ix).cPeriodFee '+ Me(ix).cJoiningFee
    Next
    
    cTax = cExTax + cExTax * TaxRate
    cCost = cCost + cCost * TaxRate
    cMargin = cTax - cCost
    
    Me.ttlCost = cCost
    Me.ttlExTax = cExTax
    Me.ttlMargin = cMargin
    Me.ttlTax = cTax
    
End Sub

Public Sub Clear()

    If Me.Count > 0 Then
        Dim ix As Long
        For ix = Me.Count To 1 Step -1
            Me.Remove ix
        Next
    End If
    
End Sub


Public Function Add(Description As String, NumOf As Integer, PeriodFee As Single, JoiningFee As Single, PerMBBlock As Single, PerHour As Single, cycType As String, cycInterval As Integer, svrCode As String, BillNow As Boolean, MBQuota As Long, cJoiningFee As Single, cPeriodFee As Single, cPerHour As Single, cPerMBBlock As Single, SessionTimeout As Long, IdleTimeout As Long, SessionsAllowed As Integer, ServiceID As Long, ListedOnRadius As Boolean, ptRecID As Long, Key As String, Optional sKey As String) As clsPackageNode
    'create a new object
    Dim objNewMember As clsPackageNode
    Set objNewMember = New clsPackageNode


    'set the properties passed into the method
    objNewMember.Description = Description
    objNewMember.NumOf = NumOf
    objNewMember.PeriodFee = PeriodFee
    objNewMember.JoiningFee = JoiningFee
    objNewMember.PerMBBlock = PerMBBlock
    objNewMember.PerHour = PerHour
    objNewMember.cycType = cycType
    objNewMember.cycInterval = cycInterval
    objNewMember.svrCode = svrCode
    objNewMember.BillNow = BillNow
    objNewMember.MBQuota = MBQuota
    objNewMember.cJoiningFee = cJoiningFee
    objNewMember.cPeriodFee = cPeriodFee
    objNewMember.cPerHour = cPerHour
    objNewMember.cPerMBBlock = cPerMBBlock
    objNewMember.SessionTimeout = SessionTimeout
    objNewMember.IdleTimeout = IdleTimeout
    objNewMember.SessionsAllowed = SessionsAllowed
    objNewMember.ServiceID = ServiceID
    objNewMember.ListedOnRadius = ListedOnRadius
    objNewMember.ptRecID = ptRecID
    objNewMember.Key = Key
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As clsPackageNode
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  If vntIndexKey = 0 Then Exit Property
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

