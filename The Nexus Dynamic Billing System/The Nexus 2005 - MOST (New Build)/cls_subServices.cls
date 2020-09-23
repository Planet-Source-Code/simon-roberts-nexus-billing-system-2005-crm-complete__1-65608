VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_subServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

Private mvarRecID As Long 'local copy
Private mvarptRecID As Long 'local copy
Private mvarServiceID As Long 'local copy
Private mvarContactName As String 'local copy
Private mvarUsername As String 'local copy
Private mvarPassword As String 'local copy
Private mvarNextCycle As Date 'local copy
Private mvarBaseURL As String 'local copy
Private mvarRadiusID As Long 'local copy
Private mvarDateCreated As Date 'local copy
Private mvaracci_RecID As Long 'local copy
Private mvarDynamicField1 As String 'local copy
Private mvarDynamicField2 As String 'local copy
Private mvarDynamicField3 As String 'local copy
Private mvarDynamicField4 As String 'local copy
Private mvarDynamicField5 As String 'local copy
Private mvarChecked As Integer 'local copy
Private mvarVirtualID As Long 'local copy
Private mvarPreviousCycle As Date 'local copy
Private mvarDomainID As Long 'local copy
Private mvarSubRecID As Long 'local copy
Private mvarMBQuota As Integer 'local copy
Private mvarActivation As Date 'local copy
Private mvarPeriodFee As Single 'local copy
Private mvarPerHour As Single 'local copy
Private mvarPerMB As Single 'local copy
Private mvarJoiningFee As Single 'local copy
Private mvarAgencyID As Long 'local copy
Private mvarDefaultShippingID As Long 'local copy
Private mvarContractID As Long 'local copy
Private mvarContractExpiry As Date 'local copy
'local variable(s) to hold property value(s)
Private mvarIDX_Radius As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarActivationSet As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvarDescription As String 'local copy
Private mvarVendorCode As String 'local copy
Private mvarProductCode As String 'local copy
Private mvarProductSubCode As String 'local copy
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
Public Property Let ProductSubCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProductSubCode = 5
    mvarProductSubCode = vData
End Property


Public Property Get ProductSubCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProductSubCode
    ProductSubCode = mvarProductSubCode
End Property



Public Property Let ProductCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProductCode = 5
    mvarProductCode = vData
End Property


Public Property Get ProductCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProductCode
    ProductCode = mvarProductCode
End Property



Public Property Let VendorCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VendorCode = 5
    mvarVendorCode = vData
End Property


Public Property Get VendorCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VendorCode
    VendorCode = mvarVendorCode
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



Public Property Let ActivationSet(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ActivationSet = 5
    mvarActivationSet = vData
End Property


Public Property Get ActivationSet() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ActivationSet
    ActivationSet = mvarActivationSet
End Property



Public Property Let IDX_Radius(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IDX_Radius = 5
    mvarIDX_Radius = vData
End Property


Public Property Get IDX_Radius() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IDX_Radius
    IDX_Radius = mvarIDX_Radius
End Property



Public Property Let ContractExpiry(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ContractExpiry = 5
    mvarContractExpiry = vData
End Property


Public Property Get ContractExpiry() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ContractExpiry
    ContractExpiry = mvarContractExpiry
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



Public Property Let DefaultShippingID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DefaultShippingID = 5
    mvarDefaultShippingID = vData
End Property


Public Property Get DefaultShippingID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DefaultShippingID
    DefaultShippingID = mvarDefaultShippingID
End Property



Public Property Let AgencyID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AgencyID = 5
    mvarAgencyID = vData
End Property


Public Property Get AgencyID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AgencyID
    AgencyID = mvarAgencyID
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



Public Property Let PerMB(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PerMB = 5
    mvarPerMB = vData
End Property


Public Property Get PerMB() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PerMB
    PerMB = mvarPerMB
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



Public Property Let Activation(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Activation = 5
    mvarActivation = vData
End Property


Public Property Get Activation() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Activation
    Activation = mvarActivation
End Property



Public Property Let MBQuota(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MBQuota = 5
    mvarMBQuota = vData
End Property


Public Property Get MBQuota() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MBQuota
    MBQuota = mvarMBQuota
End Property



Public Property Let SubRecID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SubRecID = 5
    mvarSubRecID = vData
End Property


Public Property Get SubRecID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SubRecID
    SubRecID = mvarSubRecID
End Property



Public Property Let DomainID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DomainID = 5
    mvarDomainID = vData
End Property


Public Property Get DomainID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DomainID
    DomainID = mvarDomainID
End Property



Public Property Let PreviousCycle(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PreviousCycle = 5
    mvarPreviousCycle = vData
End Property


Public Property Get PreviousCycle() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PreviousCycle
    PreviousCycle = mvarPreviousCycle
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



Public Property Let Checked(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Checked = 5
    mvarChecked = vData
End Property


Public Property Get Checked() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Checked
    Checked = mvarChecked
End Property



Public Property Let DynamicField5(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DynamicField5 = 5
    mvarDynamicField5 = vData
End Property


Public Property Get DynamicField5() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DynamicField5
    DynamicField5 = mvarDynamicField5
End Property



Public Property Let DynamicField4(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DynamicField4 = 5
    mvarDynamicField4 = vData
End Property


Public Property Get DynamicField4() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DynamicField4
    DynamicField4 = mvarDynamicField4
End Property



Public Property Let DynamicField3(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DynamicField3 = 5
    mvarDynamicField3 = vData
End Property


Public Property Get DynamicField3() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DynamicField3
    DynamicField3 = mvarDynamicField3
End Property



Public Property Let DynamicField2(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DynamicField2 = 5
    mvarDynamicField2 = vData
End Property


Public Property Get DynamicField2() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DynamicField2
    DynamicField2 = mvarDynamicField2
End Property



Public Property Let DynamicField1(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DynamicField1 = 5
    mvarDynamicField1 = vData
End Property


Public Property Get DynamicField1() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DynamicField1
    DynamicField1 = mvarDynamicField1
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



Public Property Let RadiusID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RadiusID = 5
    mvarRadiusID = vData
End Property


Public Property Get RadiusID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RadiusID
    RadiusID = mvarRadiusID
End Property



Public Property Let BaseURL(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BaseURL = 5
    mvarBaseURL = vData
End Property


Public Property Get BaseURL() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BaseURL
    BaseURL = mvarBaseURL
End Property



Public Property Let NextCycle(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NextCycle = 5
    mvarNextCycle = vData
End Property


Public Property Get NextCycle() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NextCycle
    NextCycle = mvarNextCycle
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



Public Property Let ContactName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ContactName = 5
    mvarContactName = vData
End Property


Public Property Get ContactName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ContactName
    ContactName = mvarContactName
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



