VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSysops"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvarRecID As Double 'local copy
Private mvarUsername As String 'local copy
Private mvarPublicKey As String 'local copy
Private mvarDescription As String 'local copy
Private mvarSecurityLevel As Byte 'local copy
Private mvarChecked As Boolean 'local copy
Private mvarVirtualID As Double 'local copy
Private mvarAgencyID As Double 'local copy
Private mvarbMaintain As Boolean 'local copy
Private mvarbVISP As Boolean 'local copy
Private mvarbCreateSysop As Boolean 'local copy
Private mvarbPrimary As Boolean 'local copy
Private mvarbAgency As Boolean 'local copy
Private mvarbTemplates As Boolean 'local copy
Private mvarFirstname As String 'local copy
Private mvarSurname As String 'local copy
Private mvarNextCycle As Date 'local copy
Private mvarPreviousCycle As Date 'local copy
Private mvarbRecievables As Boolean 'local copy
Private mvarbInvoice As Boolean 'local copy
Private mvarbExpenditure As Boolean 'local copy
Private mvarbHoldings As Boolean 'local copy
Private mvarbComm As Boolean 'local copy
Private mvarbRefund As Boolean 'local copy
Private mvarbAddCust As Boolean 'local copy
Private mvarbOwnership As Boolean 'local copy
Private mvarbAccSettings As Boolean 'local copy
Private mvarbVendors As Boolean 'local copy
Public Property Let bVendors(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bVendors = 5
    mvarbVendors = vData
End Property


Public Property Get bVendors() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bVendors
    bVendors = mvarbVendors
End Property



Public Property Let bAccSettings(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bAccSettings = 5
    mvarbAccSettings = vData
End Property


Public Property Get bAccSettings() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bAccSettings
    bAccSettings = mvarbAccSettings
End Property



Public Property Let bOwnership(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bOwnership = 5
    mvarbOwnership = vData
End Property


Public Property Get bOwnership() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bOwnership
    bOwnership = mvarbOwnership
End Property



Public Property Let bAddCust(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bAddCust = 5
    mvarbAddCust = vData
End Property


Public Property Get bAddCust() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bAddCust
    bAddCust = mvarbAddCust
End Property



Public Property Let bRefund(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bRefund = 5
    mvarbRefund = vData
End Property


Public Property Get bRefund() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bRefund
    bRefund = mvarbRefund
End Property



Public Property Let bComm(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bComm = 5
    mvarbComm = vData
End Property


Public Property Get bComm() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bComm
    bComm = mvarbComm
End Property



Public Property Let bHoldings(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bHoldings = 5
    mvarbHoldings = vData
End Property


Public Property Get bHoldings() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bHoldings
    bHoldings = mvarbHoldings
End Property



Public Property Let bExpenditure(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bExpenditure = 5
    mvarbExpenditure = vData
End Property


Public Property Get bExpenditure() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bExpenditure
    bExpenditure = mvarbExpenditure
End Property



Public Property Let bInvoice(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bInvoice = 5
    mvarbInvoice = vData
End Property


Public Property Get bInvoice() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bInvoice
    bInvoice = mvarbInvoice
End Property



Public Property Let bRecievables(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bRecievables = 5
    mvarbRecievables = vData
End Property


Public Property Get bRecievables() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bRecievables
    bRecievables = mvarbRecievables
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



Public Property Let Surname(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Surname = 5
    mvarSurname = vData
End Property


Public Property Get Surname() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Surname
    Surname = mvarSurname
End Property



Public Property Let Firstname(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Firstname = 5
    mvarFirstname = vData
End Property


Public Property Get Firstname() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Firstname
    Firstname = mvarFirstname
End Property



Public Property Let bTemplates(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bTemplates = 5
    mvarbTemplates = vData
End Property


Public Property Get bTemplates() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bTemplates
    bTemplates = mvarbTemplates
End Property



Public Property Let bAgency(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bAgency = 5
    mvarbAgency = vData
End Property


Public Property Get bAgency() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bAgency
    bAgency = mvarbAgency
End Property



Public Property Let bPrimary(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bPrimary = 5
    mvarbPrimary = vData
End Property


Public Property Get bPrimary() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bPrimary
    bPrimary = mvarbPrimary
End Property



Public Property Let bCreateSysop(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bCreateSysop = 5
    mvarbCreateSysop = vData
End Property


Public Property Get bCreateSysop() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bCreateSysop
    bCreateSysop = mvarbCreateSysop
End Property



Public Property Let bVISP(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bVISP = 5
    mvarbVISP = vData
End Property


Public Property Get bVISP() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bVISP
    bVISP = mvarbVISP
End Property



Public Property Let bMaintain(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bMaintain = 5
    mvarbMaintain = vData
End Property


Public Property Get bMaintain() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bMaintain
    bMaintain = mvarbMaintain
End Property



Public Property Let AgencyID(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AgencyID = 5
    mvarAgencyID = vData
End Property


Public Property Get AgencyID() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AgencyID
    AgencyID = mvarAgencyID
End Property



Public Property Let VirtualID(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VirtualID = 5
    mvarVirtualID = vData
End Property


Public Property Get VirtualID() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VirtualID
    VirtualID = mvarVirtualID
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



Public Property Let SecurityLevel(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SecurityLevel = 5
    mvarSecurityLevel = vData
End Property


Public Property Get SecurityLevel() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SecurityLevel
    SecurityLevel = mvarSecurityLevel
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



Public Property Let PublicKey(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.odb.colSalts.ReturnSalt("PublicKey") = 5
    PublicKey = vData
End Property


Public Property Get PublicKey() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.odb.colSalts.ReturnSalt("PublicKey")
    PublicKey = PublicKey
    
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



