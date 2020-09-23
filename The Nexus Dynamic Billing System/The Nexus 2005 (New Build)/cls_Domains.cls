VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Domains"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvarRecID As Long 'local copy
Private mvarDomain As String 'local copy
Private mvarAdminEmail As String 'local copy
Private mvarStatus As String 'local copy
Private mvarChecked As Integer 'local copy
Private mvarSubDomainLimit As Integer 'local copy
Private mvarvKey As String 'local copy
Private mvaracci_RecID As Long 'local copy
Private mvarTechName As String 'local copy
Private mvarTechPass As String 'local copy
Private mvarSysopID As Long 'local copy
Private mvarVirtualID As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarContactName As String 'local copy
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



Public Property Let TechPass(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TechPass = 5
    mvarTechPass = vData
End Property


Public Property Get TechPass() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TechPass
    TechPass = mvarTechPass
End Property



Public Property Let TechName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TechName = 5
    mvarTechName = vData
End Property


Public Property Get TechName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TechName
    TechName = mvarTechName
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



Public Property Let vKey(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.vKey = 5
    mvarvKey = vData
End Property


Public Property Get vKey() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.vKey
    vKey = mvarvKey
End Property



Public Property Let SubDomainLimit(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SubDomainLimit = 5
    mvarSubDomainLimit = vData
End Property


Public Property Get SubDomainLimit() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SubDomainLimit
    SubDomainLimit = mvarSubDomainLimit
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



Public Property Let Status(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Status = 5
    mvarStatus = vData
End Property


Public Property Get Status() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Status
    Status = mvarStatus
End Property



Public Property Let AdminEmail(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AdminEmail = 5
    mvarAdminEmail = vData
End Property


Public Property Get AdminEmail() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AdminEmail
    AdminEmail = mvarAdminEmail
End Property



Public Property Let Domain(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Domain = 5
    mvarDomain = vData
End Property


Public Property Get Domain() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Domain
    Domain = mvarDomain
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



