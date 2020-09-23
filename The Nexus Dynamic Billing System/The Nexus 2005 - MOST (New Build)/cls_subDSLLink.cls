VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_subDSLLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvarAccountName As String 'local copy
Private mvarAreaCode As String 'local copy
Private mvarPhoneNumber As String 'local copy
Private mvareMail As String 'local copy
Private mvaracci_RecID As Long 'local copy
Private mvarRecID As Long 'local copy
Private mvarChecked As Integer 'local copy
Private mvarUnitNo As String 'local copy
Private mvarStreetNo As String 'local copy
Private mvarStreetName As String 'local copy
Private mvarStreetType As String 'local copy
Private mvarSuburb As String 'local copy
Private mvarCountry As String 'local copy
Private mvarPostCode As String 'local copy
Private mvarState As String 'local copy
Private mvarCreated As Date 'local copy
Private mvarChurn As Byte 'local copy
Private mvarRadiusID As Long 'local copy
Private mvarAccountActive As Integer 'local copy
Private mvarcFlag As Byte 'local copy
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
Public Property Let cFlag(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.cFlag = 5
    mvarcFlag = vData
End Property


Public Property Get cFlag() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.cFlag
    cFlag = mvarcFlag
End Property



Public Property Let AccountActive(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AccountActive = 5
    mvarAccountActive = vData
End Property


Public Property Get AccountActive() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AccountActive
    AccountActive = mvarAccountActive
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



Public Property Let Churn(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Churn = 5
    mvarChurn = vData
End Property


Public Property Get Churn() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Churn
    Churn = mvarChurn
End Property



Public Property Let Created(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Created = 5
    mvarCreated = vData
End Property


Public Property Get Created() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Created
    Created = mvarCreated
End Property



Public Property Let State(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.State = 5
    mvarState = vData
End Property


Public Property Get State() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.State
    State = mvarState
End Property



Public Property Let PostCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Postcode = 5
    mvarPostCode = vData
End Property


Public Property Get PostCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Postcode
    PostCode = mvarPostCode
End Property



Public Property Let Country(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Country = 5
    mvarCountry = vData
End Property


Public Property Get Country() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Country
    Country = mvarCountry
End Property



Public Property Let Suburb(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Suburb = 5
    mvarSuburb = vData
End Property


Public Property Get Suburb() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Suburb
    Suburb = mvarSuburb
End Property



Public Property Let StreetType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StreetType = 5
    mvarStreetType = vData
End Property


Public Property Get StreetType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StreetType
    StreetType = mvarStreetType
End Property



Public Property Let StreetName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StreetName = 5
    mvarStreetName = vData
End Property


Public Property Get StreetName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StreetName
    StreetName = mvarStreetName
End Property



Public Property Let StreetNo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StreetNo = 5
    mvarStreetNo = vData
End Property


Public Property Get StreetNo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StreetNo
    StreetNo = mvarStreetNo
End Property



Public Property Let UnitNo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UnitNo = 5
    mvarUnitNo = vData
End Property


Public Property Get UnitNo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UnitNo
    UnitNo = mvarUnitNo
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



Public Property Let eMail(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.eMail = 5
    mvareMail = vData
End Property


Public Property Get eMail() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.eMail
    eMail = mvareMail
End Property



Public Property Let PhoneNumber(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PhoneNumber = 5
    mvarPhoneNumber = vData
End Property


Public Property Get PhoneNumber() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PhoneNumber
    PhoneNumber = mvarPhoneNumber
End Property



Public Property Let AreaCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AreaCode = 5
    mvarAreaCode = vData
End Property


Public Property Get AreaCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AreaCode
    AreaCode = mvarAreaCode
End Property



Public Property Let AccountName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AccountName = 5
    mvarAccountName = vData
End Property


Public Property Get AccountName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AccountName
    AccountName = mvarAccountName
End Property



