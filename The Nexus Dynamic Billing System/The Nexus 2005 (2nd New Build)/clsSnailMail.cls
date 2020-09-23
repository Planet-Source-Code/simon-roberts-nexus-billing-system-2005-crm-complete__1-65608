VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSnailMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

Private mvaracci_RecID As Long 'local copy
Private mvarFlagID As Byte 'local copy
Private mvarDateCreated As Date 'local copy
Private mvarContactName As String 'local copy
Private mvarStreet1 As String 'local copy
Private mvarStreet2 As String 'local copy
Private mvarCountry As String 'local copy
Private mvarState As String 'local copy
Private mvarPostCode As String 'local copy
Private mvarSuburb As String 'local copy
Private mvarCancelled As Integer 'local copy
Private mvarChecked As Integer 'local copy
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



Public Property Let Cancelled(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Cancelled = 5
    mvarCancelled = vData
End Property


Public Property Get Cancelled() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Cancelled
    Cancelled = mvarCancelled
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



Public Property Let Street2(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Street2 = 5
    mvarStreet2 = vData
End Property


Public Property Get Street2() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Street2
    Street2 = mvarStreet2
End Property



Public Property Let Street1(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Street1 = 5
    mvarStreet1 = vData
End Property


Public Property Get Street1() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Street1
    Street1 = mvarStreet1
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



Public Property Let FlagID(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FlagID = 5
    mvarFlagID = vData
End Property


Public Property Get FlagID() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FlagID
    FlagID = mvarFlagID
End Property



Public Property Let acci_RecID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AccI_RecID = 5
    mvaracci_RecID = vData
End Property


Public Property Get acci_RecID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AccI_RecID
    acci_RecID = mvaracci_RecID
End Property



