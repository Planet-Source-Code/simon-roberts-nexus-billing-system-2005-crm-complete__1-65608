VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsResellers_Addresses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

Private mvarRecID As Long 'local copy
Private mvarIDX As Long 'local copy
Private mvarSESSION As String 'local copy
Private mvarvisp_RecID As Long 'local copy
Private mvarFlagID As Integer 'local copy
Private mvarDateAdded As Date 'local copy
Private mvarContactName As String 'local copy
Private mvarStreet1 As String 'local copy
Private mvarStreet2 As String 'local copy
Private mvarCountry As String 'local copy
Private mvarState As String 'local copy
Private mvarSuburb As String 'local copy
Private mvarPostCode As String 'local copy
Private mvarCancelled As Integer 'local copy
Private mvarChecked As Integer 'local copy
Private mvarPhotoURL As String 'local copy
Private mvarFetchStatus As Byte 'local copy

Public Property Let FetchStatus(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FetchStatus = 5
    mvarFetchStatus = vData
End Property


Public Property Get FetchStatus() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FetchStatus
    FetchStatus = mvarFetchStatus
End Property





Public Property Let PhotoURL(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PhotoURL = 5
    mvarPhotoURL = vData
End Property


Public Property Get PhotoURL() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PhotoURL
    PhotoURL = mvarPhotoURL
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



Public Property Let PostCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PostCode = 5
    mvarPostCode = vData
End Property


Public Property Get PostCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PostCode
    PostCode = mvarPostCode
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



Public Property Let DateAdded(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateAdded = 5
    mvarDateAdded = vData
End Property


Public Property Get DateAdded() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateAdded
    DateAdded = mvarDateAdded
End Property



Public Property Let FlagID(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FlagID = 5
    mvarFlagID = vData
End Property


Public Property Get FlagID() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FlagID
    FlagID = mvarFlagID
End Property



Public Property Let visp_RecID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.visp_RecID = 5
    mvarvisp_RecID = vData
End Property


Public Property Get visp_RecID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.visp_RecID
    visp_RecID = mvarvisp_RecID
End Property



Public Property Let SESSION(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SESSION = 5
    mvarSESSION = vData
End Property


Public Property Get SESSION() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SESSION
    SESSION = mvarSESSION
End Property



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



