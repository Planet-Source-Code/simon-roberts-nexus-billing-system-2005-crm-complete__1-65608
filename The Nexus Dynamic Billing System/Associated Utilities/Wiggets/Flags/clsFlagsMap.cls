VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFlagsMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

Private mvarTopMost As Boolean 'local copy
Private mvarIDX As Long 'local copy
Private mvarSubIDX As Long 'local copy
Private mvarRecID As Long 'local copy
Private mvarSubRecID As Long 'local copy
Private mvarIconIDX As Long 'local copy
Private mvarIconSTR As String 'local copy
Private mvarDescription As String 'local copy
Private mvarTitle As String 'local copy
Private mvarURL As String 'local copy
Private mvarFormName As String 'local copy
Public Property Let FormName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FormName = 5
    mvarFormName = vData
End Property


Public Property Get FormName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FormName
    FormName = mvarFormName
End Property



Public Property Let URL(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.URL = 5
    mvarURL = vData
End Property


Public Property Get URL() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.URL
    URL = mvarURL
End Property



Public Property Let Title(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Title = 5
    mvarTitle = vData
End Property


Public Property Get Title() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Title
    Title = mvarTitle
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



Public Property Let IconSTR(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IconSTR = 5
    mvarIconSTR = vData
End Property


Public Property Get IconSTR() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IconSTR
    IconSTR = mvarIconSTR
End Property



Public Property Let IconIDX(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IconIDX = 5
    mvarIconIDX = vData
End Property


Public Property Get IconIDX() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IconIDX
    IconIDX = mvarIconIDX
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



Public Property Let SubIDX(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SubIDX = 5
    mvarSubIDX = vData
End Property


Public Property Get SubIDX() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SubIDX
    SubIDX = mvarSubIDX
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



Public Property Let TopMost(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TopMost = 5
    mvarTopMost = vData
End Property


Public Property Get TopMost() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TopMost
    TopMost = mvarTopMost
End Property



