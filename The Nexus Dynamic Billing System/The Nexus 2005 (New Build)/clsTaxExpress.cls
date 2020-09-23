VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTaxExpress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

Private mvarCode As String 'local copy
Private mvarCountry As String 'local copy
Private mvarRate As Single 'local copy
Private mvariFlag As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvarDesc As String 'local copy
'local variable(s) to hold property value(s)
Private mvarRangeMin As Currency 'local copy
Private mvarRangeMax As Currency 'local copy
Private mvarcRate As Currency 'local copy
Private mvarlGroup As Long 'local copy
Public Property Let lGroup(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lGroup = 5
    mvarlGroup = vData
End Property


Public Property Get lGroup() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lGroup
    lGroup = mvarlGroup
End Property



Public Property Let cRate(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.cRate = 5
    mvarcRate = vData
End Property


Public Property Get cRate() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.cRate
    cRate = mvarcRate
End Property



Public Property Let RangeMax(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RangeMax = 5
    mvarRangeMax = vData
End Property


Public Property Get RangeMax() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RangeMax
    RangeMax = mvarRangeMax
End Property



Public Property Let RangeMin(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RangeMin = 5
    mvarRangeMin = vData
End Property


Public Property Get RangeMin() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RangeMin
    RangeMin = mvarRangeMin
End Property



Public Property Let Desc(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Desc = 5
    mvarDesc = vData
End Property


Public Property Get Desc() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Desc
    Desc = mvarDesc
End Property



Public Property Let iFlag(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFlag = 5
    mvariFlag = vData
End Property


Public Property Get iFlag() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFlag
    iFlag = mvariFlag
End Property



Public Property Let Rate(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Rate = 5
    mvarRate = vData
End Property


Public Property Get Rate() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Rate
    Rate = mvarRate
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



Public Property Let Code(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Code = 5
    mvarCode = vData
End Property


Public Property Get Code() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Code
    Code = mvarCode
End Property



