VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPaymentSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvarddAccountName As String 'local copy
Private mvarddBSB As String 'local copy
Private mvarddAcountNo As String 'local copy
Private mvarccCardName As String 'local copy
Private mvarccCardNumber As String 'local copy
Private mvarccCardExpiry As String 'local copy
Private mvarccCIC As String 'local copy
Private mvarswWord As String 'local copy
Private mvarswNumber As String 'local copy
Private mvarIDX As Long 'local copy
Private mvarOrder As Long 'local copy
Private mvarRecID As Long 'local copy
Private mvaracci_RecID As Long 'local copy
Private mvarChecked As Integer 'local copy
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
'Syntax: X.RecID = 5
    mvaracci_RecID = vData
End Property


Public Property Get acci_RecID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RecID
    acci_RecID = mvaracci_RecID
End Property

Public Property Let Checked(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RecID = 5
    mvarChecked = vData
End Property


Public Property Get Checked() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RecID
    Checked = mvarChecked
End Property



Public Property Let Order(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Order = 5
    mvarOrder = vData
End Property


Public Property Get Order() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Order
    Order = mvarOrder
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



Public Property Let swNumber(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.swNumber = 5
    mvarswNumber = vData
End Property


Public Property Get swNumber() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.swNumber
    swNumber = mvarswNumber
End Property



Public Property Let swWord(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.swWord = 5
    mvarswWord = vData
End Property


Public Property Get swWord() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.swWord
    swWord = mvarswWord
End Property



Public Property Let ccCIC(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ccCIC = 5
    mvarccCIC = vData
End Property


Public Property Get ccCIC() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ccCIC
    ccCIC = mvarccCIC
End Property



Public Property Let ccCardExpiry(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ccCardExpiry = 5
    mvarccCardExpiry = vData
End Property


Public Property Get ccCardExpiry() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ccCardExpiry
    ccCardExpiry = mvarccCardExpiry
End Property



Public Property Let ccCardNumber(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ccCardNumber = 5
    mvarccCardNumber = vData
End Property


Public Property Get ccCardNumber() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ccCardNumber
    ccCardNumber = mvarccCardNumber
End Property



Public Property Let ccCardName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ccCardName = 5
    mvarccCardName = vData
End Property


Public Property Get ccCardName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ccCardName
    ccCardName = mvarccCardName
End Property



Public Property Let ddAcountNo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ddAcountNo = 5
    mvarddAcountNo = vData
End Property


Public Property Get ddAcountNo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ddAcountNo
    ddAcountNo = mvarddAcountNo
End Property



Public Property Let ddBSB(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ddBSB = 5
    mvarddBSB = vData
End Property


Public Property Get ddBSB() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ddBSB
    ddBSB = mvarddBSB
End Property



Public Property Let ddAccountName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ddAccountName = 5
    mvarddAccountName = vData
End Property


Public Property Get ddAccountName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ddAccountName
    ddAccountName = mvarddAccountName
End Property



