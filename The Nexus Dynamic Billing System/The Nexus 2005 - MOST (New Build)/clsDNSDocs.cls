VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDNSDocs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvarRecID As Long 'local copy
Private mvarDomainID As Long 'local copy
Private mvarDocType As String 'local copy
Private mvarDocText As String 'local copy
Private mvarbIcon As Byte 'local copy
Private mvarDescription As String 'local copy
Private mvarItemText As String 'local copy
Public Property Let ItemText(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ItemText = 5
    mvarItemText = vData
End Property


Public Property Get ItemText() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ItemText
    ItemText = mvarItemText
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



Public Property Let bIcon(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bIcon = 5
    mvarbIcon = vData
End Property


Public Property Get bIcon() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bIcon
    bIcon = mvarbIcon
End Property



Public Property Let DocText(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DocText = 5
    mvarDocText = vData
End Property


Public Property Get DocText() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DocText
    DocText = mvarDocText
End Property



Public Property Let DocType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DocType = 5
    mvarDocType = vData
End Property


Public Property Get DocType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DocType
    DocType = mvarDocType
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



