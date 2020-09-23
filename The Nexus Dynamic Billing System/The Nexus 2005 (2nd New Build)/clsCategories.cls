VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvarRecID As Double 'local copy
Private mvarSubRecID As Double 'local copy
Private mvarVirtualID As Double 'local copy
Private mvarSysopID As Double 'local copy
Private mvarIcon As Integer 'local copy
Private mvarDescription As String 'local copy
Private mvarformcode As String 'local copy
Private mvarSecLevel As Byte 'local copy
'local variable(s) to hold property value(s)
Private mvarTag As String 'local copy
Public Property Let Tag(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Tag = 5
    mvarTag = vData
End Property


Public Property Get Tag() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Tag
    Tag = mvarTag
End Property



Public Property Let SecLevel(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SecLevel = 5
    mvarSecLevel = vData
End Property


Public Property Get SecLevel() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SecLevel
    SecLevel = mvarSecLevel
End Property



Public Property Let formcode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.formcode = 5
    mvarformcode = vData
End Property


Public Property Get formcode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.formcode
    formcode = mvarformcode
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



Public Property Let Icon(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Icon = 5
    mvarIcon = vData
End Property


Public Property Get Icon() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Icon
    Icon = mvarIcon
End Property



Public Property Let SysopID(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SysopID = 5
    mvarSysopID = vData
End Property


Public Property Get SysopID() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SysopID
    SysopID = mvarSysopID
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



Public Property Let SubRecID(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SubRecID = 5
    mvarSubRecID = vData
End Property


Public Property Get SubRecID() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SubRecID
    SubRecID = mvarSubRecID
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



