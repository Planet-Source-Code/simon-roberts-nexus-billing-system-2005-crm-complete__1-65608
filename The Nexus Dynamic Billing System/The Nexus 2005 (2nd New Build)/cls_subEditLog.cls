VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_subEditLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvaracci_RecID As Long 'local copy
Private mvarDateEditMade As Date 'local copy
Private mvarSysopID As Long 'local copy
Private mvarEditTxt As String 'local copy
Private mvarIPAddress As String 'local copy
Private mvarRecID As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarUsername As String 'local copy
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



Public Property Let IPAddress(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IPAddress = 5
    mvarIPAddress = vData
End Property


Public Property Get IPAddress() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IPAddress
    IPAddress = mvarIPAddress
End Property



Public Property Let EditTxt(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EditTxt = 5
    mvarEditTxt = vData
End Property


Public Property Get EditTxt() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EditTxt
    EditTxt = mvarEditTxt
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



Public Property Let DateEditMade(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateEditMade = 5
    mvarDateEditMade = vData
End Property


Public Property Get DateEditMade() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateEditMade
    DateEditMade = mvarDateEditMade
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



