VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReseller_emailcatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvarRecID As Long 'local copy
Private mvarVirtualID As Long 'local copy
Private mvarCatchType As String 'local copy
Private mvarEmailAddy As String 'local copy
Public Property Let EmailAddy(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EmailAddy = 5
    mvarEmailAddy = vData
End Property


Public Property Get EmailAddy() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EmailAddy
    EmailAddy = mvarEmailAddy
End Property



Public Property Let CatchType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CatchType = 5
    mvarCatchType = vData
End Property


Public Property Get CatchType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CatchType
    CatchType = mvarCatchType
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



