VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvarsCommand As String 'local copy
Private mvarExample As String 'local copy
'local variable(s) to hold property value(s)
Private mvarbFields As Byte 'local copy
Public Property Let bFields(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bFields = 5
    mvarbFields = vData
End Property


Public Property Get bFields() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bFields
    bFields = mvarbFields
End Property



Public Property Let Example(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Example = 5
    mvarExample = vData
End Property


Public Property Get Example() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Example
    Example = mvarExample
End Property



Public Property Let sCommand(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCommand = 5
    mvarsCommand = vData
End Property


Public Property Get sCommand() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCommand
    sCommand = mvarsCommand
End Property



