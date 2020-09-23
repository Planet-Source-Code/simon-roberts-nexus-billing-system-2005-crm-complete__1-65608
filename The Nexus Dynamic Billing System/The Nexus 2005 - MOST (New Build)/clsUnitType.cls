VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUnitType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvarsUnitName As String 'local copy
Private mvarsFieldName As String 'local copy
Public Property Let sFieldName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sFieldName = 5
    mvarsFieldName = vData
End Property


Public Property Get sFieldName() As String
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sFieldName
    sFieldName = mvarsFieldName
End Property



Public Property Let sUnitName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sUnitName = 5
    mvarsUnitName = vData
End Property


Public Property Get sUnitName() As String
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sUnitName
    sUnitName = mvarsUnitName
End Property



