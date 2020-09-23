VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVISPs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvarsDesc As String 'local copy
Private mvarRecIDa As Long 'local copy
Private mvarRecIDb As Long 'local copy
Public Property Let RecIDb(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RecIDb = 5
    mvarRecIDb = vData
End Property


Public Property Get RecIDb() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RecIDb
    RecIDb = mvarRecIDb
End Property



Public Property Let RecIDa(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RecIDa = 5
    mvarRecIDa = vData
End Property


Public Property Get RecIDa() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RecIDa
    RecIDa = mvarRecIDa
End Property



Public Property Let sDesc(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sDesc = 5
    mvarsDesc = vData
End Property


Public Property Get sDesc() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sDesc
    sDesc = mvarsDesc
End Property



