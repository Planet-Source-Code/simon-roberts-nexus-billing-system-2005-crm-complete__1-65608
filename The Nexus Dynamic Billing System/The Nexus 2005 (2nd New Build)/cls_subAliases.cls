VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_subAliases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvaracci_RecID As Long 'local copy
Private mvarftype As String 'local copy
Private mvareMail As String 'local copy
Private mvardest As String 'local copy
Private mvarChecked As Integer 'local copy
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
Public Property Let Checked(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Checked = 5
    mvarChecked = vData
End Property


Public Property Get Checked() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Checked
    Checked = mvarChecked
End Property



Public Property Let dest(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dest = 5
    mvardest = vData
End Property


Public Property Get dest() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dest
    dest = mvardest
End Property



Public Property Let eMail(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.email = 5
    mvareMail = vData
End Property


Public Property Get eMail() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.email
    eMail = mvareMail
End Property



Public Property Let ftype(ByVal vData As String)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ftype = Form1
    mvarftype = vData
End Property


Public Property Get ftype() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ftype
    ftype = mvarftype
End Property



Public Property Let acci_RecID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.acci_RecID = 5
    mvaracci_RecID = vData
End Property


Public Property Get acci_RecID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.acci_RecID
    acci_RecID = mvaracci_RecID
End Property



