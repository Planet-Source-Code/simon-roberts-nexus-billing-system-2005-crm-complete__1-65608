VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_subReferals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvaracci_RecID As Long 'local copy
Private mvarRecID As Long 'local copy
Private mvaracci_RecID2 As Long 'local copy
Private mvarFlagID As Integer 'local copy
Private mvarDateAdded As Date 'local copy
Private mvarContactName As String 'local copy
Private mvarShortNote As String 'local copy
Private mvarCancelled As Integer 'local copy
Private mvarChecked As Integer 'local copy
Private mvaracciServiceID As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarAccountName As String 'local copy
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
Public Property Let AccountName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AccountName = 5
    mvarAccountName = vData
End Property


Public Property Get AccountName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AccountName
    AccountName = mvarAccountName
End Property



Public Property Let acciServiceID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.acciServiceID = 5
    mvaracciServiceID = vData
End Property


Public Property Get acciServiceID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.acciServiceID
    acciServiceID = mvaracciServiceID
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



Public Property Let Cancelled(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Cancelled = 5
    mvarCancelled = vData
End Property


Public Property Get Cancelled() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Cancelled
    Cancelled = mvarCancelled
End Property



Public Property Let ShortNote(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ShortNote = 5
    mvarShortNote = vData
End Property


Public Property Get ShortNote() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ShortNote
    ShortNote = mvarShortNote
End Property



Public Property Let ContactName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ContactName = 5
    mvarContactName = vData
End Property


Public Property Get ContactName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ContactName
    ContactName = mvarContactName
End Property



Public Property Let DateAdded(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateAdded = 5
    mvarDateAdded = vData
End Property


Public Property Get DateAdded() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateAdded
    DateAdded = mvarDateAdded
End Property



Public Property Let FlagID(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FlagID = 5
    mvarFlagID = vData
End Property


Public Property Get FlagID() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FlagID
    FlagID = mvarFlagID
End Property



Public Property Let acci_RecID2(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.acci_RecID2 = 5
    mvaracci_RecID2 = vData
End Property


Public Property Get acci_RecID2() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.acci_RecID2
    acci_RecID2 = mvaracci_RecID2
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



