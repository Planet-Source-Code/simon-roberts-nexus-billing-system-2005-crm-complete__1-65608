VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvarVirtualID As Long 'local copy
Private mvaracciRecID As Long 'local copy
Private mvarInvRecID As Long 'local copy
Private mvarDescription As String 'local copy
Private mvarDebit As Currency 'local copy
Private mvarCredit As Currency 'local copy
Private mvarPaid As Currency 'local copy
Private mvarChecked As Boolean 'local copy
Public Property Let Checked(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Checked = 5
    mvarChecked = vData
End Property


Public Property Get Checked() As Boolean
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Checked
    Checked = mvarChecked
End Property



Public Property Let Paid(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Paid = 5
    mvarPaid = vData
End Property


Public Property Get Paid() As Currency
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Paid
    Paid = mvarPaid
End Property



Public Property Let Credit(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Credit = 5
    mvarCredit = vData
End Property


Public Property Get Credit() As Currency
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Credit
    Credit = mvarCredit
End Property



Public Property Let Debit(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Debit = 5
    mvarDebit = vData
End Property


Public Property Get Debit() As Currency
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Debit
    Debit = mvarDebit
End Property



Public Property Let Description(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Description = 5
    mvarDescription = vData
End Property


Public Property Get Description() As String
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    Description = mvarDescription
End Property



Public Property Let InvRecID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.InvRecID = 5
    mvarInvRecID = vData
End Property


Public Property Get InvRecID() As Long
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.InvRecID
    InvRecID = mvarInvRecID
End Property



Public Property Let acciRecID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.acciRecID = 5
    mvaracciRecID = vData
End Property


Public Property Get acciRecID() As Long
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.acciRecID
    acciRecID = mvaracciRecID
End Property



Public Property Let VirtualID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VirtualID = 5
    mvarVirtualID = vData
End Property


Public Property Get VirtualID() As Long
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VirtualID
    VirtualID = mvarVirtualID
End Property



