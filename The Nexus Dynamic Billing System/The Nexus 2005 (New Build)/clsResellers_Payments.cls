VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsResellers_Payments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

Private mvarbNEW As Boolean 'local copy
Private mvarIDX As Long 'local copy
Private mvarSESSION As String 'local copy
Private mvarVirtualID As Long 'local copy
Private mvarSysopID As Long 'local copy
Private mvarInvoiceID As Long 'local copy
Private mvarDateKeyed As Date 'local copy
Private mvarDateBanked As Date 'local copy
Private mvarRemittanceSent As Boolean 'local copy
Private mvarAmountPaid As Single 'local copy
Private mvarGSTPaid As Single 'local copy
Private mvarFlagID As Integer 'local copy
Private mvarRecCreated As Date 'local copy
'local variable(s) to hold property value(s)
Private mvarFetchStatus As Byte 'local copy
Public Property Let FetchStatus(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FetchStatus = 5
    mvarFetchStatus = vData
End Property


Public Property Get FetchStatus() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FetchStatus
    FetchStatus = mvarFetchStatus
End Property






Public Property Let RecCreated(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RecCreated = 5
    mvarRecCreated = vData
End Property


Public Property Get RecCreated() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RecCreated
    RecCreated = mvarRecCreated
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



Public Property Let GSTPaid(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GSTPaid = 5
    mvarGSTPaid = vData
End Property


Public Property Get GSTPaid() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GSTPaid
    GSTPaid = mvarGSTPaid
End Property



Public Property Let AmountPaid(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AmountPaid = 5
    mvarAmountPaid = vData
End Property


Public Property Get AmountPaid() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AmountPaid
    AmountPaid = mvarAmountPaid
End Property



Public Property Let RemittanceSent(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RemittanceSent = 5
    mvarRemittanceSent = vData
End Property


Public Property Get RemittanceSent() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RemittanceSent
    RemittanceSent = mvarRemittanceSent
End Property



Public Property Let DateBanked(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateBanked = 5
    mvarDateBanked = vData
End Property


Public Property Get DateBanked() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateBanked
    DateBanked = mvarDateBanked
End Property



Public Property Let DateKeyed(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateKeyed = 5
    mvarDateKeyed = vData
End Property


Public Property Get DateKeyed() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateKeyed
    DateKeyed = mvarDateKeyed
End Property



Public Property Let InvoiceID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.InvoiceID = 5
    mvarInvoiceID = vData
End Property


Public Property Get InvoiceID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.InvoiceID
    InvoiceID = mvarInvoiceID
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



Public Property Let SESSION(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SESSION = 5
    mvarSESSION = vData
End Property


Public Property Get SESSION() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SESSION
    SESSION = mvarSESSION
End Property



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




