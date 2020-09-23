VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvarTypeID As Byte 'local copy
Private mvarAmount As Variant 'local copy
Private mvarSerialNumber As String 'local copy
Private mvarccID As Long 'local copy
Public Property Let CCID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ccID = 5
    mvarccID = vData
End Property


Public Property Get CCID() As Long
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ccID
    CCID = mvarccID
End Property



Public Property Let SerialNumber(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SerialNumber = 5
    mvarSerialNumber = vData
End Property


Public Property Get SerialNumber() As String
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SerialNumber
    SerialNumber = mvarSerialNumber
End Property



Public Property Let Amount(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Amount = 5
    mvarAmount = vData
End Property


Public Property Set Amount(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Amount = Form1
    Set mvarAmount = vData
End Property


Public Property Get Amount() As Variant
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Amount
    If IsObject(mvarAmount) Then
        Set Amount = mvarAmount
    Else
        Amount = mvarAmount
    End If
End Property



Public Property Let TypeID(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TypeID = 5
    mvarTypeID = vData
End Property


Public Property Get TypeID() As Byte
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TypeID
    TypeID = mvarTypeID
End Property



