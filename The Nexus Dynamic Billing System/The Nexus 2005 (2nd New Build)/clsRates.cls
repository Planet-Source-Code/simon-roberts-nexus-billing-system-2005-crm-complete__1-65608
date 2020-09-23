VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvarSysopID As Double 'local copy
Private mvarptRecID As Double 'local copy
Private mvarMin As Integer 'local copy
Private mvarMax As Integer 'local copy
Private mvarRate As Currency 'local copy
'local variable(s) to hold property value(s)
'local variable(s) to hold property value(s)
Private mvarClassID As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarPercentile As Single 'local copy
'local variable(s) to hold property value(s)
Private mvarFlag As Byte 'local copy
Public Property Let Flag(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Flag = 5
    mvarFlag = vData
End Property


Public Property Get Flag() As Byte
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Flag
    Flag = mvarFlag
End Property



Public Property Let Percentile(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Percentile = 5
    mvarPercentile = vData
End Property


Public Property Get Percentile() As Single
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Percentile
    Percentile = mvarPercentile
End Property



Public Property Let ttlCommPer(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ttlCommPer = 5
    mvarttlCommPer = vData
End Property


Public Property Get ttlCommPer() As Single
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ttlCommPer
    ttlCommPer = mvarttlCommPer
End Property



Public Property Let Margin(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Margin = 5
    mvarMargin = vData
End Property


Public Property Set Margin(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Margin = Form1
    Set mvarMargin = vData
End Property


Public Property Get Margin() As Variant
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Margin
    If IsObject(mvarMargin) Then
        Set Margin = mvarMargin
    Else
        Margin = mvarMargin
    End If
End Property



Public Property Let Cost(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Cost = 5
    mvarCost = vData
End Property


Public Property Get Cost() As Currency
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Cost
    Cost = mvarCost
End Property



Public Property Let RRP(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RRP = 5
    mvarRRP = vData
End Property


Public Property Get RRP() As Currency
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RRP
    RRP = mvarRRP
End Property



Public Property Let ClassID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ClassID = 5
    mvarClassID = vData
End Property


Public Property Get ClassID() As Long
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ClassID
    ClassID = mvarClassID
End Property

Public Property Let Rate(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Rate = 5
    mvarRate = vData
End Property


Public Property Get Rate() As Currency
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Rate
    Rate = mvarRate
End Property



Public Property Let Max(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Max = 5
    mvarMax = vData
End Property


Public Property Get Max() As Integer
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Max
    Max = mvarMax
End Property



Public Property Let Min(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Min = 5
    mvarMin = vData
End Property


Public Property Get Min() As Integer
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Min
    Min = mvarMin
End Property



Public Property Let ptRecID(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ptRecID = 5
    mvarptRecID = vData
End Property


Public Property Get ptRecID() As Double
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ptRecID
    ptRecID = mvarptRecID
End Property




