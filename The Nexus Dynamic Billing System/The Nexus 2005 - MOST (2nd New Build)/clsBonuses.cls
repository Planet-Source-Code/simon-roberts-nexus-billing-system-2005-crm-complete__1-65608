VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBonuses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)

Private mvarMin As Long 'local copy
Private mvarUnitsAwarded As Long 'local copy
Private mvarVirtualID As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarptRecID As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarenumSource As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarSaveFlag As Byte 'local copy
'local variable(s) to hold property value(s)
Private mvarMax As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarUnitType As Integer 'local copy
Public Property Let UnitType(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UnitType = 5
    mvarUnitType = vData
End Property


Public Property Get UnitType() As Integer
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UnitType
    UnitType = mvarUnitType
End Property



Public Property Let Max(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Max = 5
    mvarMax = vData
End Property


Public Property Get Max() As Long
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Max
    Max = mvarMax
End Property



Public Property Let SaveFlag(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SaveFlag = 5
    mvarSaveFlag = vData
End Property


Public Property Get SaveFlag() As Byte
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SaveFlag
    SaveFlag = mvarSaveFlag
End Property



Public Property Let enumSource(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.enumSource = 5
    mvarenumSource = vData
End Property


Public Property Get enumSource() As Long
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.enumSource
    enumSource = mvarenumSource
End Property



Public Property Let ptRecID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ptRecID = 5
    mvarptRecID = vData
End Property


Public Property Get ptRecID() As Long
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ptRecID
    ptRecID = mvarptRecID
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



Public Property Let UnitsAwarded(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UnitsAwarded = 5
    mvarUnitsAwarded = vData
End Property


Public Property Get UnitsAwarded() As Long
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UnitsAwarded
    UnitsAwarded = mvarUnitsAwarded
End Property


Public Property Let Min(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Min = 5
    mvarMin = vData
End Property


Public Property Get Min() As Long
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Min
    Min = mvarMin
End Property







