VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSpiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"colSpiff"
Attribute VB_Ext_KEY = "Member1" ,"colUnitType"
Private mvarcolSpiff As colSpiff
Private mvarcolUnitType As colUnitType


Public Property Get colUnitType() As colUnitType
    If mvarcolUnitType Is Nothing Then
        Set mvarcolUnitType = New colUnitType
    End If


    Set colUnitType = mvarcolUnitType
End Property


Public Property Set colUnitType(vData As colUnitType)
    Set mvarcolUnitType = vData
End Property



Public Property Get colSpiff() As colSpiff
    If mvarcolSpiff Is Nothing Then
        Set mvarcolSpiff = New colSpiff
    End If


    Set colSpiff = mvarcolSpiff
End Property


Public Property Set colSpiff(vData As colSpiff)
    Set mvarcolSpiff = vData
End Property
Private Sub Class_Terminate()
  Set mvarcolUnitType = Nothing
    Set mvarcolSpiff = Nothing
End Sub



