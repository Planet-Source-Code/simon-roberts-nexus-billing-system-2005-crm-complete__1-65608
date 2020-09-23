VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"clsSpiff"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"colNames"
Attribute VB_Ext_KEY = "Member2" ,"colRates"
Private mvarcolRates As colRates
Private mvarcolNames As colNames
Private mvarclsSpiff As clsSpiff


Private Sub Class_Initialize()
    'create the mclsSpiff object when the clsCommission class is created
    Set mvarclsSpiff = New clsSpiff
End Sub


Public Property Get clsSpiff() As clsSpiff
    Set clsSpiff = mvarclsSpiff
End Property


Public Property Set clsSpiff(vData As clsSpiff)
    Set mvarclsSpiff = vData
End Property

Public Sub LoadDB()
End Sub



Public Property Get colNames() As colNames
    If mvarcolNames Is Nothing Then
        Set mvarcolNames = New colNames
    End If


    Set colNames = mvarcolNames
End Property


Public Property Set colNames(vData As colNames)
    Set mvarcolNames = vData
End Property



Public Property Get colRates() As colRates
    If mvarcolRates Is Nothing Then
        Set mvarcolRates = New colRates
    End If


    Set colRates = mvarcolRates
End Property


Public Property Set colRates(vData As colRates)
    Set mvarcolRates = vData
End Property
Private Sub Class_Terminate()
  Set mvarclsSpiff = Nothing
  Set mvarcolNames = Nothing
    Set mvarcolRates = Nothing
End Sub



