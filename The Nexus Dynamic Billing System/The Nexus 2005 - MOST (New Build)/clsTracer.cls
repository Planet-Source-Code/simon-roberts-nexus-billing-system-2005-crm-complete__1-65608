VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTracer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"ColIPs"
Attribute VB_Ext_KEY = "Member1" ,"colResult"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarColIPs As ColIPs
Private mvarcolResult As colResult


Public Property Get colResult() As colResult
    If mvarcolResult Is Nothing Then
        Set mvarcolResult = New colResult
    End If


    Set colResult = mvarcolResult
End Property


Public Property Set colResult(vData As colResult)
    Set mvarcolResult = vData
End Property



Public Property Get ColIPs() As ColIPs
    If mvarColIPs Is Nothing Then
        Set mvarColIPs = New ColIPs
    End If


    Set ColIPs = mvarColIPs
End Property


Public Property Set ColIPs(vData As ColIPs)
    Set mvarColIPs = vData
End Property
Private Sub Class_Terminate()
  Set mvarcolResult = Nothing
    Set mvarColIPs = Nothing
End Sub



