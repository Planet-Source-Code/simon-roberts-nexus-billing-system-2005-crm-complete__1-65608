VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"colKeys"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarcolKeys As colKeys


Public Property Get colKeys() As colKeys
    If mvarcolKeys Is Nothing Then
        Set mvarcolKeys = New colKeys
    End If


    Set colKeys = mvarcolKeys
End Property


Public Property Set colKeys(vData As colKeys)
    Set mvarcolKeys = vData
End Property
Private Sub Class_Terminate()
    Set mvarcolKeys = Nothing
End Sub



