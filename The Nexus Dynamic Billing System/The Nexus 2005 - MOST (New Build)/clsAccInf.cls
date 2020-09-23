VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAccInf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"colAccounts"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarcolAccounts As colAccounts


Public Property Get colAccounts() As colAccounts
    If mvarcolAccounts Is Nothing Then
        Set mvarcolAccounts = New colAccounts
    End If


    Set colAccounts = mvarcolAccounts
End Property


Public Property Set colAccounts(vData As colAccounts)
    Set mvarcolAccounts = vData
End Property
Private Sub Class_Terminate()
    Set mvarcolAccounts = Nothing
End Sub



