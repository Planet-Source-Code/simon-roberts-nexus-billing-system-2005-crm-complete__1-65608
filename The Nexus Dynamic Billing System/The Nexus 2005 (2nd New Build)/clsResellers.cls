VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsResellers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"colReseller"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"


Private mvarcolReseller As colReseller




Public Property Get colReseller() As colReseller
    If mvarcolReseller Is Nothing Then
        Set mvarcolReseller = New colReseller
    End If


    Set colReseller = mvarcolReseller
End Property


Public Property Set colReseller(vData As colReseller)
    Set mvarcolReseller = vData
End Property
Private Sub Class_Terminate()
    Set mvarcolReseller = Nothing
End Sub



