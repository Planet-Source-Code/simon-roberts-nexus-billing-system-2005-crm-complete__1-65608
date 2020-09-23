VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFlags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"colFlags_IconCache"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarcolFlags_IconCache As colFlags_IconCache


Public Property Get colFlags_IconCache() As colFlags_IconCache
    If mvarcolFlags_IconCache Is Nothing Then
        Set mvarcolFlags_IconCache = New colFlags_IconCache
    End If


    Set colFlags_IconCache = mvarcolFlags_IconCache
End Property


Public Property Set colFlags_IconCache(vData As colFlags_IconCache)
    Set mvarcolFlags_IconCache = vData
End Property
Private Sub Class_Terminate()
    Set mvarcolFlags_IconCache = Nothing
End Sub



