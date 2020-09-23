VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDBConstraints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"colInsert"
Private mvarcolInsert As colInsert
Public Function ContainsCommand(ByVal sBuffer As String) As Boolean

    If Me.colInsert.Count > 0 Then
    
        Dim ix As Long
        For ix = 1 To Me.colInsert.Count
            If InStr(LCase(sBuffer), LCase(Me.colInsert(ix).sCommand)) > 0 Then
                ContainsCommand = True
                Exit Function
            End If
        Next ix
    End If
    
    
End Function



Public Property Get colInsert() As colInsert
    If mvarcolInsert Is Nothing Then
        Set mvarcolInsert = New colInsert
    End If


    Set colInsert = mvarcolInsert
End Property


Public Property Set colInsert(vData As colInsert)
    Set mvarcolInsert = vData
End Property
Private Sub Class_Terminate()
    Set mvarcolInsert = Nothing
End Sub



