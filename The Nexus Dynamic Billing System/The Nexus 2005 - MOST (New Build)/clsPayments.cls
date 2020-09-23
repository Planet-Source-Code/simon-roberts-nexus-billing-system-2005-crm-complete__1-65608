VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"colPayments"
Private mvarcolPayments As colPayments


Public Property Get colPayments() As colPayments
    If mvarcolPayments Is Nothing Then
        Set mvarcolPayments = New colPayments
    End If


    Set colPayments = mvarcolPayments
End Property


Public Property Set colPayments(vData As colPayments)
    Set mvarcolPayments = vData
End Property
Private Sub Class_Terminate()
    Set mvarcolPayments = Nothing
End Sub



