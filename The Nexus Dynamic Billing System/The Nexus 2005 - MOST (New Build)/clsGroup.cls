VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"colInvoices"
Private mvarcolInvoices As colInvoices


Public Property Get colInvoices() As colInvoices
    If mvarcolInvoices Is Nothing Then
        Set mvarcolInvoices = New colInvoices
    End If


    Set colInvoices = mvarcolInvoices
End Property


Public Property Set colInvoices(vData As colInvoices)
    Set mvarcolInvoices = vData
End Property
Private Sub Class_Terminate()
    Set mvarcolInvoices = Nothing
End Sub



