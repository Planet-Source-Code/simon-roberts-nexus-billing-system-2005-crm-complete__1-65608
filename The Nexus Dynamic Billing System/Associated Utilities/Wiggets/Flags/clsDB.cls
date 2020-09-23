VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"clsDBConstraints"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"colADOConn"
Attribute VB_Ext_KEY = "Member2" ,"colDBObjects"
Attribute VB_Ext_KEY = "Member3" ,"colSalts"
Private mvarcolDBObjects As colDBObjects
Private mvarclsdbConstraints As clsDBConstraints
Private mvarcolADOConn As colADOConn
Public Enum enumSoftwareAction
    PersonalsBrowser
    BrowseData
    AdminsCenter
    Query
End Enum
Private mvarcolSalts As colSalts


Public Property Get colSalts() As colSalts
    If mvarcolSalts Is Nothing Then
        Set mvarcolSalts = New colSalts
    End If


    Set colSalts = mvarcolSalts
End Property


Public Property Set colSalts(vData As colSalts)
    Set mvarcolSalts = vData
End Property

Public Function GetActionNodesFilled(ByVal fldName As String, ByVal tblName As String, Optional ByVal dbName As String, Optional NodeS As Node = Nothing, Optional tv As TreeView = Nothing, Optional Search As Boolean = True) As enumSoftwareAction

    Static LastdbName As String
    
    If NodeS Is Nothing Then
    
    
    Else
        
        If tblName = "personals" And dbName = "spraci" And fldName = "id" Then
            tv.NodeS.Add NodeS, tvwChild, "personals.spraci", "Load Personals Explorer", "project"
        End If
    
    End If
    
    
End Function




Public Property Get colADOConn() As colADOConn
    If mvarcolADOConn Is Nothing Then
        Set mvarcolADOConn = New colADOConn
    End If


    Set colADOConn = mvarcolADOConn
End Property


Public Property Set colADOConn(vData As colADOConn)
    Set mvarcolADOConn = vData
End Property



Private Sub Class_Initialize()
    'create the mclsdbConstraints object when the clsDB class is created
    Set mvarclsdbConstraints = New clsDBConstraints
End Sub


Public Property Get clsDBConstraints() As clsDBConstraints
    Set clsDBConstraints = mvarclsdbConstraints
End Property


Public Property Set clsDBConstraints(vData As clsDBConstraints)
    Set mvarclsdbConstraints = vData
End Property



Public Property Get colDBObjects() As colDBObjects
    If mvarcolDBObjects Is Nothing Then
        Set mvarcolDBObjects = New colDBObjects
    End If


    Set colDBObjects = mvarcolDBObjects
End Property


Public Property Set colDBObjects(vData As colDBObjects)
    Set mvarcolDBObjects = vData
End Property
Private Sub Class_Terminate()
  Set mvarcolSalts = Nothing
  Set mvarcolADOConn = Nothing
  Set mvarclsdbConstraints = Nothing
    Set mvarcolDBObjects = Nothing
End Sub



