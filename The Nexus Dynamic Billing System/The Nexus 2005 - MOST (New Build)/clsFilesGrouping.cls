VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFilesGrouping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"colFilesFolderMap"
Attribute VB_Ext_KEY = "Member1" ,"colFilesTree"
Public Key As String

Public Enum enumFileAllocation
    fa_VendorFile
    fa_SubscriberFile
    fa_ResellerFile
    fa_Invoice
    fa_Refund
    fa_Commission
    fa_Administration
    fa_Sysops
End Enum

Private mvarFileAllocation As enumFileAllocation 'local copy
Private mvarGroupingFolder As String 'local copy
Private mvarIDX As Long 'local copy
Private mvarSESSION As String 'local copy
Private mvarSubFolder As String 'local copy
Private mvarVirtualID As Long 'local copy
Private mvarClientID As Long 'local copy
Private mvarVendorID As Long 'local copy
Private mvarRefundID As Long 'local copy
Private mvarSysopID As Long 'local copy
Private mvarCommsID As Long 'local copy
Private mvarSerializedFolder As String 'local copy
Private mvarLastQuery As Object 'local copy
Public Sub FetchDirStructure(Optional sKey As String, Optional FolderIDX As Long, Optional FileIDX As Long)
End Sub

Public Property Set LastQuery(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.LastQuery = Form1
    Set mvarLastQuery = vData
End Property


Public Property Get LastQuery() As Object
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastQuery
    Set LastQuery = mvarLastQuery
End Property



Public Property Let SerializedFolder(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SerializedFolder = 5
    mvarSerializedFolder = vData
End Property


Public Property Get SerializedFolder() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SerializedFolder
    SerializedFolder = mvarSerializedFolder
End Property



Public Property Let CommsID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CommsID = 5
    mvarCommsID = vData
End Property


Public Property Get CommsID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CommsID
    CommsID = mvarCommsID
End Property



Public Property Let SysopID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SysopID = 5
    mvarSysopID = vData
End Property


Public Property Get SysopID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SysopID
    SysopID = mvarSysopID
End Property



Public Property Let RefundID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RefundID = 5
    mvarRefundID = vData
End Property


Public Property Get RefundID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RefundID
    RefundID = mvarRefundID
End Property



Public Property Let VendorID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VendorID = 5
    mvarVendorID = vData
End Property


Public Property Get VendorID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VendorID
    VendorID = mvarVendorID
End Property



Public Property Let ClientID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ClientID = 5
    mvarClientID = vData
End Property


Public Property Get ClientID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ClientID
    ClientID = mvarClientID
End Property



Public Property Let VirtualID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VirtualID = 5
    mvarVirtualID = vData
End Property


Public Property Get VirtualID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VirtualID
    VirtualID = mvarVirtualID
End Property



Public Property Let SubFolder(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SubFolder = 5
    mvarSubFolder = vData
End Property


Public Property Get SubFolder() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SubFolder
    SubFolder = mvarSubFolder
End Property



Public Property Let SESSION(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SESSION = 5
    mvarSESSION = vData
End Property


Public Property Get SESSION() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SESSION
    SESSION = mvarSESSION
End Property



Public Property Let IDX(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IDX = 5
    mvarIDX = vData
End Property


Public Property Get IDX() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IDX
    IDX = mvarIDX
End Property



Public Property Let GroupingFolder(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GroupingFolder = 5
    mvarGroupingFolder = vData
End Property


Public Property Get GroupingFolder() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GroupingFolder
    GroupingFolder = mvarGroupingFolder
End Property



Public Property Let FileAllocation(ByVal vData As enumFileAllocation)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileAllocation = 5
    mvarFileAllocation = vData
End Property


Public Property Get FileAllocation() As enumFileAllocation
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileAllocation
    Set FileAllocation = mvarFileAllocation
End Property



