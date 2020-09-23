VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFilesIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

Private mvarobjFTP As Object 'local copy
Public Enum enumIODirection
    SendToServer
    GetFromServer
End Enum
Private mvarIOState As enumIODirection 'local copy
Private mvarFilename As String 'local copy
Private mvarLocalPath As String 'local copy
Private mvarRemotePath As String 'local copy
Private mvarVirtualID As Long 'local copy
Private mvarSysopID As Long 'local copy
Private mvarDateStarted As Date 'local copy
Private mvarDateFinished As Date 'local copy
Public Enum enumIOStatus
    Waiting
    Transferring
    Idle
    Finished
End Enum
Private mvarIOStatus As enumIOStatus 'local copy
Private mvarEstimatedTimeRemaining As String 'local copy
Private mvarFileSize As Long 'local copy
Private mvarBytesRetrieved As Long 'local copy
Private mvarIdle As Boolean 'local copy
Private mvarFinished As Boolean 'local copy
Public Property Let Finished(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Finished = 5
    mvarFinished = vData
End Property


Public Property Get Finished() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Finished
    Finished = mvarFinished
End Property



Public Property Let Idle(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Idle = 5
    mvarIdle = vData
End Property


Public Property Get Idle() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Idle
    Idle = mvarIdle
End Property



Public Property Let BytesRetrieved(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BytesRetrieved = 5
    mvarBytesRetrieved = vData
End Property


Public Property Get BytesRetrieved() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BytesRetrieved
    BytesRetrieved = mvarBytesRetrieved
End Property



Public Property Let FileSize(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileSize = 5
    mvarFileSize = vData
End Property


Public Property Get FileSize() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileSize
    FileSize = mvarFileSize
End Property



Public Property Let EstimatedTimeRemaining(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EstimatedTimeRemaining = 5
    mvarEstimatedTimeRemaining = vData
End Property


Public Property Get EstimatedTimeRemaining() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EstimatedTimeRemaining
    EstimatedTimeRemaining = mvarEstimatedTimeRemaining
End Property



Public Property Let IOStatus(ByVal vData As enumIOStatus)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IOStatus = 5
    mvarIOStatus = vData
End Property


Public Property Get IOStatus() As enumIOStatus
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IOStatus
    Set IOStatus = mvarIOStatus
End Property




Public Property Let DateFinished(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateFinished = 5
    mvarDateFinished = vData
End Property


Public Property Get DateFinished() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateFinished
    DateFinished = mvarDateFinished
End Property



Public Property Let DateStarted(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateStarted = 5
    mvarDateStarted = vData
End Property


Public Property Get DateStarted() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateStarted
    DateStarted = mvarDateStarted
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



Public Property Let RemotePath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RemotePath = 5
    mvarRemotePath = vData
End Property


Public Property Get RemotePath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RemotePath
    RemotePath = mvarRemotePath
End Property



Public Property Let LocalPath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LocalPath = 5
    mvarLocalPath = vData
End Property


Public Property Get LocalPath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LocalPath
    LocalPath = mvarLocalPath
End Property



Public Property Let Filename(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Filename = 5
    mvarFilename = vData
End Property


Public Property Get Filename() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Filename
    Filename = mvarFilename
End Property



Public Property Let IOState(ByVal vData As enumIODirection)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IOState = 5
    mvarIOState = vData
End Property


Public Property Get IOState() As enumIODirection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IOState
    Set IOState = mvarIOState
End Property




Public Property Set objFTP(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.objFTP = Form1
    Set mvarobjFTP = vData
End Property


Public Property Get objFTP() As Object
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.objFTP
    Set objFTP = mvarobjFTP
End Property



