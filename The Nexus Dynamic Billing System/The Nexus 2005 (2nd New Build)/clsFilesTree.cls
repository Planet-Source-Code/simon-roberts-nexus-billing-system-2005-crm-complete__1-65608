VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFilesTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

Private mvarFilename As String 'local copy
Private mvarMD5 As String 'local copy
Private mvarFileSize As Long 'local copy
Private mvarIDX As Long 'local copy
Private mvarSESSION As String 'local copy
Private mvarLocalPath As String 'local copy
Private mvarLocalMD5 As String 'local copy
Private mvarExtension As String 'local copy
Private mvarRemotePath As String 'local copy
Private mvarVirtualID As Long 'local copy
Private mvarSysopID As Long 'local copy
Private mvarDateSubmitted As Date 'local copy
Private mvarPercentage As Single 'local copy
Private mvarToBeDeleted As Boolean 'local copy
Private mvarHasBeenArchived As Boolean 'local copy
Private mvarDateAction As Date 'local copy
Private mvarDateFileCreated As Date 'local copy
Private mvarDateLastRequested As Date 'local copy
Private mvarFFMIDX As Long 'local copy
Private mvarRecID As Long 'local copy
Public Property Let RecID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RecID = 5
    mvarRecID = vData
End Property


Public Property Get RecID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RecID
    RecID = mvarRecID
End Property



Public Property Let FFMIDX(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FFMIDX = 5
    mvarFFMIDX = vData
End Property


Public Property Get FFMIDX() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FFMIDX
    FFMIDX = mvarFFMIDX
End Property



Public Property Let DateLastRequested(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateLastRequested = 5
    mvarDateLastRequested = vData
End Property


Public Property Get DateLastRequested() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateLastRequested
    DateLastRequested = mvarDateLastRequested
End Property



Public Property Let DateFileCreated(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateFileCreated = 5
    mvarDateFileCreated = vData
End Property


Public Property Get DateFileCreated() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateFileCreated
    DateFileCreated = mvarDateFileCreated
End Property



Public Property Let DateAction(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateAction = 5
    mvarDateAction = vData
End Property


Public Property Get DateAction() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateAction
    DateAction = mvarDateAction
End Property



Public Property Let HasBeenArchived(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HasBeenArchived = 5
    mvarHasBeenArchived = vData
End Property


Public Property Get HasBeenArchived() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HasBeenArchived
    HasBeenArchived = mvarHasBeenArchived
End Property



Public Property Let ToBeDeleted(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ToBeDeleted = 5
    mvarToBeDeleted = vData
End Property


Public Property Get ToBeDeleted() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ToBeDeleted
    ToBeDeleted = mvarToBeDeleted
End Property



Public Property Let Percentage(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Percentage = 5
    mvarPercentage = vData
End Property


Public Property Get Percentage() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Percentage
    Percentage = mvarPercentage
End Property



Public Property Let DateSubmitted(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateSubmitted = 5
    mvarDateSubmitted = vData
End Property


Public Property Get DateSubmitted() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateSubmitted
    DateSubmitted = mvarDateSubmitted
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



Public Property Let Extension(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Extension = 5
    mvarExtension = vData
End Property


Public Property Get Extension() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Extension
    Extension = mvarExtension
End Property



Public Property Let LocalMD5(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LocalMD5 = 5
    mvarLocalMD5 = vData
End Property


Public Property Get LocalMD5() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LocalMD5
    LocalMD5 = mvarLocalMD5
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



Public Property Let MD5(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MD5 = 5
    mvarMD5 = vData
End Property


Public Property Get MD5() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MD5
    MD5 = mvarMD5
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



