VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFiles_AllVersions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
'local variable(s) to hold property value(s)
Private mvarDatePutInArchives As Date 'local copy
Private mvarFileID As Double 'local copy
Private mvarFilename As String 'local copy
Private mvarOldFilename As String 'local copy
Private mvarMD5_Crypt As String 'local copy
Private mvarMD5_Orginal As String 'local copy
Private mvarsaltAES As String 'local copy
Private mvarsaltRC4 As String 'local copy
Private mvarusedAES As Boolean 'local copy
Private mvarusedRC4 As Boolean 'local copy
Private mvarusedZIP As Boolean 'local copy
Private mvarServerPath As String 'local copy
Private mvarVersion As Long 'local copy
Private mvarMaskedFolder As String 'local copy
Private mvarSysopID As Long 'local copy
Private mvarFlagID As Integer 'local copy
Private mvarScheduledRollbackDate As Date 'local copy
Private mvarScheduledRollbackVersion As Long 'local copy
Private mvarSysopID_RequestingRollback As Long 'local copy
Private mvarRollBackASAP As Boolean 'local copy
Public EditStatus As enum_FileDB_EditStates
Public Key As String


Public Property Let RollBackASAP(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RollBackASAP = 5
    mvarRollBackASAP = vData
End Property


Public Property Get RollBackASAP() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RollBackASAP
    RollBackASAP = mvarRollBackASAP
End Property



Public Property Let SysopID_RequestingRollback(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SysopID_RequestingRollback = 5
    mvarSysopID_RequestingRollback = vData
End Property


Public Property Get SysopID_RequestingRollback() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SysopID_RequestingRollback
    SysopID_RequestingRollback = mvarSysopID_RequestingRollback
End Property



Public Property Let ScheduledRollbackVersion(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ScheduledRollbackVersion = 5
    mvarScheduledRollbackVersion = vData
End Property


Public Property Get ScheduledRollbackVersion() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ScheduledRollbackVersion
    ScheduledRollbackVersion = mvarScheduledRollbackVersion
End Property



Public Property Let ScheduledRollbackDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ScheduledRollbackDate = 5
    mvarScheduledRollbackDate = vData
End Property


Public Property Get ScheduledRollbackDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ScheduledRollbackDate
    ScheduledRollbackDate = mvarScheduledRollbackDate
End Property



Public Property Let FlagID(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FlagID = 5
    mvarFlagID = vData
End Property


Public Property Get FlagID() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FlagID
    FlagID = mvarFlagID
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



Public Property Let MaskedFolder(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MaskedFolder = 5
    mvarMaskedFolder = vData
End Property


Public Property Get MaskedFolder() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MaskedFolder
    MaskedFolder = mvarMaskedFolder
End Property



Public Property Let Version(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Version = 5
    mvarVersion = vData
End Property


Public Property Get Version() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Version
    Version = mvarVersion
End Property



Public Property Let ServerPath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ServerPath = 5
    mvarServerPath = vData
End Property


Public Property Get ServerPath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ServerPath
    ServerPath = mvarServerPath
End Property



Public Property Let usedZIP(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.usedZIP = 5
    mvarusedZIP = vData
End Property


Public Property Get usedZIP() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.usedZIP
    usedZIP = mvarusedZIP
End Property



Public Property Let usedRC4(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.usedRC4 = 5
    mvarusedRC4 = vData
End Property


Public Property Get usedRC4() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.usedRC4
    usedRC4 = mvarusedRC4
End Property



Public Property Let usedAES(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.usedAES = 5
    mvarusedAES = vData
End Property


Public Property Get usedAES() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.usedAES
    usedAES = mvarusedAES
End Property



Public Property Let saltRC4(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.saltRC4 = 5
    mvarsaltRC4 = vData
End Property


Public Property Get saltRC4() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.saltRC4
    saltRC4 = mvarsaltRC4
End Property



Public Property Let saltAES(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.saltAES = 5
    mvarsaltAES = vData
End Property


Public Property Get saltAES() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.saltAES
    saltAES = mvarsaltAES
End Property



Public Property Let MD5_Orginal(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MD5_Orginal = 5
    mvarMD5_Orginal = vData
End Property


Public Property Get MD5_Orginal() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MD5_Orginal
    MD5_Orginal = mvarMD5_Orginal
End Property



Public Property Let MD5_Crypt(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MD5_Crypt = 5
    mvarMD5_Crypt = vData
End Property


Public Property Get MD5_Crypt() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MD5_Crypt
    MD5_Crypt = mvarMD5_Crypt
End Property



Public Property Let oldFilename(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OldFilename = 5
    mvarOldFilename = vData
End Property


Public Property Get oldFilename() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OldFilename
    oldFilename = mvarOldFilename
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



Public Property Let FileID(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileID = 5
    mvarFileID = vData
End Property


Public Property Get FileID() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileID
    FileID = mvarFileID
End Property



Public Property Let DatePutInArchives(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DatePutInArchives = 5
    mvarDatePutInArchives = vData
End Property


Public Property Get DatePutInArchives() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DatePutInArchives
    DatePutInArchives = mvarDatePutInArchives
End Property



