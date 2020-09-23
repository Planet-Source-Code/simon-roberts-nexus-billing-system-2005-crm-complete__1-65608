VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFilesFolderMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

Private mvarIDX As Long 'local copy
Private mvarParentFolder As String 'local copy
Private mvarChildFolder As String 'local copy
Private mvarDepth As Integer 'local copy
Private mvarToBeDeleted As Boolean 'local copy
Private mvarHasBeenArchived As Boolean 'local copy
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



Public Property Let Depth(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Depth = 5
    mvarDepth = vData
End Property


Public Property Get Depth() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Depth
    Depth = mvarDepth
End Property



Public Property Let ChildFolder(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ChildFolder = 5
    mvarChildFolder = vData
End Property


Public Property Get ChildFolder() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ChildFolder
    ChildFolder = mvarChildFolder
End Property



Public Property Let ParentFolder(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ParentFolder = 5
    mvarParentFolder = vData
End Property


Public Property Get ParentFolder() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ParentFolder
    ParentFolder = mvarParentFolder
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



