VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFilesLogicGate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Public Key As String

Private mvarIcon As String 'local copy
Private mvarFolderIDX As Long 'local copy
Private mvarFileIDX As Variant 'local copy
Public Property Let FileIDX(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileIDX = 5
    mvarFileIDX = vData
End Property


Public Property Set FileIDX(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.FileIDX = Form1
    Set mvarFileIDX = vData
End Property


Public Property Get FileIDX() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileIDX
    If IsObject(mvarFileIDX) Then
        Set FileIDX = mvarFileIDX
    Else
        FileIDX = mvarFileIDX
    End If
End Property



Public Property Let FolderIDX(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FolderIDX = 5
    mvarFolderIDX = vData
End Property


Public Property Get FolderIDX() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FolderIDX
    FolderIDX = mvarFolderIDX
End Property



Public Property Let Icon(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Icon = 5
    mvarIcon = vData
End Property


Public Property Get Icon() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Icon
    Icon = mvarIcon
End Property



