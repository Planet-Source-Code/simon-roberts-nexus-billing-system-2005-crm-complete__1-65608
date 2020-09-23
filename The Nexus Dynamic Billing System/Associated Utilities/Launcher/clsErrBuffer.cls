VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsErrBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvarCaseStatement As String 'local copy
'local variable(s) to hold property value(s)
Private mvarErrNumber As Double 'local copy
'local variable(s) to hold property value(s)
Private mvarRoutineName As String 'local copy
Private mvarContainerName As String 'local copy
'local variable(s) to hold property value(s)
Private mvarLBL As String 'local copy
Private mvarDateTime As Date 'local copy
Public Property Let DateTime(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateTime = 5
    mvarDateTime = vData
End Property


Public Property Get DateTime() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateTime
    DateTime = mvarDateTime
End Property



Public Property Let LBL(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LBL = 5
    mvarLBL = vData
End Property


Public Property Get LBL() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LBL
    LBL = mvarLBL
End Property



Public Property Let ContainerName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ContainerName = 5
    mvarContainerName = vData
End Property


Public Property Get ContainerName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ContainerName
    ContainerName = mvarContainerName
End Property



Public Property Let RoutineName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RoutineName = 5
    mvarRoutineName = vData
End Property


Public Property Get RoutineName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RoutineName
    RoutineName = mvarRoutineName
End Property



Public Property Let ErrNumber(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ErrNumber = 5
    mvarErrNumber = vData
End Property


Public Property Get ErrNumber() As Double
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ErrNumber
    ErrNumber = mvarErrNumber
End Property



Public Property Let CaseStatement(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CaseStatement = 5
    mvarCaseStatement = vData
End Property


Public Property Get CaseStatement() As String
'used when retrueving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CaseStatement
    CaseStatement = mvarCaseStatement
End Property



