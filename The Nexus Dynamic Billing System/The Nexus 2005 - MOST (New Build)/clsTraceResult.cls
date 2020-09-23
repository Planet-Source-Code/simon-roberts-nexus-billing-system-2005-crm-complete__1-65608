VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTraceResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvarIP As String 'local copy
Private mvarOne As String 'local copy
Private mvarTwo As String 'local copy
Private mvarThree As String 'local copy
Private mvarIpfrom As String 'local copy
Private mvarIndexIP As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarHOP As Byte 'local copy
Public Property Let HOP(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HOP = 5
    mvarHOP = vData
End Property


Public Property Get HOP() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HOP
    HOP = mvarHOP
End Property



Public Property Let IndexIP(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IndexIP = 5
    mvarIndexIP = vData
End Property


Public Property Get IndexIP() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IndexIP
    IndexIP = mvarIndexIP
End Property



Public Property Let Ipfrom(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Ipfrom = 5
    mvarIpfrom = vData
End Property


Public Property Get Ipfrom() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Ipfrom
    Ipfrom = mvarIpfrom
End Property



Public Property Let Three(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Three = 5
    mvarThree = vData
End Property


Public Property Get Three() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Three
    Three = mvarThree
End Property



Public Property Let Two(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Two = 5
    mvarTwo = vData
End Property


Public Property Get Two() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Two
    Two = mvarTwo
End Property



Public Property Let One(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.One = 5
    mvarOne = vData
End Property


Public Property Get One() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.One
    One = mvarOne
End Property



Public Property Let IP(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IP = 5
    mvarIP = vData
End Property


Public Property Get IP() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IP
    IP = mvarIP
End Property



