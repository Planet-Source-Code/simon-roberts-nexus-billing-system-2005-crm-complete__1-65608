VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_subHardware"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvarRecID As Long 'local copy
Private mvarModem As String 'local copy
Private mvarProcessor As String 'local copy
Private mvarVideoCard As String 'local copy
Private mvarMonitor As String 'local copy
Private mvarPCType As String 'local copy
Private mvarNetworkCard As String 'local copy
Private mvarOS As String 'local copy
Private mvarPrinter As String 'local copy
Private mvarMainboard As String 'local copy
Private mvaracci_RecID As Long 'local copy
Private mvarExtraXML As String 'local copy
'local variable(s) to hold property value(s)
Private mvaracciServiceID As Long 'local copy
Private mvarIDX As Long 'local copy
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
Public Property Let acciServiceID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.acciServiceID = 5
    mvaracciServiceID = vData
End Property


Public Property Get acciServiceID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.acciServiceID
    acciServiceID = mvaracciServiceID
End Property



Public Property Let ExtraXML(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ExtraXML = 5
    mvarExtraXML = vData
End Property


Public Property Get ExtraXML() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ExtraXML
    ExtraXML = mvarExtraXML
End Property



Public Property Let acci_RecID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.acci_RecID = 5
    mvaracci_RecID = vData
End Property


Public Property Get acci_RecID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.acci_RecID
    acci_RecID = mvaracci_RecID
End Property



Public Property Let Mainboard(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Mainboard = 5
    mvarMainboard = vData
End Property


Public Property Get Mainboard() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Mainboard
    Mainboard = mvarMainboard
End Property



Public Property Let Printer(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Printer = 5
    mvarPrinter = vData
End Property


Public Property Get Printer() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Printer
    Printer = mvarPrinter
End Property



Public Property Let OS(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OS = 5
    mvarOS = vData
End Property


Public Property Get OS() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OS
    OS = mvarOS
End Property



Public Property Let NetworkCard(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NetworkCard = 5
    mvarNetworkCard = vData
End Property


Public Property Get NetworkCard() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NetworkCard
    NetworkCard = mvarNetworkCard
End Property



Public Property Let PCType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PCType = 5
    mvarPCType = vData
End Property


Public Property Get PCType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PCType
    PCType = mvarPCType
End Property



Public Property Let Monitor(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Monitor = 5
    mvarMonitor = vData
End Property


Public Property Get Monitor() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Monitor
    Monitor = mvarMonitor
End Property



Public Property Let VideoCard(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VideoCard = 5
    mvarVideoCard = vData
End Property


Public Property Get VideoCard() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VideoCard
    VideoCard = mvarVideoCard
End Property



Public Property Let Processor(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Processor = 5
    mvarProcessor = vData
End Property


Public Property Get Processor() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Processor
    Processor = mvarProcessor
End Property



Public Property Let Modem(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Modem = 5
    mvarModem = vData
End Property


Public Property Get Modem() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Modem
    Modem = mvarModem
End Property



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



