VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAuditTrail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvarsystemstamp As Date 'local copy
Private mvarsysnow As Date 'local copy
Private mvarlocaltime As Date 'local copy
Private mvarappname As String 'local copy
Private mvarappversion As String 'local copy
Private mvarapphdc As Long 'local copy
Private mvarformname As String 'local copy
Private mvarformhwnd As Long 'local copy
Private mvarvartype As String 'local copy
Private mvarvarname As String 'local copy
Private mvaroldcur As Single 'local copy
Private mvarnewcur As Single 'local copy
Private mvaroldvalue As String 'local copy
Private mvarnewvalue As String 'local copy
Private mvaroldpointer As Double 'local copy
Private mvarnewpointer As Double 'local copy
Private mvarIDX As Long 'local copy
Private mvarSysopID As Long 'local copy
Private mvarVirtualID As Long 'local copy
Private mvarAgencyID As Long 'local copy
Private mvaracci_RecID As Long 'local copy
Private mvarRefundID As Long 'local copy
Private mvarinvtrxrid As Long 'local copy
Private mvarFlagID As Long 'local copy
Private mvarDescription As String 'local copy
Private mvarChecked As Integer 'local copy
Public Property Let Checked(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Checked = 5
    mvarChecked = vData
End Property


Public Property Get Checked() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Checked
    Checked = mvarChecked
End Property



Public Property Let Description(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Description = 5
    mvarDescription = vData
End Property


Public Property Get Description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    Description = mvarDescription
End Property



Public Property Let FlagID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.flagid = 5
    mvarFlagID = vData
End Property


Public Property Get FlagID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.flagid
    FlagID = mvarFlagID
End Property



Public Property Let invtrxrid(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.invtrxrid = 5
    mvarinvtrxrid = vData
End Property


Public Property Get invtrxrid() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.invtrxrid
    invtrxrid = mvarinvtrxrid
End Property



Public Property Let RefundID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.refundid = 5
    mvarRefundID = vData
End Property


Public Property Get RefundID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.refundid
    RefundID = mvarRefundID
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



Public Property Let AgencyID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.agencyid = 5
    mvarAgencyID = vData
End Property


Public Property Get AgencyID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.agencyid
    AgencyID = mvarAgencyID
End Property



Public Property Let VirtualID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.virtualid = 5
    mvarVirtualID = vData
End Property


Public Property Get VirtualID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.virtualid
    VirtualID = mvarVirtualID
End Property



Public Property Let SysopID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sysopid = 5
    mvarSysopID = vData
End Property


Public Property Get SysopID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sysopid
    SysopID = mvarSysopID
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



Public Property Let newpointer(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.newpointer = 5
    mvarnewpointer = vData
End Property


Public Property Get newpointer() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.newpointer
    newpointer = mvarnewpointer
End Property



Public Property Let oldpointer(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.oldpointer = 5
    mvaroldpointer = vData
End Property


Public Property Get oldpointer() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.oldpointer
    oldpointer = mvaroldpointer
End Property



Public Property Let newvalue(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.newvalue = 5
    mvarnewvalue = vData
End Property


Public Property Get newvalue() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.newvalue
    newvalue = mvarnewvalue
End Property



Public Property Let oldvalue(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.oldvalue = 5
    mvaroldvalue = vData
End Property


Public Property Get oldvalue() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.oldvalue
    oldvalue = mvaroldvalue
End Property



Public Property Let newcur(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.newcur = 5
    mvarnewcur = vData
End Property


Public Property Get newcur() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.newcur
    newcur = mvarnewcur
End Property



Public Property Let oldcur(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.oldcur = 5
    mvaroldcur = vData
End Property


Public Property Get oldcur() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.oldcur
    oldcur = mvaroldcur
End Property



Public Property Let varname(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.varname = 5
    mvarvarname = vData
End Property


Public Property Get varname() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.varname
    varname = mvarvarname
End Property



Public Property Let vartype(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.vartype = 5
    mvarvartype = vData
End Property


Public Property Get vartype() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.vartype
    vartype = mvarvartype
End Property



Public Property Let formhwnd(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.formhwnd = 5
    mvarformhwnd = vData
End Property


Public Property Get formhwnd() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.formhwnd
    formhwnd = mvarformhwnd
End Property



Public Property Let formname(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.formname = 5
    mvarformname = vData
End Property


Public Property Get formname() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.formname
    formname = mvarformname
End Property



Public Property Let apphdc(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.apphdc = 5
    mvarapphdc = vData
End Property


Public Property Get apphdc() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.apphdc
    apphdc = mvarapphdc
End Property



Public Property Let appversion(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.appversion = 5
    mvarappversion = vData
End Property


Public Property Get appversion() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.appversion
    appversion = mvarappversion
End Property



Public Property Let appname(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.appname = 5
    mvarappname = vData
End Property


Public Property Get appname() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.appname
    appname = mvarappname
End Property



Public Property Let localtime(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.localtime = 5
    mvarlocaltime = vData
End Property


Public Property Get localtime() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.localtime
    localtime = mvarlocaltime
End Property



Public Property Let sysnow(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sysnow = 5
    mvarsysnow = vData
End Property


Public Property Get sysnow() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sysnow
    sysnow = mvarsysnow
End Property



Public Property Let systemstamp(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.systemstamp = 5
    mvarsystemstamp = vData
End Property


Public Property Get systemstamp() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.systemstamp
    systemstamp = mvarsystemstamp
End Property



