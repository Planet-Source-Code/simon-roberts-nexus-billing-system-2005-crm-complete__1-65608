VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_subTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvarRecID As Long 'local copy
Private mvaracci_RecID As Long 'local copy
Private mvarAmountDue As Currency 'local copy
Private mvarGSTCharged As Currency 'local copy
Private mvarPaymentDue As Date 'local copy
Private mvarAmountPaid As Currency 'local copy
Private mvarPaidWhen As Date 'local copy
Private mvarChecked As Integer 'local copy
Private mvarFlagID As Integer 'local copy
Private mvarTotalDue As Currency 'local copy
Private mvarsfCycle_Upload As Double 'local copy
Private mvarsfCycle_Download As Double 'local copy
Private mvarsfCycle_Mins As Double 'local copy
Private mvarAgencyID As Long 'local copy
Private mvarVirtualID As Long 'local copy
Private mvarDescription As String 'local copy
Private mvarTraxrID As Long 'local copy
Private mvarRefundID As Long 'local copy
Private mvarPlanServiceID As Long 'local copy
Private mvarAmountRefunded As Currency 'local copy
Private mvarGSTRefunded As Currency 'local copy
Private mvarSysopID As Long 'local copy
Private mvarCreated As Date 'local copy
Private mvarSubRecID As Long 'local copy
Private mvarVISPStatementID As Long 'local copy
Private mvarStatementID As Long 'local copy
Private mvarInvoiceFlagID As Integer 'local copy
Private mvarStartCycle As Date 'local copy
Private mvarEndCycle As Date 'local copy
Private mvarptRecID As Long 'local copy
Private mvarServiceID As Long 'local copy
Private mvarDomainID As Long 'local copy
Private mvarRadiusID As Long 'local copy
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
Public Property Let RadiusID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RadiusID = 5
    mvarRadiusID = vData
End Property


Public Property Get RadiusID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RadiusID
    RadiusID = mvarRadiusID
End Property



Public Property Let DomainID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DomainID = 5
    mvarDomainID = vData
End Property


Public Property Get DomainID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DomainID
    DomainID = mvarDomainID
End Property



Public Property Let ServiceID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ServiceID = 5
    mvarServiceID = vData
End Property


Public Property Get ServiceID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ServiceID
    ServiceID = mvarServiceID
End Property



Public Property Let ptRecID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ptRecID = 5
    mvarptRecID = vData
End Property


Public Property Get ptRecID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ptRecID
    ptRecID = mvarptRecID
End Property



Public Property Let EndCycle(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EndCycle = 5
    mvarEndCycle = vData
End Property


Public Property Get EndCycle() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EndCycle
    EndCycle = mvarEndCycle
End Property



Public Property Let StartCycle(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StartCycle = 5
    mvarStartCycle = vData
End Property


Public Property Get StartCycle() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StartCycle
    StartCycle = mvarStartCycle
End Property



Public Property Let InvoiceFlagID(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.InvoiceFlagID = 5
    mvarInvoiceFlagID = vData
End Property


Public Property Get InvoiceFlagID() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.InvoiceFlagID
    InvoiceFlagID = mvarInvoiceFlagID
End Property



Public Property Let StatementID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StatementID = 5
    mvarStatementID = vData
End Property


Public Property Get StatementID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StatementID
    StatementID = mvarStatementID
End Property



Public Property Let VISPStatementID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VISPStatementID = 5
    mvarVISPStatementID = vData
End Property


Public Property Get VISPStatementID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VISPStatementID
    VISPStatementID = mvarVISPStatementID
End Property



Public Property Let SubRecID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SubRecID = 5
    mvarSubRecID = vData
End Property


Public Property Get SubRecID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SubRecID
    SubRecID = mvarSubRecID
End Property



Public Property Let Created(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Created = 5
    mvarCreated = vData
End Property


Public Property Get Created() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Created
    Created = mvarCreated
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



Public Property Let GSTRefunded(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GSTRefunded = 5
    mvarGSTRefunded = vData
End Property


Public Property Get GSTRefunded() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GSTRefunded
    GSTRefunded = mvarGSTRefunded
End Property



Public Property Let AmountRefunded(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AmountRefunded = 5
    mvarAmountRefunded = vData
End Property


Public Property Get AmountRefunded() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AmountRefunded
    AmountRefunded = mvarAmountRefunded
End Property



Public Property Let PlanServiceID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PlanServiceID = 5
    mvarPlanServiceID = vData
End Property


Public Property Get PlanServiceID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PlanServiceID
    PlanServiceID = mvarPlanServiceID
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



Public Property Let TraxrID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TraxrID = 5
    mvarTraxrID = vData
End Property


Public Property Get TraxrID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TraxrID
    TraxrID = mvarTraxrID
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



Public Property Let AgencyID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AgencyID = 5
    mvarAgencyID = vData
End Property


Public Property Get AgencyID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AgencyID
    AgencyID = mvarAgencyID
End Property



Public Property Let sfCycle_Mins(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sfCycle_Mins = 5
    mvarsfCycle_Mins = vData
End Property


Public Property Get sfCycle_Mins() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sfCycle_Mins
    sfCycle_Mins = mvarsfCycle_Mins
End Property



Public Property Let sfCycle_Download(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sfCycle_Download = 5
    mvarsfCycle_Download = vData
End Property


Public Property Get sfCycle_Download() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sfCycle_Download
    sfCycle_Download = mvarsfCycle_Download
End Property



Public Property Let sfCycle_Upload(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sfCycle_Upload = 5
    mvarsfCycle_Upload = vData
End Property


Public Property Get sfCycle_Upload() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sfCycle_Upload
    sfCycle_Upload = mvarsfCycle_Upload
End Property



Public Property Let TotalDue(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TotalDue = 5
    mvarTotalDue = vData
End Property


Public Property Get TotalDue() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TotalDue
    TotalDue = mvarTotalDue
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



Public Property Let PaidWhen(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PaidWhen = 5
    mvarPaidWhen = vData
End Property


Public Property Get PaidWhen() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PaidWhen
    PaidWhen = mvarPaidWhen
End Property



Public Property Let AmountPaid(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AmountPaid = 5
    mvarAmountPaid = vData
End Property


Public Property Get AmountPaid() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AmountPaid
    AmountPaid = mvarAmountPaid
End Property



Public Property Let PaymentDue(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PaymentDue = 5
    mvarPaymentDue = vData
End Property


Public Property Get PaymentDue() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PaymentDue
    PaymentDue = mvarPaymentDue
End Property



Public Property Let GSTCharged(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GSTCharged = 5
    mvarGSTCharged = vData
End Property


Public Property Get GSTCharged() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GSTCharged
    GSTCharged = mvarGSTCharged
End Property



Public Property Let AmountDue(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AmountDue = 5
    mvarAmountDue = vData
End Property


Public Property Get AmountDue() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AmountDue
    AmountDue = mvarAmountDue
End Property



Public Property Let acci_RecID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AccI_RecID = 5
    mvaracci_RecID = vData
End Property


Public Property Get acci_RecID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AccI_RecID
    acci_RecID = mvaracci_RecID
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



