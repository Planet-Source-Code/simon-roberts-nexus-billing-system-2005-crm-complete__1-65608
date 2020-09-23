VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "col_subTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"cls_subTrans"
Attribute VB_Ext_KEY = "Collection" ,"UserDefinedCollection"
'local variable to hold collection
Private mCol As Collection
'local variable(s) to hold property value(s)
Private mvarttl_TotalPaid As Currency 'local copy
Private mvarttl_TotalDue As Currency 'local copy
Private mvarttl_TotalTax As Currency 'local copy
Private mvarttl_TotalCredited As Currency 'local copy
Private mvarttl_TotalCreditedTax As Currency 'local copy
Public Sub calcTotals()
End Sub

Public Property Let ttl_TotalCreditedTax(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ttl_TotalCreditedTax = 5
    mvarttl_TotalCreditedTax = vData
End Property


Public Property Get ttl_TotalCreditedTax() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ttl_TotalCreditedTax
    ttl_TotalCreditedTax = mvarttl_TotalCreditedTax
End Property



Public Property Let ttl_TotalCredited(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ttl_TotalCredited = 5
    mvarttl_TotalCredited = vData
End Property


Public Property Get ttl_TotalCredited() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ttl_TotalCredited
    ttl_TotalCredited = mvarttl_TotalCredited
End Property



Public Property Let ttl_TotalTax(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ttl_TotalTax = 5
    mvarttl_TotalTax = vData
End Property


Public Property Get ttl_TotalTax() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ttl_TotalTax
    ttl_TotalTax = mvarttl_TotalTax
End Property



Public Property Let ttl_TotalDue(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ttl_TotalDue = 5
    mvarttl_TotalDue = vData
End Property


Public Property Get ttl_TotalDue() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ttl_TotalDue
    ttl_TotalDue = mvarttl_TotalDue
End Property



Public Property Let ttl_TotalPaid(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ttl_TotalPaid = 5
    mvarttl_TotalPaid = vData
End Property


Public Property Get ttl_TotalPaid() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ttl_TotalPaid
    ttl_TotalPaid = mvarttl_TotalPaid
End Property



Public Function Commit(ByVal oConn As ADODB.Connection) As Long

    Dim SubRecID As Long
    Dim lastSession As String
    
    If oConn.State = adStateOpen Then
        If Me.Count > 0 Then
            Dim oAls As cls_subTrans
            For Each oAls In Me
                bUpdate = False
                If Left(oAls.Key, 3) = "NEW" Or Left(oAls.Key, 3) = "EML" Then
                    Commit = Commit + 1: bUpdate = True
                    If lastSession <> Mid(oAls.Key, 4, 14) Then
                        If Left(oAls.Key, 3) = "EML" Then
                            oAls.Key = "RECID" & MySQL.GetTMPRecID("invoiceout", oConn, "RecID")
                        Else
                            lastSession = Mid(oAls.Key, 4, 14)
                            SubRecID = MySQL.GetTMPRecID("invoiceout", oConn, "RecID")
                            oAls.Key = "RECID" & SubRecID
                        End If
                    Else
                        oAls.Key = "RECID" & MySQL.GetTMPRecID("invoiceout", oConn, "RecID")
                    End If
                ElseIf Left(oAls.Key, 5) = "UPDAT" Then
                    Commit = Commit + 1: bUpdate = True
                End If
                
                If bUpdate = True Then
                    SQL = "update invoiceout set"
                    SQL = SQL + " AccI_RecID = '" & oAls.acci_RecID
                    SQL = SQL + "', AgencyID = '" & oAls.AgencyID
                    SQL = SQL + "', AmountDue = '" & oAls.AmountDue
                    SQL = SQL + "', AmountPaid = '" & oAls.AmountPaid
                    SQL = SQL + "', AmountRefunded = '" & oAls.AmountRefunded
                    SQL = SQL + "', TotalDue = '" & oAls.TotalDue
                    SQL = SQL + "', Checked = '" & oAls.Checked
                    SQL = SQL + "', Description = '" & MySQL.ESC(oAls.Description)
                    SQL = SQL + "', DomainID = '" & oAls.DomainID
                    SQL = SQL + "', EndCycle = '" & Format(oAls.EndCycle, "yyyy-mm-dd ttttt")
                    SQL = SQL + "', FlagID = '" & oAls.FlagID
                    SQL = SQL + "', GSTCharged = '" & oAls.GSTCharged
                    SQL = SQL + "', GSTRefunded = '" & oAls.GSTRefunded
                    SQL = SQL + "', InvoiceFlagID = '" & oAls.InvoiceFlagID
                    SQL = SQL + "', PaidWhen = '" & Format(oAls.PaidWhen, "yyyy-mm-dd ttttt")
                    SQL = SQL + "', PaymentDue = '" & Format(oAls.PaymentDue, "yyyy-mm-dd ttttt")
                    SQL = SQL + "', PlanServiceID = '" & oAls.PlanServiceID
                    SQL = SQL + "', ptRecID = '" & oAls.ptRecID
                    SQL = SQL + "', RadiusID = '" & oAls.RadiusID
                    SQL = SQL + "', ServiceID = '" & oAls.ServiceID
                    SQL = SQL + "', RefundID = '" & oAls.RefundID
                    SQL = SQL + "', StartCycle = '" & Format(oAls.StartCycle, "yyyy-mm-dd ttttt")
                    SQL = SQL + "', sfCycle_Download = '" & oAls.sfCycle_Download
                    SQL = SQL + "', sfCycle_Mins = '" & oAls.sfCycle_Mins
                    SQL = SQL + "', sfCycle_Upload = '" & oAls.sfCycle_Upload
                    SQL = SQL + "', StatementID = '" & oAls.StatementID
                    SQL = SQL + "', SubRecID  = '" & IIf(SubRecID <> 0, SubRecID, oAls.SubRecID)
                    SQL = SQL + "', VISPStatementID = '" & oAls.VISPStatementID
                    
                    SQL = SQL + "' WHERE  RecID = '" & Mid(oAls.Key, 6) & "'"
                    
                    MySQL.Execute oConn, CStr(SQL)
                    
                    oAls.Key = "RECID" & Mid(oAls.Key, 6)
                End If
                
                
'                MsgBox SQL
            Next
        End If
    End If


End Function

Public Function GetClient(ByVal ClientID As Long, ByVal oConn As ADODB.Connection, Optional TraxrID As Long = 0, Optional RefundID As Long = 0, Optional SQLCLause As String = "::") As Byte

 Const bNoReturn = 0
    Const bResult = 1
    
    Dim rsload As ADODB.Recordset
    Dim lx As Long
    Dim bFound As Boolean
    Dim SQL As String
    
    If SQLCLause = "::" Then
        If TraxrID <> 0 Then
            SQL = "select * from invoiceout Where TraxrID = " & TraxrID
        ElseIf RefundID <> 0 Then
            SQL = "select * from invoiceout Where RefundID = " & RefundID
        Else
            SQL = "select * from invoiceout Where acci_RecID = " & ClientID
        End If
    Else
        SQL = SQLCLause
    End If
    
    If MySQL.OpenTable(oConn, rsload, , SQL) = True Then
        If rsload.State = adStateOpen Then
        
            If rsload.RecordCount > 0 Then
                While Not rsload.EOF
                    bFound = False
'                    For lx = 1 To Me.Count
'
'
'                            If Me(lx).Key = "RECID" & rsLoad!RecID Then
'                                bFound = True
'                            End If
'
'                    Next
'
                    If bFound = False Then
                    
                        With Me.Add("RECID" & rsload!RecID, rsload!RecID, IIf(IsNull(rsload!acci_RecID), ClientID, rsload!acci_RecID), IIf(IsNull(rsload!AmountDue), 0, rsload!AmountDue), IIf(IsNull(rsload!GSTCharged), 0, rsload!GSTCharged), IIf(IsNull(rsload!PaymentDue), DateAdd("d", 2, sysnow), rsload!PaymentDue), _
                                IIf(IsNull(rsload!AmountPaid), 0, rsload!AmountPaid), IIf(IsNull(rsload!PaidWhen), sysnow, rsload!PaidWhen), IIf(IsNull(rsload!Checked), True, Val(rsload!Checked)), IIf(IsNull(rsload!FlagID), 0, Val(rsload!FlagID)), IIf(IsNull(rsload!TotalDue), 0, rsload!TotalDue), _
                                IIf(IsNull(rsload!sfCycle_Upload), 0, rsload!sfCycle_Upload), IIf(IsNull(rsload!sfCycle_Download), 0, rsload!sfCycle_Download), IIf(IsNull(rsload!sfCycle_Mins), True, Val(rsload!sfCycle_Mins)), IIf(IsNull(rsload!AgencyID), 0, Val(rsload!AgencyID)), IIf(IsNull(rsload!VirtualID), Login.lVirtualID, rsload!VirtualID), _
                                IIf(IsNull(rsload!Description), 0, rsload!Description), IIf(IsNull(rsload!TraxrID), 0, rsload!TraxrID), IIf(IsNull(rsload!RefundID), True, Val(rsload!RefundID)), IIf(IsNull(rsload!PlanServiceID), 0, Val(rsload!PlanServiceID)), IIf(IsNull(rsload!AmountRefunded), 0, rsload!AmountRefunded), _
                                IIf(IsNull(rsload!GSTRefunded), 0, rsload!GSTRefunded), IIf(IsNull(rsload!SysopID), Login.lSysopID, rsload!SysopID), IIf(IsNull(rsload!Created), sysnow, rsload!Created), IIf(IsNull(rsload!SubRecID), 0, Val(rsload!SubRecID)), IIf(IsNull(rsload!VISPStatementID), 0, rsload!VISPStatementID), _
                                IIf(IsNull(rsload!StatementID), 0, rsload!StatementID), IIf(IsNull(rsload!InvoiceFlagID), 0, rsload!InvoiceFlagID), IIf(IsNull(rsload!StartCycle), DateAdd("d", -32, sysnow), rsload!StartCycle), IIf(IsNull(rsload!EndCycle), DateAdd("d", -1, sysnow), rsload!EndCycle), IIf(IsNull(rsload!ptRecID), 0, rsload!ptRecID), _
                                IIf(IsNull(rsload!ServiceID), 0, rsload!ServiceID), IIf(IsNull(rsload!DomainID), sysnow, rsload!DomainID), IIf(IsNull(rsload!RadiusID), 0, rsload!RadiusID), "RECID" & rsload!RecID)
                            
                        End With
                    End If
                    rsload.MoveNext
                Wend
                GetClient = bResult
            End If
        End If
    End If
    
End Function


Public Function Add(Key As String, RecID As Long, acci_RecID As Long, AmountDue As Currency, GSTCharged As Currency, PaymentDue As Date, AmountPaid As Currency, PaidWhen As Date, Checked As Integer, FlagID As Integer, TotalDue As Currency, sfCycle_Upload As Double, sfCycle_Download As Double, sfCycle_Mins As Double, AgencyID As Long, VirtualID As Long, Description As String, TraxrID As Long, RefundID As Long, PlanServiceID As Long, AmountRefunded As Currency, GSTRefunded As Currency, SysopID As Long, Created As Date, SubRecID As Long, VISPStatementID As Long, StatementID As Long, InvoiceFlagID As Integer, StartCycle As Date, EndCycle As Date, ptRecID As Long, ServiceID As Long, DomainID As Long, RadiusID As Long, Optional sKey As String) As cls_subTrans
    'create a new object
    Dim objNewMember As cls_subTrans
    Set objNewMember = New cls_subTrans


    'set the properties passed into the method
    objNewMember.Key = Key
    objNewMember.RecID = RecID
    objNewMember.acci_RecID = acci_RecID
    objNewMember.AmountDue = AmountDue
    objNewMember.GSTCharged = GSTCharged
    objNewMember.PaymentDue = PaymentDue
    objNewMember.AmountPaid = AmountPaid
    objNewMember.PaidWhen = PaidWhen
    objNewMember.Checked = Checked
    objNewMember.FlagID = FlagID
    objNewMember.TotalDue = TotalDue
    objNewMember.sfCycle_Upload = sfCycle_Upload
    objNewMember.sfCycle_Download = sfCycle_Download
    objNewMember.sfCycle_Mins = sfCycle_Mins
    objNewMember.AgencyID = AgencyID
    objNewMember.VirtualID = VirtualID
    objNewMember.Description = Description
    objNewMember.TraxrID = TraxrID
    objNewMember.RefundID = RefundID
    objNewMember.PlanServiceID = PlanServiceID
    objNewMember.AmountRefunded = AmountRefunded
    objNewMember.GSTRefunded = GSTRefunded
    objNewMember.SysopID = SysopID
    objNewMember.Created = Created
    objNewMember.SubRecID = SubRecID
    objNewMember.VISPStatementID = VISPStatementID
    objNewMember.StatementID = StatementID
    objNewMember.InvoiceFlagID = InvoiceFlagID
    objNewMember.StartCycle = StartCycle
    objNewMember.EndCycle = EndCycle
    objNewMember.ptRecID = ptRecID
    objNewMember.ServiceID = ServiceID
    objNewMember.DomainID = DomainID
    objNewMember.RadiusID = RadiusID
    objNewMember.IDX = Me.Count + 1
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As cls_subTrans
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

