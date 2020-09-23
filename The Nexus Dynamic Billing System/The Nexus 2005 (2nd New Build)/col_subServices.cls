VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "col_subServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"cls_subServices"
Attribute VB_Ext_KEY = "Member0" ,"cls_subServices"
'local variable to hold collection
Private mCol As Collection

Public Function Add(Key As String, RecID As Long, ptRecID As Long, ServiceID As Long, ContactName As String, Username As String, Password As String, NextCycle As Date, BaseURL As String, RadiusID As Long, DateCreated As Date, acci_RecID As Long, DynamicField1 As String, DynamicField2 As String, DynamicField3 As String, DynamicField4 As String, DynamicField5 As String, Checked As Integer, VirtualID As Long, PreviousCycle As Date, DomainID As Long, SubRecID As Long, MBQuota As Integer, Activation As Date, PeriodFee As Single, PerHour As Single, PerMB As Single, JoiningFee As Single, AgencyID As Long, DefaultShippingID As Long, ContractID As Long, ContractExpiry As Date, ActivationSet As Byte, Optional sKey As String) As cls_subServices
    'create a new object
    Dim objNewMember As cls_subServices
    Set objNewMember = New cls_subServices


    'set the properties passed into the method
    objNewMember.Key = Key
    objNewMember.RecID = RecID
    objNewMember.ptRecID = ptRecID
    objNewMember.ServiceID = ServiceID
    objNewMember.ContactName = ContactName
    objNewMember.Username = Username
    objNewMember.Password = Password
    objNewMember.NextCycle = NextCycle
    objNewMember.BaseURL = BaseURL
    objNewMember.RadiusID = RadiusID
    objNewMember.DateCreated = DateCreated
    objNewMember.acci_RecID = acci_RecID
    objNewMember.DynamicField1 = DynamicField1
    objNewMember.DynamicField2 = DynamicField2
    objNewMember.DynamicField3 = DynamicField3
    objNewMember.DynamicField4 = DynamicField4
    objNewMember.DynamicField5 = DynamicField5
    objNewMember.Checked = Checked
    objNewMember.VirtualID = VirtualID
    objNewMember.PreviousCycle = PreviousCycle
    objNewMember.DomainID = DomainID
    objNewMember.SubRecID = SubRecID
    objNewMember.MBQuota = MBQuota
    objNewMember.Activation = Activation
    objNewMember.PeriodFee = PeriodFee
    objNewMember.PerHour = PerHour
    objNewMember.PerMB = PerMB
    objNewMember.JoiningFee = JoiningFee
    objNewMember.AgencyID = AgencyID
    objNewMember.DefaultShippingID = DefaultShippingID
    objNewMember.ContractID = ContractID
    objNewMember.ContractExpiry = ContractExpiry
    objNewMember.ActivationSet = ActivationSet
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

Public Property Get Item(vntIndexKey As Variant) As cls_subServices
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property

Public Function Commit(ByVal oConn As ADODB.Connection) As Long

    Dim SubRecID As Long
    Dim RecID As Long
    
    If oConn.State = adStateOpen Then
        If Me.Count > 0 Then
            Dim oAls As cls_subServices
            For Each oAls In Me
                bUpdate = False
                If Left(oAls.Key, 3) = "NEW" Then
                    Commit = Commit + 1: bUpdate = True
                    If lastSession <> Mid(oAls.Key, 4, 14) Then
                        lastSession = Mid(oAls.Key, 4, 14)
                        
                        RecID = MySQL.GetTMPRecID("acci_services", oConn, "RecID")
                        If oAls.SubRecID = oAls.IDX Then
                            SubRecID = RecID
                        End If
                        oAls.Key = "RECID" & RecID
                    Else
                        oAls.Key = "RECID" & MySQL.GetTMPRecID("acci_services", oConn, "RecID")
                    End If
                ElseIf Left(oAls.Key, 5) = "UPDAT" Then
                    Commit = Commit + 1: bUpdate = True
                End If
                
                If bUpdate = True Then
                    SQL = "update acci_services set"
                    SQL = SQL + " acci_RecID = '" & oAls.acci_RecID
                    SQL = SQL + "', Activation = '" & Format(oAls.Activation, "yyyy-mm-dd ttttt")
                    SQL = SQL + "', ActivationSet = '" & oAls.ActivationSet
                    SQL = SQL + "', AgencyID = '" & oAls.AgencyID
                    SQL = SQL + "', BaseURL = '" & MySQL.ESC(oAls.BaseURL)
                    SQL = SQL + "', Checked = '" & oAls.Checked
                    SQL = SQL + "', ContactName = '" & MySQL.ESC(oAls.ContactName)
                    SQL = SQL + "', ContractExpiry = '" & oAls.ContractExpiry
                    SQL = SQL + "', ContractID = '" & oAls.ContractID
                    SQL = SQL + "', DefaultShippingID = '" & oAls.DefaultShippingID
                    SQL = SQL + "', DomainID = '" & oAls.DomainID
                    SQL = SQL + "', DynamicField1 = '" & MySQL.ESC(oAls.DynamicField1)
                    SQL = SQL + "', DynamicField2 = '" & MySQL.ESC(oAls.DynamicField2)
                    SQL = SQL + "', DynamicField3 = '" & MySQL.ESC(oAls.DynamicField3)
                    SQL = SQL + "', DynamicField4 = '" & MySQL.ESC(oAls.DynamicField4)
                    SQL = SQL + "', DynamicField5 = '" & MySQL.ESC(oAls.DynamicField5)
                    SQL = SQL + "', JoiningFee = '" & oAls.JoiningFee
                    SQL = SQL + "', MBQuota = '" & oAls.MBQuota
                    SQL = SQL + "', NextCycle = '" & Format(oAls.NextCycle, "yyyy-mm-dd ttttt")
                    SQL = SQL + "', Password = AES_ENCRYPT('" & MySQL.ESC(oAls.Password) & "','" + odb.colSalts.ReturnSalt(PWSalt) + "')"
                    SQL = SQL + ", PerHour = '" & oAls.PerHour
                    SQL = SQL + "', PeriodFee = '" & oAls.PeriodFee
                    SQL = SQL + "', PerMB = '" & oAls.PerMB
                    SQL = SQL + "', PreviousCycle = '" & Format(oAls.PreviousCycle, "yyyy-mm-dd ttttt")
                    SQL = SQL + "', ptRecID = '" & oAls.ptRecID
                    SQL = SQL + "', RadiusID = '" & oAls.RadiusID
                    SQL = SQL + "', ServiceID = '" & oAls.ServiceID
                    SQL = SQL + "', SubRecID = '" & SubRecID
                    SQL = SQL + "', Username = '" & MySQL.ESC(oAls.Username)
                    SQL = SQL + "', VirtualID = '" & oAls.VirtualID
                    SQL = SQL + "' WHERE  RecID = '" & Mid(oAls.Key, 6) & "'"
                    
                    MySQL.Execute oConn, CStr(SQL)
                    
                    oAls.Key = "RECID" & Mid(oAls.Key, 6)
                End If
                
                
            Next
        End If
    End If

End Function

Public Function GetClient(ByVal ClientID As Long, ByVal oConn As ADODB.Connection) As Byte

   Const bNoReturn = 0
    Const bResult = 1
    
    Dim rsload As ADODB.Recordset
    Dim lx As Long
    Dim bFound As Boolean
    
    If MySQL.OpenTable(oConn, rsload, , "select acci_services.*, AES_DECRYPT(acci_services.Password,'" & odb.colSalts.ReturnSalt(PWSalt) & "') as PassDec, plantypes.Description, plantypes.CatNo from acci_services, plantypes Where plantypes.RecID = acci_services.ptRecID AND AccI_RecID =  '" & ClientID & "'") = True Then
        If rsload.State = adStateOpen Then
            If rsload.RecordCount > 0 Then
                While Not rsload.EOF
                    bFound = False
                    For lx = 1 To Me.Count
                        If Me(lx).Key = "RECID" & rsload!RecID Then
                            bFound = True
                        End If
                    Next
                    
                    If bFound = False Then
                        With Me.Add("RECID" & rsload!RecID, rsload!RecID, IIf(IsNull(rsload!ptRecID), 0, rsload!ptRecID), IIf(IsNull(rsload!ServiceID), 0, rsload!ServiceID), IIf(IsNull(rsload!ContactName), "", rsload!ContactName), IIf(IsNull(rsload!Username), "guest" & Round(Rnd * 999999), rsload!Username), _
                                IIf(IsNull(rsload!PassDec), "myvoice", rsload!PassDec), IIf(IsNull(rsload!NextCycle), DateAdd("d", 1, sysnow), rsload!NextCycle), IIf(IsNull(rsload!BaseURL), "", rsload!BaseURL), IIf(IsNull(rsload!RadiusID), 0, rsload!RadiusID), IIf(IsNull(rsload!DateCreated), sysnow, rsload!DateCreated), _
                                IIf(IsNull(rsload!acci_RecID), ClientID, rsload!acci_RecID), IIf(IsNull(rsload!DynamicField1), "", rsload!DynamicField1), IIf(IsNull(rsload!DynamicField2), "", rsload!DynamicField2), IIf(IsNull(rsload!DynamicField3), "", rsload!DynamicField3), IIf(IsNull(rsload!DynamicField4), "", rsload!DynamicField4), _
                                IIf(IsNull(rsload!DynamicField5), "", rsload!DynamicField5), IIf(IsNull(rsload!Checked), True, rsload!Checked), IIf(IsNull(rsload!VirtualID), 0, rsload!VirtualID), IIf(IsNull(rsload!PreviousCycle), DateAdd("m", -1, sysnow), rsload!PreviousCycle), IIf(IsNull(rsload!DomainID), 0, rsload!DomainID), IIf(IsNull(rsload!SubRecID), 0, rsload!SubRecID), _
                                IIf(IsNull(rsload!MBQuota), 20, rsload!MBQuota), IIf(IsNull(rsload!Activation), DateAdd("d", 12, sysnow), rsload!Activation), IIf(IsNull(rsload!PeriodFee), 0, rsload!PeriodFee), IIf(IsNull(rsload!PerHour), 0, rsload!PerHour), IIf(IsNull(rsload!PerMB), 0, rsload!PerMB), _
                                IIf(IsNull(rsload!JoiningFee), 0, rsload!JoiningFee), IIf(IsNull(rsload!AgencyID), 0, rsload!AgencyID), IIf(IsNull(rsload!DefaultShippingID), 0, rsload!DefaultShippingID), IIf(IsNull(rsload!ContractID), 0, rsload!ContractID), IIf(IsNull(rsload!ContractExpiry), DateAdd("d", -1, sysnow), rsload!ContractExpiry), Val(rsload!ActivationSet), "RECID" & rsload!RecID)
                                
                            .Description = IIf(IsNull(rsload!Description), "::", rsload!Description)
                            .ProductCode = IIf(IsNull(rsload!CatNo), "", rsload!CatNo)
                        End With
                    End If
                    rsload.MoveNext
                Wend
                GetClient = bResult
            End If
        End If
    End If
End Function



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

