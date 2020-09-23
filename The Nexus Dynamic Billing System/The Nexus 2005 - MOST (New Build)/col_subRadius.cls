VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "col_subRadius"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"cls_subRadius"
Attribute VB_Ext_KEY = "Collection" ,"UserDefinedCollection"
'local variable to hold collection
Private mCol As Collection
Private mvarcls_subRadius As cls_subRadius


Public Property Get cls_subRadius() As cls_subRadius
    Set cls_subRadius = mvarcls_subRadius
End Property


Public Property Set cls_subRadius(vData As cls_subRadius)
    Set mvarcls_subRadius = vData
End Property

Public Function Commit(ByVal oConn As ADODB.Connection) As Long
    
    Dim bUpdate As Boolean
    Dim SQL As String
    
    If oConn.State = adStateOpen Then
        If Me.Count > 0 Then
            Dim oAls As cls_subRadius
            For Each oAls In Me
                bUpdate = False
                If Left(oAls.Key, 3) = "NEW" Then
                    Commit = Commit + 1: bUpdate = True
                    oAls.Key = "RECID" & MySQL.GetTMPRecID("radiusaccounts", oConn, "RecID")
                ElseIf Left(oAls.Key, 5) = "UPDAT" Then
                    Commit = Commit + 1: bUpdate = True
                End If
                
                If bUpdate = True Then
                    SQL = "update radiusaccounts set"
                    SQL = SQL + " acci_RecID = '" & oAls.acci_RecID
                    SQL = SQL + "', Acct_Session_ID = '" & oAls.Acct_Session_ID
                    SQL = SQL + "', Activate = '" & oAls.Activate
                    SQL = SQL + "', AutoActivateFlag = '" & oAls.AutoActivateFlag
                    SQL = SQL + "', Checked = '" & oAls.Checked
                    SQL = SQL + "', Deactivate = '" & oAls.Deactivate
                    SQL = SQL + "', FramedProtocol = '" & oAls.FramedProtocol
                    SQL = SQL + "', IdleTimeout = '" & oAls.IdleTimeout
                    SQL = SQL + "', Password = AES_ENCRYPT('" & MySQL.ESC(oAls.Password) & "','" & odb.colSalts.ReturnSalt(PWSalt) & "')"
                    SQL = SQL + ", PrimaryDNS = '" & MySQL.ESC(oAls.PrimaryDNS)
                    SQL = SQL + "', ptRecID = '" & oAls.ptRecID
                    SQL = SQL + "', SecondaryDNS = '" & MySQL.ESC(oAls.SecondaryDNS)
                    SQL = SQL + "', ServiceType = '" & oAls.ServiceType
                    SQL = SQL + "', SessionsAllowed = '" & oAls.SessionsAllowed
                    SQL = SQL + "', SessionTimeout = '" & oAls.SessionTimeout
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
    
    If MySQL.OpenTable(oConn, rsload, , "select *, AES_DECRYPT(Password,'" + odb.colSalts.ReturnSalt(PWSalt) + "') as PassDec from radiusaccounts where acci_RecID = '" & ClientID & "'") = True Then
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
                        Call Me.Add("RECID" & rsload!RecID, rsload!RecID, IIf(IsNull(rsload!Username), "guest" & Round(Rnd(99999)), rsload!Username), _
                                IIf(IsNull(rsload!PassDec), "myvoice", rsload!PassDec), IIf(IsNull(rsload!SessionsAllowed), 1, rsload!SessionsAllowed), _
                                IIf(IsNull(rsload!AutoActivateFlag), 0, rsload!AutoActivateFlag), IIf(IsNull(rsload!Activate), CDate(Format(Now, "yyyy-mm-dd ") & "07:00 AM"), rsload!Activate), _
                                IIf(IsNull(rsload!Deactivate), CDate(Format(Now, "yyyy-mm-dd ") & "6:30 PM"), rsload!Deactivate), IIf(IsNull(rsload!SessionTimeout), 0, rsload!SessionTimeout), _
                                IIf(IsNull(rsload!IdleTimeout), 0, rsload!IdleTimeout), IIf(IsNull(rsload!acci_RecID), ClientID, rsload!acci_RecID), _
                                IIf(IsNull(rsload!ServiceType), "", rsload!ServiceType), IIf(IsNull(rsload!FramedProtocol), "", rsload!FramedProtocol), _
                                IIf(IsNull(rsload!Checked), False, Val(rsload!Checked)), IIf(IsNull(rsload!sfStartTime), sysnow, rsload!sfStartTime), _
                                IIf(IsNull(rsload!sfStopTime), sysnow, rsload!sfStopTime), IIf(IsNull(rsload!sfAliveTime), sysnow, rsload!sfAliveTime), _
                                IIf(IsNull(rsload!sfCycle_Upload), 0, rsload!sfCycle_Upload), IIf(IsNull(rsload!sfCycle_Download), 0, rsload!sfCycle_Download), _
                                IIf(IsNull(rsload!sfCycle_Mins), 0, rsload!sfCycle_Mins), IIf(IsNull(rsload!Acct_Session_ID), "", rsload!Acct_Session_ID), _
                                IIf(IsNull(rsload!ptRecID), 0, rsload!ptRecID), IIf(IsNull(rsload!VirtualID), 0, rsload!VirtualID), _
                                IIf(IsNull(rsload!PrimaryDNS), "", rsload!PrimaryDNS), IIf(IsNull(rsload!SecondaryDNS), "", rsload!SecondaryDNS), _
                                IIf(IsNull(rsload!DateCreated), sysnow, rsload!DateCreated), "RECID" & rsload!RecID)
                                
                    End If
                    rsload.MoveNext
                Wend
                GetClient = bResult
            End If
        End If
    End If
End Function


Public Function Add(Key As String, RecID As Long, Username As String, Password As String, SessionsAllowed As Integer, AutoActivateFlag As Byte, Activate As Date, Deactivate As Date, SessionTimeout As Long, IdleTimeout As Long, acci_RecID As Long, ServiceType As String, FramedProtocol As String, Checked As Integer, sfStartTime As Date, sfStopTime As Date, sfAliveTime As Date, sfCycle_Upload As Double, sfCycle_Download As Double, sfCycle_Mins As Double, Acct_Session_ID As String, ptRecID As Long, VirtualID As Long, PrimaryDNS As String, SecondaryDNS As String, DateCreated As Date, Optional sKey As String) As cls_subRadius
    'create a new object
    Dim objNewMember As cls_subRadius
    Set objNewMember = New cls_subRadius


    'set the properties passed into the method
    objNewMember.Key = Key
    objNewMember.RecID = RecID
    
        objNewMember.Username = Username
    
    objNewMember.Password = Password
    objNewMember.SessionsAllowed = SessionsAllowed
    objNewMember.AutoActivateFlag = AutoActivateFlag
    objNewMember.Activate = Activate
    objNewMember.Deactivate = Deactivate
    objNewMember.SessionTimeout = SessionTimeout
    objNewMember.IdleTimeout = IdleTimeout
    objNewMember.acci_RecID = acci_RecID
    objNewMember.ServiceType = ServiceType
    objNewMember.FramedProtocol = FramedProtocol
    objNewMember.Checked = Checked
    objNewMember.sfStartTime = sfStartTime
    objNewMember.sfStopTime = sfStopTime
    objNewMember.sfAliveTime = sfAliveTime
    objNewMember.sfCycle_Upload = sfCycle_Upload
    objNewMember.sfCycle_Download = sfCycle_Download
    objNewMember.sfCycle_Mins = sfCycle_Mins
    objNewMember.Acct_Session_ID = Acct_Session_ID
    objNewMember.ptRecID = ptRecID
    objNewMember.VirtualID = VirtualID
    objNewMember.PrimaryDNS = PrimaryDNS
    objNewMember.SecondaryDNS = SecondaryDNS
    objNewMember.DateCreated = DateCreated
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

Public Property Get Item(vntIndexKey As Variant) As cls_subRadius
Attribute Item.VB_UserMemId = 0
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
  'create the mcls_subRadius object when the col_subRadius class is created
  Set mvarcls_subRadius = New cls_subRadius
End Sub


Private Sub Class_Terminate()
  Set mvarcls_subRadius = Nothing
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

