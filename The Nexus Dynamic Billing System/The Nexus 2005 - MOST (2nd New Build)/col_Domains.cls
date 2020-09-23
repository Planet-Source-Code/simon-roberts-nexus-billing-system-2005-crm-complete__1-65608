VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "col_Domains"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"cls_Domains"
Attribute VB_Ext_KEY = "Collection" ,"UserDefinedCollection"
'local variable to hold collection
Private mCol As Collection
Public Function Commit(ByVal oConn As ADODB.Connection) As Long


    If oConn.State = adStateOpen Then
        If Me.Count > 0 Then
            Dim oAls As cls_Domains
            For Each oAls In Me
                bUpdate = False
                If Left(oAls.Key, 3) = "NEW" Then
                    Commit = Commit + 1: bUpdate = True
                    oAls.Key = "RECID" & MySQL.GetTMPRecID("domainlist", oConn, "RecID")
                ElseIf Left(oAls.Key, 5) = "UPDAT" Then
                    Commit = Commit + 1: bUpdate = True
                End If
                
                If bUpdate = True Then
                    SQL = "update domainlist set"
                    SQL = SQL + " acci_RecID = '" & oAls.acci_RecID
                    SQL = SQL + "', AdminEmail = '" & MySQL.ESC(oAls.AdminEmail)
                    SQL = SQL + "', Checked = '" & oAls.Checked
                    SQL = SQL + "', ContactName = '" & MySQL.ESC(oAls.ContactName)
                    SQL = SQL + "', Domain = '" & MySQL.ESC(oAls.Domain)
                    SQL = SQL + "', Status = '" & MySQL.ESC(oAls.Status)
                    SQL = SQL + "', SubDomainLimit = '" & oAls.SubDomainLimit
                    SQL = SQL + "', SysopID = '" & oAls.SysopID
                    SQL = SQL + "', TechName = '" & MySQL.ESC(oAls.TechName)
                    SQL = SQL + "', TechPass = AES_ENCRYPT('" & MySQL.ESC(oAls.TechPass) & "','" & odb.colSalts(PWSalt) & "')"
                    SQL = SQL + "', VirtualID = '" & oAls.VirtualID
                    SQL = SQL + "', vKey = AES_ENCRYPT('" & MySQL.ESC(oAls.vKey) & "','" & odb.colSalts(CCSalt) & "')"
                    SQL = SQL + " WHERE  RecID = '" & Mid(oAls.Key, 6) & "'"
                    
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
    
    If MySQL.OpenTable(oConn, rsload, , "select *, AES_DECRYPT(vKey,'" + odb.colSalts.ReturnSalt(CCSalt) + "') as vKeyDesc, AES_DECRYPT(TechPass,'" + odb.colSalts.ReturnSalt(PWSalt) + "') as TechPassDesc from domainlist where acci_RecID = '" & ClientID & "'") = True Then
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
                        With Me.Add("RECID" & rsload!RecID, rsload!RecID, IIf(IsNull(rsload!AdminEmail), "", rsload!AdminEmail), IIf(IsNull(rsload!AdminEmail), "", rsload!AdminEmail), IIf(IsNull(rsload!Status), "Z", rsload!Status), _
                                    IIf(IsNull(rsload!Checked), "", rsload!Checked), IIf(IsNull(rsload!SubDomainLimit), 1, rsload!SubDomainLimit), IIf(IsNull(rsload!vKey), "", rsload!vKeyDesc), _
                                    IIf(IsNull(rsload!acci_RecID), "", rsload!acci_RecID), IIf(IsNull(rsload!TechName), "", rsload!TechName), IIf(IsNull(rsload!TechPassDesc), "", rsload!TechPassDesc), IIf(IsNull(rsload!SysopID), 0, rsload!SysopID), _
                                    IIf(IsNull(rsload!VirtualID), 0, rsload!VirtualID), "RECID" & rsload!RecID)
                            
                            .ContactName = IIf(IsNull(rsload!ContactName), "", rsload!ContactName)
                        End With
                    End If
                    rsload.MoveNext
                Wend
                GetClient = bResult
            End If
        End If
    End If
    
End Function


Public Function Add(Key As String, RecID As Long, Domain As String, AdminEmail As String, Status As String, Checked As Integer, SubDomainLimit As Integer, vKey As String, acci_RecID As Long, TechName As String, TechPass As String, SysopID As Long, VirtualID As Long, Optional sKey As String) As cls_Domains
    'create a new object
    Dim objNewMember As cls_Domains
    Set objNewMember = New cls_Domains


    'set the properties passed into the method
    objNewMember.Key = Key
    objNewMember.RecID = RecID
    objNewMember.Domain = Domain
    objNewMember.AdminEmail = AdminEmail
    objNewMember.Status = Status
    objNewMember.Checked = Checked
    objNewMember.SubDomainLimit = SubDomainLimit
    objNewMember.vKey = vKey
    objNewMember.acci_RecID = acci_RecID
    objNewMember.TechName = TechName
    objNewMember.TechPass = TechPass
    objNewMember.SysopID = SysopID
    objNewMember.VirtualID = VirtualID
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

Public Property Get Item(vntIndexKey As Variant) As cls_Domains
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
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

