VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "col_subDSLLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"cls_subDSLLink"
Attribute VB_Ext_KEY = "Collection" ,"UserDefinedCollection"
'local variable to hold collection
Private mCol As Collection
Public Function Commit(ByVal oConn As ADODB.Connection) As Long

    Dim bUpdate As Boolean
    Dim SQL As String
    
    If oConn.State = adStateOpen Then
        If Me.Count > 0 Then
            Dim oAls As cls_subDSLLink
            For Each oAls In Me
                bUpdate = False
                If Left(oAls.Key, 3) = "NEW" Then
                    Commit = Commit + 1: bUpdate = True
                    oAls.Key = "RECID" & MySQL.GetTMPRecID("acci_dslconnections", oConn, "RecID")
                ElseIf Left(oAls.Key, 5) = "UPDAT" Then
                    Commit = Commit + 1: bUpdate = True
                End If
                
                If bUpdate = True Then
                    SQL = "update acci_dslconnections set"
                    SQL = SQL + " acci_RecID = '" & oAls.acci_RecID
                    SQL = SQL + "', AccountActive = '" & oAls.AccountActive
                    SQL = SQL + "', AccountName = '" & MySQL.ESC(oAls.AccountName)
                    SQL = SQL + "', AreaCode = '" & oAls.AreaCode
                    SQL = SQL + "', cFlag = '" & oAls.cFlag
                    SQL = SQL + "', Checked = '" & oAls.Checked
                    SQL = SQL + "', Churn = '" & oAls.Churn
                    SQL = SQL + "', Country = '" & MySQL.ESC(oAls.Country)
                    SQL = SQL + "', eMail = '" & MySQL.ESC(oAls.eMail)
                    SQL = SQL + "', PhoneNumber = '" & MySQL.ESC(oAls.PhoneNumber)
                    SQL = SQL + "', Postcode = '" & oAls.PostCode
                    SQL = SQL + "', RadiusID = '" & oAls.RadiusID
                    SQL = SQL + "', State = '" & MySQL.ESC(oAls.State)
                    SQL = SQL + "', StreetName = '" & MySQL.ESC(oAls.StreetName)
                    SQL = SQL + "', StreetNo = '" & MySQL.ESC(oAls.StreetNo)
                    SQL = SQL + "', StreetType = '" & MySQL.ESC(oAls.StreetType)
                    SQL = SQL + "', Suburb = '" & MySQL.ESC(oAls.Suburb)
                    SQL = SQL + "', UnitNo = '" & oAls.UnitNo
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
    
    If MySQL.OpenTable(oConn, rsload, , "select * from acci_dslconnections where acci_RecID = '" & ClientID & "'") = True Then
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
                        Call Me.Add("RECID" & rsload!RecID, IIf(IsNull(rsload!AccountName), "", rsload!AccountName), IIf(IsNull(rsload!AreaCode), "", rsload!AreaCode), IIf(IsNull(rsload!PhoneNumber), "", rsload!PhoneNumber), IIf(IsNull(rsload!eMail), "", rsload!eMail), _
                                IIf(IsNull(rsload!acci_RecID), -1, rsload!acci_RecID), IIf(IsNull(rsload!RecID), -1, rsload!RecID), IIf(IsNull(rsload!Checked), True, rsload!Checked), IIf(IsNull(rsload!UnitNo), "", rsload!UnitNo), IIf(IsNull(rsload!StreetNo), "", rsload!StreetNo), _
                                IIf(IsNull(rsload!StreetName), "", rsload!StreetName), IIf(IsNull(rsload!StreetType), "", rsload!StreetType), IIf(IsNull(rsload!Suburb), "", rsload!Suburb), IIf(IsNull(rsload!Country), "", rsload!Country), IIf(IsNull(rsload!PostCode), "", rsload!PostCode), _
                                IIf(IsNull(rsload!State), "", rsload!State), IIf(IsNull(rsload!Created), sysnow, rsload!Created), IIf(IsNull(rsload!Churn), 1, rsload!Churn), IIf(IsNull(rsload!RadiusID), -1, rsload!RadiusID), IIf(IsNull(rsload!AccountActive), 0, rsload!AccountActive), IIf(IsNull(rsload!cFlag), 0, rsload!cFlag), "RECID" & rsload!RecID)
                                
                    End If
                    rsload.MoveNext
                Wend
                GetClient = bResult
            End If
        End If
    End If
End Function


Public Function Add(Key As String, AccountName As String, AreaCode As String, PhoneNumber As String, eMail As String, acci_RecID As Long, RecID As Long, Checked As Integer, UnitNo As String, StreetNo As String, StreetName As String, StreetType As String, Suburb As String, Country As String, PostCode As String, State As String, Created As Date, Churn As Byte, RadiusID As Long, AccountActive As Integer, cFlag As Byte, Optional sKey As String) As cls_subDSLLink
    'create a new object
    Dim objNewMember As cls_subDSLLink
    Set objNewMember = New cls_subDSLLink


    'set the properties passed into the method
    objNewMember.Key = Key
    objNewMember.AccountName = AccountName
    objNewMember.AreaCode = AreaCode
    objNewMember.PhoneNumber = PhoneNumber
    objNewMember.eMail = eMail
    objNewMember.acci_RecID = acci_RecID
    objNewMember.RecID = RecID
    objNewMember.Checked = Checked
    objNewMember.UnitNo = UnitNo
    objNewMember.StreetNo = StreetNo
    objNewMember.StreetName = StreetName
    objNewMember.StreetType = StreetType
    objNewMember.Suburb = Suburb
    objNewMember.Country = Country
    objNewMember.PostCode = PostCode
    objNewMember.State = State
    objNewMember.Created = Created
    objNewMember.Churn = Churn
    objNewMember.RadiusID = RadiusID
    objNewMember.AccountActive = AccountActive
    objNewMember.cFlag = cFlag
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

Public Property Get Item(vntIndexKey As Variant) As cls_subDSLLink
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

