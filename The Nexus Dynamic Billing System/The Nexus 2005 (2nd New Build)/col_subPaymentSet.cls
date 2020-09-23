VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "col_subPaymentSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"clsPaymentSet"
Attribute VB_Ext_KEY = "Member0" ,"clsPaymentSet"
'local variable to hold collection
Private mCol As Collection

Public Function Add(Key As String, ddAccountName As String, ddBSB As String, ddAcountNo As String, ccCardName As String, ccCardNumber As String, ccCardExpiry As String, ccCIC As String, swWord As String, swNumber As String, Order As Long, RecID As Long, Optional sKey As String) As clsPaymentSet
    'create a new object
    Dim objNewMember As clsPaymentSet
    Set objNewMember = New clsPaymentSet


    'set the properties passed into the method
    objNewMember.Key = Key
    objNewMember.ddAccountName = ddAccountName
    objNewMember.ddBSB = ddBSB
    objNewMember.ddAcountNo = ddAcountNo
    objNewMember.ccCardName = ccCardName
    objNewMember.ccCardNumber = ccCardNumber
    objNewMember.ccCardExpiry = ccCardExpiry
    objNewMember.ccCIC = ccCIC
    objNewMember.swWord = swWord
    objNewMember.swNumber = swNumber
    objNewMember.IDX = Me.Count + 1
    objNewMember.Order = Order
    objNewMember.RecID = RecID
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As clsPaymentSet
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
            Dim oAls As clsPaymentSet
            For Each oAls In Me
                bUpdate = False
                If Left(oAls.Key, 3) = "NEW" Then
                    Commit = Commit + 1: bUpdate = True
                    oAls.RecID = MySQL.GetTMPRecID("acci_paymentsettings", oConn, "RecID")
                ElseIf Left(oAls.Key, 5) = "UPDAT" Then
                    Commit = Commit + 1: bUpdate = True
                End If
                
                If bUpdate = True Then
                    SQL = "update acci_paymentsettings set"
                    SQL = SQL + " acci_RecID = '" & oAls.acci_RecID
                    If Len(oAls.ddAccountName) > 0 Then
                        SQL = SQL + "', ddAccountName = AES_ENCRYPT('" & oAls.ddAccountName & "','" & odb.colSalts.ReturnSalt(CCSalt) & "')"
                        SQL = SQL + ", ddBSB = encode('" & oAls.ddBSB & "','" & odb.colSalts.ReturnSalt(CCSalt) & "')"
                        SQL = SQL + ", ddAccountNumber = AES_ENCRYPT('" & oAls.ddAcountNo & "','" & odb.colSalts.ReturnSalt(CCSalt) & "')"
                    End If
                    If Len(oAls.ccCardNumber) > 9 Then
                        SQL = SQL + ", ccCardName = AES_ENCRYPT('" & oAls.ccCardName & "','" & odb.colSalts.ReturnSalt(CCSalt) & "')"
                        SQL = SQL + ", ccCardNumber = AES_ENCRYPT('" & oAls.ccCardNumber & "','" & odb.colSalts.ReturnSalt(CCSalt) & odb.colSalts.ReturnSalt(PWSalt) & "')"
                        SQL = SQL + ", ccExpiryDate = '" & oAls.ccCardExpiry
                        SQL = SQL + "', ccCIC = '" & oAls.ccCIC & "'"
                    End If
                    SQL = SQL + ", swWord = '" & oAls.swWord
                    SQL = SQL + "', swNumber = '" & oAls.swNumber
                    SQL = SQL + "', cOrder = '" & oAls.Order
                    SQL = SQL + "', Checked = '" & oAls.Checked
                    SQL = SQL + "' WHERE  RecID = '" & oAls.RecID & "'"
                    
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
    Dim SQL As String
    Dim bFound As Boolean
    
    SQL = "Select "
    SQL = SQL + "AES_DECRYPT(ddAccountName,'" & odb.colSalts.ReturnSalt(CCSalt) & "') as ddAccountName"
    SQL = SQL + ", decode(ddBSB,'" & odb.colSalts.ReturnSalt(PWSalt) & "') as ddBSB"
    SQL = SQL + ", AES_DECRYPT(ddAccountNumber,'" & odb.colSalts.ReturnSalt(CCSalt) & "') as ddAccountNumber"
    SQL = SQL + ", AES_DECRYPT(ccCardName,'" & odb.colSalts.ReturnSalt(CCSalt) & "') as ccCardName"
    SQL = SQL + ", RIGHT(AES_DECRYPT(ccCardNumber ,'" & odb.colSalts.ReturnSalt(CCSalt) & odb.colSalts.ReturnSalt(PWSalt) & "'),5) as ccCardNumber"
    SQL = SQL + ", ccExpiryDate"
    SQL = SQL + ", ccCIC"
    SQL = SQL + ", swWord"
    SQL = SQL + ", swNumber"
    SQL = SQL + ", cOrder"
    SQL = SQL + ", acci_RecID"
    SQL = SQL + ", Checked"
    SQL = SQL + ", RecID from acci_paymentsettings Where acci_RecID = '" & ClientID & "' order by cOrder"
    
    
    If MySQL.OpenTable(oConn, rsload, , SQL) = True Then
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
                        With Me.Add("RECID" & rsload!RecID, IIf(IsNull(rsload!ddAccountName), "", rsload!ddAccountName), IIf(IsNull(rsload!ddBSB), "", rsload!ddBSB), IIf(IsNull(rsload!ddAccountnumber), "", rsload!ddAccountnumber), _
                            IIf(IsNull(rsload!ccCardName), "", rsload!ccCardName), IIf(IsNull(rsload!ccCardNumber), "", rsload!ccCardNumber), IIf(IsNull(rsload!ccExpiryDate), "", rsload!ccExpiryDate), IIf(IsNull(rsload!ccCIC), "", rsload!ccCIC), _
                            IIf(IsNull(rsload!swWord), "", rsload!swWord), IIf(IsNull(rsload!swNumber), "", rsload!swNumber), IIf(IsNull(rsload!cOrder), "", rsload!cOrder), IIf(IsNull(rsload!RecID), "", rsload!RecID), "RECID" & rsload!RecID)
                            .Checked = IIf(IsNull(rsload!Checked), "", rsload!Checked)
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

