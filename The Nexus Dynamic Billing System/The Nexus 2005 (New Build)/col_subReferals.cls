VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "col_subReferals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"cls_subReferals"
Attribute VB_Ext_KEY = "Collection" ,"UserDefinedCollection"
'local variable to hold collection
Private mCol As Collection

Public Function Commit(ByVal oConn As ADODB.Connection) As Long

    Dim bUpdate As Boolean
    Dim SQL As String
    
    If oConn.State = adStateOpen Then
        If Me.Count > 0 Then
            Dim oAls As cls_subReferals
            For Each oAls In Me
                bUpdate = False
                If Left(oAls.Key, 3) = "NEW" Then
                    Commit = Commit + 1: bUpdate = True
                    oAls.Key = "RECID" & MySQL.GetTMPRecID("acci_referedby", oConn, "RecID")
                ElseIf Left(oAls.Key, 5) = "UPDAT" Then
                    Commit = Commit + 1: bUpdate = True
                End If
                
                If bUpdate = True Then
                    SQL = "update acci_referedby set"
                    SQL = SQL + " acci_RecID = '" & oAls.acci_RecID
                    SQL = SQL + "', acci_RecID2 = '" & oAls.acci_RecID2
                    SQL = SQL + "', acciServiceID = '" & oAls.acciServiceID
                    SQL = SQL + "', Cancelled = '" & oAls.Cancelled
                    SQL = SQL + "', Checked = '" & oAls.Checked
                    SQL = SQL + "', ContactName = '" & MySQL.ESC(oAls.ContactName)
                    SQL = SQL + "', FlagID = '" & oAls.FlagID
                    SQL = SQL + "', ShortNote = '" & MySQL.ESC(oAls.ShortNote)
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
    
    If MySQL.OpenTable(oConn, rsload, , "select accountinfo.AccountName, acci_referedby.* from acci_referedby, accountinfo Where accountinfo.RecID = acci_referedby.acci_RecID2 AND acci_referedby.AccI_RecID = '" & ClientID & "'") = True Then
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
                    
                        With Me.Add("RECID" & rsload!RecID, IIf(IsNull(rsload!acci_RecID), ClientID, rsload!acci_RecID), rsload!RecID, IIf(IsNull(rsload!acci_RecID2), 0, rsload!acci_RecID2), IIf(IsNull(rsload!FlagID), 0, rsload!FlagID), IIf(IsNull(rsload!DateAdded), sysnow, rsload!DateAdded), _
                                IIf(IsNull(rsload!ContactName), "", rsload!ContactName), IIf(IsNull(rsload!ShortNote), "", rsload!ShortNote), IIf(IsNull(rsload!Cancelled), "", Val(rsload!Cancelled)), IIf(IsNull(rsload!Checked), "", Val(rsload!Checked)), IIf(IsNull(rsload!acciServiceID), 0, rsload!acciServiceID), "RECID" & rsload!RecID)
                            .AccountName = IIf(IsNull(rsload!AccountName), "::", rsload!AccountName)
                        End With
                    End If
                    rsload.MoveNext
                Wend
                GetClient = bResult
            End If
        End If
    End If

End Function


Public Function Add(Key As String, acci_RecID As Long, RecID As Long, acci_RecID2 As Long, FlagID As Integer, DateAdded As Date, ContactName As String, ShortNote As String, Cancelled As Integer, Checked As Integer, acciServiceID As Long, Optional sKey As String) As cls_subReferals
    'create a new object
    Dim objNewMember As cls_subReferals
    Set objNewMember = New cls_subReferals


    'set the properties passed into the method
    objNewMember.Key = Key
    objNewMember.acci_RecID = acci_RecID
    objNewMember.RecID = RecID
    objNewMember.acci_RecID2 = acci_RecID2
    objNewMember.FlagID = FlagID
    objNewMember.DateAdded = DateAdded
    objNewMember.ContactName = ContactName
    objNewMember.ShortNote = ShortNote
    objNewMember.Cancelled = Cancelled
    objNewMember.Checked = Checked
    objNewMember.acciServiceID = acciServiceID
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

Public Property Get Item(vntIndexKey As Variant) As cls_subReferals
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

