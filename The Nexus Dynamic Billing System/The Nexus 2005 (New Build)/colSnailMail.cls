VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colSnailMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"clsSnailMail"
Attribute VB_Ext_KEY = "Collection" ,"UserDefinedCollection"
'local variable to hold collection
Private mCol As Collection
Public Function Commit(ByVal oConn As ADODB.Connection) As Long

    Dim bUpdate As Boolean
    Dim SQL As String
    
    If oConn.State = adStateOpen Then
        If Me.Count > 0 Then
            Dim oAls As clsSnailMail
            For Each oAls In Me
                bUpdate = False
                If Left(oAls.Key, 3) = "NEW" Then
                    Commit = Commit + 1: bUpdate = True
                    oAls.Key = "RECID" & MySQL.GetTMPRecID("acci_addresses", oConn, "RecID")
                ElseIf Left(oAls.Key, 5) = "UPDAT" Then
                    Commit = Commit + 1: bUpdate = True
                End If
                
                If bUpdate = True Then
                    SQL = "update acci_addresses set"
                    SQL = SQL + " acci_RecID = '" & oAls.acci_RecID
                    SQL = SQL + "', Cancelled = '" & oAls.Cancelled
                    SQL = SQL + "', Checked = '" & oAls.Checked
                    SQL = SQL + "', ContactName = '" & MySQL.ESC(oAls.ContactName)
                    SQL = SQL + "', Country = '" & MySQL.ESC(oAls.Country)
                    SQL = SQL + "', FlagID = '" & oAls.FlagID
                    SQL = SQL + "', Postcode = '" & MySQL.ESC(oAls.PostCode)
                    SQL = SQL + "', State = '" & MySQL.ESC(oAls.State)
                    SQL = SQL + "', Street1 = '" & MySQL.ESC(oAls.Street1)
                    SQL = SQL + "', Street2 = '" & MySQL.ESC(oAls.Street2)
                    SQL = SQL + "', Suburb = '" & MySQL.ESC(oAls.Suburb)
                    SQL = SQL + "' WHERE  RecID = '" & Mid(oAls.Key, 6) & "'"
                    
                    MySQL.Execute oConn, CStr(SQL)
                    
                    oAls.Key = "RECID" & Mid(oAls.Key, 6)
                End If
                
            Next
        End If
    End If

End Function

Public Function GetClient(ByVal ClientID As Long, ByVal oConn As ADODB.Connection) As Byte
    
    If oConn.State = adStateOpen Then
    
        Dim rsGet As ADODB.Recordset
        Dim SQL As String
        
        SQL = "Select * from acci_addresses where AccI_RecID = '" & ClientID & "'"
        
        Call MySQL.OpenTable(oConn, rsGet, , SQL)
        
        If rsGet.State = adStateOpen Then
            While Not rsGet.EOF
                With Me.Add("RECID" & rsGet!RecID, ClientID, rsGet!FlagID, rsGet!DateCreated, rsGet!ContactName, rsGet!Street1, rsGet!Street2, _
                            rsGet!Country, rsGet!State, rsGet!PostCode, rsGet!Suburb, rsGet!Cancelled, rsGet!Checked, "RECID" & rsGet!RecID)
                End With
                rsGet.MoveNext
            Wend
        End If
        
    End If

End Function


Public Function Add(Key As String, acci_RecID As Long, FlagID As Byte, DateCreated As Date, ContactName As String, Street1 As String, Street2 As String, Country As String, State As String, PostCode As String, Suburb As String, Cancelled As Integer, Checked As Integer, Optional sKey As String) As clsSnailMail
    'create a new object
    Dim objNewMember As clsSnailMail
    Set objNewMember = New clsSnailMail


    'set the properties passed into the method
    objNewMember.Key = Key
    objNewMember.acci_RecID = acci_RecID
    objNewMember.FlagID = FlagID
    objNewMember.DateCreated = DateCreated
    objNewMember.ContactName = ContactName
    objNewMember.Street1 = Street1
    objNewMember.Street2 = Street2
    objNewMember.Country = Country
    objNewMember.State = State
    objNewMember.PostCode = PostCode
    objNewMember.Suburb = Suburb
    objNewMember.Cancelled = Cancelled
    objNewMember.Checked = Checked
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

Public Property Get Item(vntIndexKey As Variant) As clsSnailMail
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

