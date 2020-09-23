VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "col_subPhoneNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"col_subServices"
Attribute VB_Ext_KEY = "Collection" ,"UserDefinedCollection"
Attribute VB_Ext_KEY = "Member1" ,"col_subServices"
'local variable to hold collection
Private mCol As Collection
Private mvarCipherSalt As String 'local copy
Private mvarcol_subServices As col_subServices
Public Function Commit(ByVal oConn As ADODB.Connection) As Long

    Dim bUpdate As Boolean
    Dim SQL As String
    
    If oConn.State = adStateOpen Then
        If Me.Count > 0 Then
            Dim oAls As cls_subPhoneNo
            For Each oAls In Me
                bUpdate = False
                If Left(oAls.Key, 3) = "NEW" Then
                    Commit = Commit + 1: bUpdate = True
                    oAls.Key = "RECID" & MySQL.GetTMPRecID("acci_phonenumbers", oConn, "RecID")
                ElseIf Left(oAls.Key, 5) = "UPDAT" Then
                    Commit = Commit + 1: bUpdate = True
                End If
                
                If bUpdate = True Then
                    SQL = "update acci_phonenumbers set"
                    SQL = SQL + " acci_RecID = '" & oAls.acci_RecID
                    SQL = SQL + "', Cancelled = '" & oAls.Cancelled
                    SQL = SQL + "', Checked = '" & oAls.Checked
                    SQL = SQL + "', ContactName = '" & MySQL.ESC(oAls.ContactName)
                    SQL = SQL + "', Extension = '" & MySQL.ESC(oAls.Extension)
                    SQL = SQL + "', FlagID = '" & oAls.FlagID
                    SQL = SQL + "', PhoneNumber = '" & MySQL.ESC(oAls.PhoneNumber)
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
    
    If MySQL.OpenTable(oConn, rsload, , "select * from acci_phonenumbers where acci_RecID = '" & ClientID & "'") = True Then
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
                        Call Me.Add("RECID" & rsload!RecID, rsload!RecID, IIf(IsNull(rsload!acci_RecID), ClientID, rsload!acci_RecID), IIf(IsNull(rsload!FlagID), 0, rsload!FlagID), IIf(IsNull(rsload!DateAdded), sysnow, rsload!DateAdded), IIf(IsNull(rsload!PhoneNumber), "", rsload!PhoneNumber), _
                                IIf(IsNull(rsload!Extension), "", rsload!Extension), IIf(IsNull(rsload!ContactName), "", rsload!ContactName), IIf(IsNull(rsload!Cancelled), "", Val(rsload!Cancelled)), IIf(IsNull(rsload!Checked), "", Val(rsload!Checked)), IIf(IsNull(rsload!ShortNote), "", rsload!ShortNote), "RECID" & rsload!RecID)
                                
                    End If
                    rsload.MoveNext
                Wend
                GetClient = bResult
            End If
        End If
    End If
End Function



Public Property Get col_subServices() As col_subServices
    If mvarcol_subServices Is Nothing Then
        Set mvarcol_subServices = New col_subServices
    End If


    Set col_subServices = mvarcol_subServices
End Property


Public Property Set col_subServices(vData As col_subServices)
    Set mvarcol_subServices = vData
End Property

Public Property Let CipherSalt(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CipherSalt = 5
    mvarCipherSalt = vData
End Property


Public Property Get CipherSalt() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CipherSalt
    CipherSalt = mvarCipherSalt
End Property




Public Function Add(Key As String, RecID As Long, acci_RecID As Long, FlagID As Integer, DateAdded As Date, PhoneNumber As String, Extension As String, ContactName As String, Cancelled As Integer, Checked As Integer, ShortNote As String, Optional sKey As String) As cls_subPhoneNo
    'create a new object
    Dim objNewMember As cls_subPhoneNo
    Set objNewMember = New cls_subPhoneNo


    'set the properties passed into the method
    objNewMember.Key = Key
    objNewMember.RecID = RecID
    objNewMember.acci_RecID = acci_RecID
    objNewMember.FlagID = FlagID
    objNewMember.DateAdded = DateAdded
    objNewMember.PhoneNumber = PhoneNumber
    objNewMember.Extension = Extension
    objNewMember.ContactName = ContactName
    objNewMember.Cancelled = Cancelled
    objNewMember.Checked = Checked
    objNewMember.ShortNote = ShortNote
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

Public Property Get Item(vntIndexKey As Variant) As cls_subPhoneNo
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
  Set mvarcol_subServices = Nothing
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

