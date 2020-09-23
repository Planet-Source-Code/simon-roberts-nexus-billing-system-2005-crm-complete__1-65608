VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "col_subEditLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"cls_subEditLog"
Attribute VB_Ext_KEY = "Collection" ,"UserDefinedCollection"
'local variable to hold collection
Private mCol As Collection
Public Function Commit(ByVal oConn As ADODB.Connection) As Long

Dim bUpdate As Boolean
    Dim SQL As String
    
    If oConn.State = adStateOpen Then
        If Me.Count > 0 Then
            Dim oAls As cls_subEditLog
            For Each oAls In Me
                bUpdate = False
                If Left(oAls.Key, 3) = "NEW" Then
                    Commit = Commit + 1: bUpdate = True
                    oAls.Key = "RECID" & MySQL.GetTMPRecID("acci_editlog", oConn, "RecID")
                ElseIf Left(oAls.Key, 5) = "UPDAT" Then
                    Commit = Commit + 1: bUpdate = True
                End If
                
                If bUpdate = True Then
                    SQL = "update acci_editlog set"
                    SQL = SQL + " acci_RecID = '" & oAls.acci_RecID
                    SQL = SQL + "', DateEditMade = '" & oAls.DateEditMade
                    SQL = SQL + "', EditTxt = '" & MySQL.ESC(oAls.EditTxt)
                    SQL = SQL + "', IPAddress = '" & oAls.IPAddress
                    SQL = SQL + "', SysopID = '" & oAls.SysopID
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
    Dim rsSysops As ADODB.Recordset
    Dim lx As Long
    Dim bFound As Boolean
    

     
    If MySQL.OpenTable(oConn, rsload, , "select distinct SysopID, EditTxt from acci_editlog where acci_RecID = '" & ClientID & "'", adOpenStatic, adLockReadOnly) = True Then
        If rsload.State = adStateOpen Then
            On Error Resume Next
            
            If rsload.RecordCount > 0 Then
                Call MySQL.OpenTable(oConn, rsSysops, , "select RecID, Username from Sysops")
                    
                For lx = Me.Count To 1 Step -1
                    Me.Remove lx
                Next

                While Not rsload.EOF Or Err.Number = 0
                    bFound = False
                    
                        With Me.Add("RECID" & rsload!RecID, IIf(IsNull(rsload!acci_RecID), ClientID, rsload!acci_RecID), IIf(IsNull(rsload!DateEditMade), sysnow, rsload!DateEditMade), IIf(IsNull(rsload!SysopID), 0, rsload!SysopID), IIf(IsNull(rsload!EditTxt), "", rsload!EditTxt), _
                                IIf(IsNull(rsload!IPAddress), "", rsload!IPAddress), IIf(IsNull(rsload!RecID), 0, rsload!RecID), "RECID" & rsload!RecID)
                            rsSysops.Filter = "RecID = " & rsload!SysopID
                            If rsSysops.RecordCount > 0 Then
                                .Username = IIf(IsNull(rsSysops!Username), "Unknown User", rsSysops!Username)
                            Else
                                .Username = "Unknown User"
                            End If
                        End With
                    rsload.MoveNext
                Wend
                GetClient = bResult
            End If
        End If
    End If
End Function


Public Function Add(Key As String, acci_RecID As Long, DateEditMade As Date, SysopID As Long, EditTxt As String, IPAddress As String, RecID As Long, Optional sKey As String) As cls_subEditLog
    
    On Error Resume Next
    
    'create a new object
    Dim objNewMember As cls_subEditLog
    Set objNewMember = New cls_subEditLog


    'set the properties passed into the method
    objNewMember.Key = Key
    objNewMember.acci_RecID = acci_RecID
    objNewMember.DateEditMade = DateEditMade
    objNewMember.SysopID = SysopID
    objNewMember.EditTxt = EditTxt
    objNewMember.IPAddress = IPAddress
    objNewMember.RecID = RecID
    objNewMember.IDX = Me.Count + 1
    
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    If Err.Number <> 0 Then
        Set Add = mCol(objNewMember.Key)
    Else
        'return the object created
        Set Add = objNewMember
    End If
    
    Set objNewMember = Nothing
    
    
End Function

Public Property Get Item(vntIndexKey As Variant) As cls_subEditLog
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

