VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "col_subHardware"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"cls_subHardware"
Attribute VB_Ext_KEY = "Collection" ,"UserDefinedCollection"
'local variable to hold collection
Private mCol As Collection
Public Function Commit(ByVal oConn As ADODB.Connection) As Long
End Function

Public Function GetClient(ByVal ClientID As Long, ByVal oConn As ADODB.Connection) As Byte

   Const bNoReturn = 0
    Const bResult = 1
    
    Dim rsload As ADODB.Recordset
    Dim lx As Long
    Dim bFound As Boolean
    
    If MySQL.OpenTable(oConn, rsload, , "select * from acci_hardware where acci_RecID = '" & ClientID & "'") = True Then
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
                        Call Me.Add("RECID" & rsload!RecID, IIf(IsNull(rsload!RecID), 0, rsload!RecID), IIf(IsNull(rsload!Modem), "", rsload!Modem), IIf(IsNull(rsload!Processor), "", rsload!Processor), IIf(IsNull(rsload!VideoCard), "", rsload!VideoCard), IIf(IsNull(rsload!Monitor), "", rsload!Monitor), _
                                IIf(IsNull(rsload!PCType), "", rsload!PCType), IIf(IsNull(rsload!NetworkCard), "", rsload!NetworkCard), IIf(IsNull(rsload!OS), "", rsload!OS), IIf(IsNull(rsload!Printer), "", rsload!Printer), IIf(IsNull(rsload!Mainboard), "", rsload!Mainboard), _
                                IIf(IsNull(rsload!acci_RecID), 0, rsload!acci_RecID), IIf(IsNull(rsload!ExtraXML), "", rsload!ExtraXML), IIf(IsNull(rsload!acciServiceID), 0, rsload!acciServiceID), "RECID" & rsload!RecID)
                                
                    End If
                    rsload.MoveNext
                Wend
                GetClient = bResult
            End If
        End If
    End If
End Function


Public Function Add(Key As String, RecID As Long, Modem As String, Processor As String, VideoCard As String, Monitor As String, PCType As String, NetworkCard As String, OS As String, Printer As String, Mainboard As String, acci_RecID As Long, ExtraXML As String, acciServicesID As Long, Optional sKey As String) As cls_subHardware
    'create a new object
    Dim objNewMember As cls_subHardware
    Set objNewMember = New cls_subHardware


    'set the properties passed into the method
    objNewMember.Key = Key
    objNewMember.RecID = RecID
    objNewMember.Modem = Modem
    objNewMember.Processor = Processor
    objNewMember.VideoCard = VideoCard
    objNewMember.Monitor = Monitor
    objNewMember.PCType = PCType
    objNewMember.NetworkCard = NetworkCard
    objNewMember.OS = OS
    objNewMember.Printer = Printer
    objNewMember.Mainboard = Mainboard
    objNewMember.acci_RecID = acci_RecID
    objNewMember.ExtraXML = ExtraXML
    objNewMember.acciServiceID = acciServicesID
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

Public Property Get Item(vntIndexKey As Variant) As cls_subHardware
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

