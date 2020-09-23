VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colAuditTrail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"clsAuditTrail"
Attribute VB_Ext_KEY = "Member0" ,"clsAuditTrail"
'local variable to hold collection
Private mCol As Collection

Public Function Add(Key As String, systemstamp As Date, sysnow As Date, localtime As Date, appname As String, appversion As String, apphdc As Long, formname As String, formhwnd As Long, vartype As String, varname As String, oldcur As Single, newcur As Single, oldvalue As String, newvalue As String, oldpointer As Double, newpointer As Double, IDX As Long, SysopID As Long, VirtualID As Long, AgencyID As Long, acci_RecID As Long, RefundID As Long, invtrxrid As Long, FlagID As Long, Description As String, Checked As Integer, Optional sKey As String) As clsAuditTrail
    'create a new object
    Dim objNewMember As clsAuditTrail
    Set objNewMember = New clsAuditTrail


    'set the properties passed into the method
    objNewMember.Key = Key
    objNewMember.systemstamp = systemstamp
    objNewMember.sysnow = sysnow
    objNewMember.localtime = localtime
    objNewMember.appname = appname
    objNewMember.appversion = appversion
    objNewMember.apphdc = apphdc
    objNewMember.formname = formname
    objNewMember.formhwnd = formhwnd
    objNewMember.vartype = vartype
    objNewMember.varname = varname
    objNewMember.oldcur = oldcur
    objNewMember.newcur = newcur
    objNewMember.oldvalue = oldvalue
    objNewMember.newvalue = newvalue
    objNewMember.oldpointer = oldpointer
    objNewMember.newpointer = newpointer
    objNewMember.IDX = IDX
    objNewMember.SysopID = SysopID
    objNewMember.VirtualID = VirtualID
    objNewMember.AgencyID = AgencyID
    objNewMember.acci_RecID = acci_RecID
    objNewMember.RefundID = RefundID
    objNewMember.invtrxrid = invtrxrid
    objNewMember.FlagID = FlagID
    objNewMember.Description = Description
    objNewMember.Checked = Checked
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As clsAuditTrail
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

