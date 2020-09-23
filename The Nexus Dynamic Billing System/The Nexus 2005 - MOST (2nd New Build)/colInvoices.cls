VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colInvoices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"clsInvoice"
Attribute VB_Ext_KEY = "Member0" ,"clsInvoice"
'local variable to hold collection
Private mCol As Collection
Public Function Clear()

    Dim ix As Long
    
    If Me.Count > 0 Then
        For ix = Me.Count To 1 Step -1
            Me.Remove ix
        Next
    End If
    
End Function

Public Function TotalPaid(Checked As Boolean) As Currency

    If Me.Count > 0 Then
        For ix = Me.Count To 1 Step -1
            If Me(ix).Checked = Checked Then
                TotalPaid = TotalPaid + Me(ix).Paid
            End If
        Next
    End If
    
End Function

Public Function TotalCredit(Checked As Boolean) As Currency

    If Me.Count > 0 Then
        For ix = Me.Count To 1 Step -1
            If Me(ix).Checked = Checked Then
                TotalCredit = TotalCredit + Me(ix).Credit
            End If
        Next
    End If
    
End Function

Public Function TotalDebit(Checked As Boolean) As Currency


    If Me.Count > 0 Then
        For ix = Me.Count To 1 Step -1
            If Me(ix).Checked = Checked Then
                TotalDebit = TotalDebit + Me(ix).Debit
            End If
        Next
    End If
    
End Function


Public Function Add(Key As String, VirtualID As Long, acciRecID As Long, InvRecID As Long, Description As String, Debit As Currency, Credit As Currency, Paid As Currency, Checked As Boolean, Optional sKey As String) As clsInvoice
    'create a new object
    Dim objNewMember As clsInvoice
    Set objNewMember = New clsInvoice


    'set the properties passed into the method
    objNewMember.Key = Key
    objNewMember.VirtualID = VirtualID
    objNewMember.acciRecID = acciRecID
    objNewMember.InvRecID = InvRecID
    objNewMember.Description = Description
    objNewMember.Debit = Debit
    objNewMember.Credit = Credit
    objNewMember.Paid = Paid
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

Public Property Get Item(vntIndexKey As Variant) As clsInvoice
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrueving the number of elements in the
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

