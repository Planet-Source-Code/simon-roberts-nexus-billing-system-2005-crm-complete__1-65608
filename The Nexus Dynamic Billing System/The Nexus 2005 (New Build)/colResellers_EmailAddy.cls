VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colResellers_EmailAddy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"clsResellers_EmailAddy"
Attribute VB_Ext_KEY = "Member0" ,"clsResellers_EmailAddy"
Enum enumPrivateFetchStatusd
    NoChanges = 1
    Edited_Update = 2
    NewLine_Insert = 3
    DeleteRecord = 4
    LoadingData = 5
    Saving = 6
    Deleting = 7
    Idle = 8
    CreateNewViSP = 9
    AccountHeader = 10
    Addresses = 11
    EmailAddresses = 12
    PhoneNumbers = 13
    InvoiceItems = 14
    PaymentHistory = 15
    LoadHeader = 16
    LoadInvoice = 17
    LoadAllContactDetails = 18
    LoadPaymentHistory = 19
    LoadAll = 20
    LoadMinimum = 21
    LoadEmail = 22
    LoadPhone = 23
    LoadAddress = 24
End Enum


'local variable to hold collection
Private mCol As Collection
Public Function Clear(Optional SESSIONKey As String = "")

    Dim ix As Long
    
    If Me.Count > 0 Then
        For ix = Me.Count To 1 Step -1
            If Len(SESSION) > 0 Then
                If Me(ix).SESSION = SESSIONKey Then
                    Me.Remove ix
                End If
            Else
                Me.Remove ix
            End If
        Next
    End If
    
End Function
Public Function Add(Key As String, RecID As Long, visp_RecID As Long, FlagID As Integer, DateAdded As Date, EmailAddress As String, ContactName As String, Cancelled As Integer, Checked As Integer, PhotoURL As String, FetchStatus As enumPrivateFetchStatusd, SESSION As String, Optional sKey As String) As clsResellers_EmailAddy
    
    
    If Me.IsKeyInstance(sKey) = -1 Then
    
       
        'create a new object
        Dim objNewMember As clsResellers_EmailAddy
        Set objNewMember = New clsResellers_EmailAddy
    
    
        objNewMember.Key = Key
        objNewMember.RecID = RecID
        objNewMember.visp_RecID = visp_RecID
        objNewMember.IDX = Me.Count + 1
        objNewMember.SESSION = SESSION
        objNewMember.FlagID = FlagID
        objNewMember.DateAdded = DateAdded
        objNewMember.EmailAddress = EmailAddress
        objNewMember.ContactName = ContactName
        objNewMember.Cancelled = Cancelled
        objNewMember.Checked = Checked
        objNewMember.PhotoURL = PhotoURL
        objNewMember.FetchStatus = Val(FetchStatus)
    
        
        If Len(sKey) = 0 Then
            mCol.Add objNewMember
        Else
            mCol.Add objNewMember, sKey
        End If
    
    
        'return the object created
        Set Add = objNewMember
        Set objNewMember = Nothing
    Else
        Set Add = Me(Me.IsKeyInstance(sKey))
    End If
       

End Function
Public Function IsKeyInstance(sKey As String) As Long

    Static LastsKey As String
    Static Intialised As Boolean
    Static LastIndex As Long
    
    IsKeyInstance = -1
    
    If Me.Count > 0 Then
        If Not sKey = LastsKey Or Not Initialised = True Then
            Dim lk As Long
            For lk = 1 To Me.Count
                If Me(lk).Key = sKey Then
                    LastIndex = lk
                    IsKeyInstance = lk
                    Initialised = True
                    Exit For
                End If
            Next
        Else
            IsKeyInstance = LastIndex
        End If
    Else
        IsKeyInstance = -1
    End If

End Function



Public Property Get Item(vntIndexKey As Variant) As clsResellers_EmailAddy
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

