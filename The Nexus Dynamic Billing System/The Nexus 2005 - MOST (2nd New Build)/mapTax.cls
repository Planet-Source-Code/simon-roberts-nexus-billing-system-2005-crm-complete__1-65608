VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mapTax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"clsTaxExpress"
Attribute VB_Ext_KEY = "Member0" ,"clsTaxExpress"
'local variable to hold collection
Private mCol As Collection
Public Sub Clear(Optional TaxCode As String, Optional CountryCode As String, Optional lGroup As Long = -1)

    Dim Xnt As Long
    
    If lGroup <> -1 Or CountryCode <> "" Or TaxCode <> "" Then
        For Xnt = Me.Count To 1 Step -1
            If lGroup = Me(Xnt).lGroup Or CountryCode = Me(Xnt).Country Or TaxCode = Me(Xnt).Code Then
                Me.Remove Xnt
            End If
        Next
    Else
        For Xnt = Me.Count To 1 Step -1
            Me.Remove Xnt
        Next
    End If
    
End Sub

Public Function FindTaxRate(Optional Codes As String, Optional CountryCode As String, Optional AmountBeingTaxed As Currency = -1.01) As Single

    Dim iOverall As Single
    Dim iTaxes As Long
    
    Select Case AmountBeingTaxed
    Case -1.01
    
        If Me.Count > 0 Then
            
            For iTaxes = Me.Count To 1 Step -1
                        
                If Me(iTaxes).Country = CountryCode Then
                    If InStr(Codes, ",") > 0 Then
                        If InStr(Me(iTaxes).Code, Codes) > 0 Then
                            FindTaxRate = FindTaxRate + Me(iTaxes).Rate
                        End If
                    ElseIf Codes = Me(iTaxes).Code Then
                        FindTaxRate = Me(iTaxes).Rate
                        Exit For
                    End If
                End If
                                    
            Next
        Else
            FindTaxRate = 0
        End If
        
    Case Else
    
        If Me.Count > 0 Then
            
            For iTaxes = Me.Count To 1 Step -1
                        
                If Me(iTaxes).Country = CountryCode Then
                    If InStr(Codes, ",") > 0 Then
                        If InStr(Codes, Me(iTaxes).Code) > 0 Then
                            If Me(iTaxes).RangeMin >= AmountBeingTaxed Then
                                If Me(iTaxes).RangeMax = -1 Then
                                    FindTaxRate = Me(iTaxes).cRate
                                    Exit For
                                ElseIf Me(iTaxes).RangeMax <= AmountBeingTaxed Then
                                    FindTaxRate = Me(iTaxes).cRate
                                    Exit For
                                End If
                            End If
                        End If
                    ElseIf Codes = Me(iTaxes).Code Then
                        If Me(iTaxes).RangeMin >= AmountBeingTaxed Then
                            If Me(iTaxes).RangeMax = -1 Then
                                FindTaxRate = Me(iTaxes).cRate
                                Exit For
                            ElseIf Me(iTaxes).RangeMax <= AmountBeingTaxed Then
                                FindTaxRate = Me(iTaxes).cRate
                                Exit For
                            End If
                        End If
                        
                    End If
                End If
                                    
            Next
        Else
            FindTaxRate = 0
        End If
        
    End Select
    
End Function


Public Function Add(Key As String, Code As String, Country As String, Rate As Single, iFlag As Integer, Desc As String, lGroup As Long, RangeMin As Currency, RangeMax As Currency, FlatRate As Currency, Optional sKey As String) As clsTaxExpress
    'create a new object
    Dim objNewMember As clsTaxExpress
    Set objNewMember = New clsTaxExpress


    'set the properties passed into the method
    objNewMember.Key = Key
    objNewMember.Code = Code
    objNewMember.Country = Country
    objNewMember.Rate = Rate
    objNewMember.iFlag = iFlag
    objNewMember.Desc = Desc
    objNewMember.lGroup = lGroup
    objNewMember.RangeMin = RangeMin
    objNewMember.RangeMax = RangeMax
    objNewMember.cRate = FlatRate
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As clsTaxExpress
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

