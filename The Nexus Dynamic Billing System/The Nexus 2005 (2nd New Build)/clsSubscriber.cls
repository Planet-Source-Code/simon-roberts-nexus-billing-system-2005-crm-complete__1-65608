VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubscriber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"col_Domains"
Attribute VB_Ext_KEY = "Member1" ,"col_subAliases"
Attribute VB_Ext_KEY = "Member2" ,"col_subDSLLink"
Attribute VB_Ext_KEY = "Member3" ,"col_subEditLog"
Attribute VB_Ext_KEY = "Member4" ,"col_subEmails"
Attribute VB_Ext_KEY = "Member5" ,"col_subHardware"
Attribute VB_Ext_KEY = "Member6" ,"col_subPaymentSet"
Attribute VB_Ext_KEY = "Member7" ,"col_subPhoneNo"
Attribute VB_Ext_KEY = "Member8" ,"col_subRadius"
Attribute VB_Ext_KEY = "Member9" ,"col_subReferals"
Attribute VB_Ext_KEY = "Member10" ,"col_subServices"
Attribute VB_Ext_KEY = "Member11" ,"col_subTrans"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member12" ,"colSnailMail"
'local variable(s) to hold property value(s)
Private mvarfAccountName As String 'local copy
Private mvarfRecID As Long 'local copy
Private mvarfActivationDate As Date 'local copy
Private mvarfExpiryDate As Date 'local copy
Private mvarfCreationDate As Date 'local copy
Private mvarfDOB As Date 'local copy
Private mvarfPaymentID As Long 'local copy
Private mvarfActivateDateSet As Byte 'local copy
Private mvarfExpiryDateSet As Byte 'local copy
Private mvarfProcessFlag As Byte 'local copy
Private mvarfCancelled As Integer 'local copy
Private mvarfChecked As Integer 'local copy
Private mvarfClassification As Long 'local copy
Private mvarfRealm As String 'local copy
Private mvarfPayIntervalType As String 'local copy
Private mvarfPayInterval As Integer 'local copy
Private mvarfFlagA_RecID As Long 'local copy
Private mvarfFlagB_RecID As Long 'local copy
Private mvarfSysopID As Long 'local copy
Private mvarfVirtualID As Long 'local copy
Private mvarfAgencyID As Long 'local copy
Private mvarfsfStartTime As Date 'local copy
Private mvarfsfCycle_Upload As Double 'local copy
Private mvarfsfCycle_Download As Double 'local copy
Private mvarfsfCycle_Mins As Double 'local copy
Private mvarfCrystallise As Byte 'local copy
Private mvarfOmnivorous As Byte 'local copy
Private mvarFCarnivorous As Byte 'local copy
Private mvarfFlagASet As Date 'local copy
Private mvarfFlagBSet As Date 'local copy
Private mvarfBillingDate As Date 'local copy
Private mvarfgUsername As String 'local copy
Private mvarfgPassword As String 'local copy
Private mvarfAboutUs As Byte 'local copy
Private mvarbInitalised As Boolean 'local copy
Private mvarcolSnailMail As New colSnailMail 'local copy
Private mvarcol_subAliases As New col_subAliases  'local copy
Private mvarcol_subDSLLink As New col_subDSLLink 'local copy
Private mvarcol_subEditLog As New col_subEditLog 'local copy
Private mvarcol_subHardware As New col_subHardware 'local copy
Private mvarcol_subPhoneNo As New col_subPhoneNo 'local copy
Private mvarcol_subRadius As New col_subRadius 'local copy
Private mvarcol_subServices As New col_subServices 'local copy
Private mvarcol_subEmails As New col_subEmails 'local copy
Private mvarcol_subReferals As New col_subReferals 'local copy
Private mvarcol_subTrans As New col_subTrans 'local copy
Private mvarcol_Domains As New col_Domains 'local copy
Private mvarfftpPathKey As String 'local copy
Private mvarcolPaymentSet As col_subPaymentSet


Public Property Get colPaymentSet() As col_subPaymentSet
    If mvarcolPaymentSet Is Nothing Then
        Set mvarcolPaymentSet = New col_subPaymentSet
    End If


    Set colPaymentSet = mvarcolPaymentSet
End Property


Public Property Set colPaymentSet(vData As col_subPaymentSet)
    Set mvarcolPaymentSet = vData
End Property
Private Sub Class_Terminate()
    Set mvarcolPaymentSet = Nothing
End Sub



Public Property Let fftpPathKey(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fftpPathKey = 5
    mvarfftpPathKey = vData
End Property


Public Property Get fftpPathKey() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fftpPathKey
    fftpPathKey = mvarfftpPathKey
End Property



Public Function ClearSession(ByVal SESSION As String, Optional lv1 As ListView, Optional lv2 As ListView, Optional lv3 As ListView, Optional lv4 As ListView, Optional lv5 As ListView, Optional lv6 As ListView, Optional lv7 As ListView, Optional lv8 As ListView) As Long

    With Me
        
        ClearSession = ClearSession + .CheckAndRemove(.col_Domains, SESSION)
        ClearSession = ClearSession + .CheckAndRemove(.col_subAliases, SESSION)
        ClearSession = ClearSession + .CheckAndRemove(.col_subDSLLink, SESSION)
        ClearSession = ClearSession + .CheckAndRemove(.col_subEmails, SESSION)
        ClearSession = ClearSession + .CheckAndRemove(.col_subHardware, SESSION)
        ClearSession = ClearSession + .CheckAndRemove(.col_subPhoneNo, SESSION)
        ClearSession = ClearSession + .CheckAndRemove(.col_subRadius, SESSION)
        ClearSession = ClearSession + .CheckAndRemove(.col_subServices, SESSION)
        ClearSession = ClearSession + .CheckAndRemove(.col_subTrans, SESSION)
        
        If Not lv1 Is Nothing Then ClearSession = ClearSession + .CheckAndRemove(lv1, SESSION)
        If Not lv2 Is Nothing Then ClearSession = ClearSession + .CheckAndRemove(lv2, SESSION)
        If Not lv3 Is Nothing Then ClearSession = ClearSession + .CheckAndRemove(lv3, SESSION)
        If Not lv4 Is Nothing Then ClearSession = ClearSession + .CheckAndRemove(lv4, SESSION)
        If Not lv5 Is Nothing Then ClearSession = ClearSession + .CheckAndRemove(lv5, SESSION)
        If Not lv6 Is Nothing Then ClearSession = ClearSession + .CheckAndRemove(lv6, SESSION)
        If Not lv7 Is Nothing Then ClearSession = ClearSession + .CheckAndRemove(lv7, SESSION)
        If Not lv8 Is Nothing Then ClearSession = ClearSession + .CheckAndRemove(lv8, SESSION)
        
    End With
    
End Function

Public Function CheckAndRemove(Optional mCol As Object, Optional SESSION As String, Optional lv As ListView) As Long
    Dim lx As Long
    
    If Not mCol Is Nothing Then
        If mCol.Count > 0 Then
    
            For lx = mCol.Count To 1 Step -1
                If InStr(mCol(lx).Key, SESSION) > 0 Then
                    mCol.Remove lx
                    CheckAndRemove = CheckAndRemove + 1
                End If
            Next
        End If
    End If
    
    If Not lv Is Nothing Then
        If lv.ListItems.Count > 0 Then
            
            For lx = lv.ListItems.Count To 1 Step -1
                If InStr(lv.ListItems(lx).Key, SESSION) > 0 Then
                    lv.ListItems.Remove lx
                    CheckAndRemove = CheckAndRemove + 1
                End If
            Next
        End If
    End If
    
End Function

Public Property Set col_Domains(ByVal vData As col_Domains)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.col_Domains = Form1
    Set mvarcol_Domains = vData
End Property


Public Property Get col_Domains() As col_Domains
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.col_Domains
    Set col_Domains = mvarcol_Domains
End Property



Public Property Set col_subTrans(ByVal vData As col_subTrans)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.col_subTrans = Form1
    Set mvarcol_subTrans = vData
End Property


Public Property Get col_subTrans() As col_subTrans
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.col_subTrans
    Set col_subTrans = mvarcol_subTrans
End Property



Public Property Set col_subReferals(ByVal vData As col_subReferals)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.col_subReferals = Form1
    Set mvarcol_subReferals = vData
End Property


Public Property Get col_subReferals() As col_subReferals
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.col_subReferals
    Set col_subReferals = mvarcol_subReferals
End Property



Public Property Set col_subEmails(ByVal vData As col_subEmails)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.col_subEmails = Form1
    Set mvarcol_subEmails = vData
End Property


Public Property Get col_subEmails() As col_subEmails
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.col_subEmails
    Set col_subEmails = mvarcol_subEmails
End Property



Public Property Set col_subServices(ByVal vData As col_subServices)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.col_subServices = Form1
    Set mvarcol_subServices = vData
End Property


Public Property Get col_subServices() As col_subServices
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.col_subServices
    Set col_subServices = mvarcol_subServices
End Property



Public Function GetClient(ByVal ClientID As Long, ByVal oConn As ADODB.Connection) As Byte

    Const bNoReturn = 0
    Const bResult = 1
    
    Dim rsload As ADODB.Recordset
    
    If MySQL.OpenTable(oConn, rsload, , "select *, AES_DECRYPT(gPassword,'" + odb.colSalts.ReturnSalt(PWSalt) + "') as gPWD from accountinfo where RecID = '" & ClientID & "'") = True Then
        If rsload.State = adStateOpen Then
            If rsload.RecordCount > 0 Then
                GetClient = bResult
                Me.fftpPathKey = IIf(IsNull(rsload!ftpPathKey), "", rsload!ftpPathKey)
                Me.fAboutUs = IIf(IsNull(rsload!AboutUS), 0, rsload!AboutUS)
                Me.fAccountName = IIf(IsNull(rsload!AccountName), "", rsload!AccountName)
                Me.fActivateDateSet = IIf(IsNull(rsload!ActivateDateSet), 0, rsload!ActivateDateSet)
                Me.fActivationDate = IIf(IsNull(rsload!ActivationDate), DateAdd("w", 2, sysnow), rsload!ActivationDate)
                Me.fAgencyID = IIf(IsNull(rsload!AgencyID), 0, rsload!AgencyID)
                Me.fBillingDate = IIf(IsNull(rsload!BillingDate), DateAdd("d", 1, sysnow), rsload!BillingDate)
                Me.fCancelled = IIf(IsNull(rsload!Cancelled), 0, Val(rsload!Cancelled))
                Me.FCarnivorous = IIf(IsNull(rsload!Carnivorous), 0, rsload!Carnivorous)
                Me.fChecked = IIf(IsNull(rsload!Checked), 0, Val(rsload!Checked))
                Me.fClassification = IIf(IsNull(rsload!Classification), 0, rsload!Classification)
                Me.fDOB = IIf(IsNull(rsload!DOB), sysnow, rsload!DOB)
                Me.fExpiryDate = IIf(IsNull(rsload!ExpiryDate), sysnow, rsload!ExpiryDate)
                Me.fExpiryDateSet = IIf(IsNull(rsload!ExpiryDateSet), 0, rsload!ExpiryDateSet)
                Me.fFlagA_RecID = IIf(IsNull(rsload!FlagA_RecID), 0, rsload!FlagA_RecID)
                Me.fFlagASet = IIf(IsNull(rsload!FlagASet), 0, rsload!FlagASet)
                Me.fFlagB_RecID = IIf(IsNull(rsload!FlagB_RecID), 0, rsload!FlagB_RecID)
                Me.fFlagBSet = IIf(IsNull(rsload!FlagBSet), 0, rsload!FlagBSet)
                Me.fgPassword = IIf(IsNull(rsload!gPWD), "myvoice", rsload!gPWD)
                Me.fgUsername = IIf(IsNull(rsload!gUsername), "guest" & Round(Rnd * 99999), rsload!gUsername)
                Me.fOmnivorous = IIf(IsNull(rsload!Omnivorous), 0, rsload!Omnivorous)
                Me.fPayInterval = IIf(IsNull(rsload!PayInterval), "m", rsload!PayInterval)
                Me.fPayIntervalType = IIf(IsNull(rsload!PayIntervalType), "1", rsload!PayIntervalType)
                Me.fPaymentID = IIf(IsNull(rsload!PaymentID), 0, rsload!PaymentID)
                Me.fProcessFlag = IIf(IsNull(rsload!PaymentID), 0, rsload!ProcessFlag)
                Me.fRealm = IIf(IsNull(rsload!Realm), "", rsload!Realm)
                Me.fRecID = IIf(IsNull(rsload!RecID), 0, rsload!RecID)
                Me.fsfCycle_Download = IIf(IsNull(rsload!sfCycle_Download), 0, rsload!sfCycle_Download)
                Me.fsfCycle_Mins = IIf(IsNull(rsload!sfCycle_Mins), 0, rsload!sfCycle_Mins)
                Me.fsfCycle_Upload = IIf(IsNull(rsload!sfCycle_Upload), 0, rsload!sfCycle_Upload)
                Me.fsfStartTime = IIf(IsNull(rsload!sfStartTime), sysnow, rsload!sfStartTime)
                Me.fSysopID = IIf(IsNull(rsload!SysopID), 0, rsload!SysopID)
                Me.fVirtualID = IIf(IsNull(rsload!VirtualID), 0, rsload!VirtualID)
                                
                Me.col_subPhoneNo.CipherSalt = odb.colSalts.ReturnSalt("phonesalt")
                Me.col_subEmails.CipherSalt = odb.colSalts.ReturnSalt(EMAILSalt)
                
                Call Me.col_subTrans.GetClient(Me.fRecID, oConn)
                Call Me.col_subAliases.GetClient(Me.fRecID, oConn)
                Call Me.col_subDSLLink.GetClient(Me.fRecID, oConn)
                Call Me.col_subEditLog.GetClient(Me.fRecID, oConn)
                Call Me.col_subEmails.GetClient(Me.fRecID, oConn)
                Call Me.col_subHardware.GetClient(Me.fRecID, oConn)
                Call Me.col_subPhoneNo.GetClient(Me.fRecID, oConn)
                Call Me.col_subRadius.GetClient(Me.fRecID, oConn)
                Call Me.col_subServices.GetClient(Me.fRecID, oConn)
                Call Me.col_subReferals.GetClient(Me.fRecID, oConn)
                Call Me.colSnailMail.GetClient(Me.fRecID, oConn)
                Call Me.col_Domains.GetClient(Me.fRecID, oConn)
                Call Me.colPaymentSet.GetClient(Me.fRecID, oConn)
                
                Me.bInitalised = True
            End If
        End If
    End If
    
    
End Function

Public Property Set col_subRadius(ByVal vData As col_subRadius)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.col_subRadius = Form1
    Set mvarcol_subRadius = vData
End Property


Public Property Get col_subRadius() As col_subRadius
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.col_subRadius
    Set col_subRadius = mvarcol_subRadius
End Property



Public Function Commit(ByVal oConn As ADODB.Connection) As Long

    If oConn.State = adStateOpen Then
    
        If Me.fRecID = 0 Then
            Me.fRecID = MySQL.GetTMPRecID("accountinfo", oConn, "RecID")
        End If
        
        Dim SQL As String
        
        SQL = "update accountinfo set "
        SQL = SQL + "AboutUs = '" & Me.fAboutUs & "', "
        SQL = SQL + "AccountName = '" & MySQL.ESC(Me.fAccountName) & "', "
        SQL = SQL + "ActivateDateSet = '" & Me.fActivateDateSet & "', "
        SQL = SQL + "ActivationDate = '" & Format(Me.fActivationDate, "yyyy-mm-dd ttttt") & "', "
        SQL = SQL + "AgencyID = '" & Me.fAgencyID & "', "
        SQL = SQL + "BillingDate = '" & Format(Me.fBillingDate, "dd mmm yyyy ttttt") & "', "
        SQL = SQL + "Cancelled = '" & Me.fCancelled & "', "
        SQL = SQL + "Carnivorous = '" & Me.FCarnivorous & "', "
        SQL = SQL + "Checked = '" & Me.fChecked & "', "
        SQL = SQL + "Classification = '" & Me.fClassification & "', "
        SQL = SQL + "Crystallise = '" & Me.fCrystallise & "', "
        SQL = SQL + "DOB = '" & Format(Me.fDOB, "dd mmm yyyy ttttt") & "', "
        SQL = SQL + "ExpiryDate = '" & Format(Me.fExpiryDate, "yyyy-mm-dd ttttt") & "', "
        SQL = SQL + "ExpiryDateSet = '" & Me.fExpiryDateSet & "', "
        SQL = SQL + "FlagA_RecID = '" & Me.fFlagA_RecID & "', "
        SQL = SQL + "FlagB_RecID = '" & Format(Me.fFlagASet, "yyyy-mm-dd ttttt") & "', "
        SQL = SQL + "FlagB_RecID = '" & Me.fFlagB_RecID & "', "
        SQL = SQL + "FlagBSet = '" & Format(Me.fFlagBSet, "yyyy-mm-dd ttttt") & "', "
        SQL = SQL + "gPassword = AES_ENCRYPT('" & Me.fgPassword & "','" + odb.colSalts.ReturnSalt(PWSalt) + "'), "
        SQL = SQL + "gUsername = '" & Me.fgUsername & "', "
        SQL = SQL + "Omnivorous = '" & Me.fOmnivorous & "', "
        SQL = SQL + "PayInterval  = '" & Me.fPayInterval & "', "
        SQL = SQL + "PayIntervalType = '" & Me.fPayIntervalType & "', "
        SQL = SQL + "PaymentID = '" & Me.fPaymentID & "', "
        SQL = SQL + "ProcessFlag = '" & Me.fProcessFlag & "', "
        SQL = SQL + "Realm = '" & MySQL.ESC(Me.fRealm) & "', "
        SQL = SQL + "SysopID = '" & Me.fSysopID & "', "
        SQL = SQL + "ftpPathKey = '" & Me.fftpPathKey & "', "
        SQL = SQL + "VirtualID = '" & Me.fVirtualID & "' "
        SQL = SQL + "where RecID = '" & Me.fRecID & "'"
        
        Call MySQL.Execute(oConn, SQL)
        Call Me.Refesh_acci_RecID
        
        If Me.colSnailMail.Commit(oConn) > 0 Then
            Call Me.LinkIDX("DefaultShip")
        End If
        
        If Me.col_subRadius.Commit(oConn) > 0 Then
            Call Me.LinkIDX("Radius")
        End If
        
        If Me.col_subServices.Commit(oConn) > 0 Then
            Call Me.LinkIDX("Services")
            Call Me.LinkIDX("Referals")
        End If

        
        If Me.col_subTrans.Commit(oConn) > 0 Then
            
        End If
    
        If Me.col_subReferals.Commit(oConn) > 0 Then
            
        End If
        
        If Me.col_subPhoneNo.Commit(oConn) > 0 Then
            
        End If
        
        If Me.col_subHardware.Commit(oConn) > 0 Then
            
        End If
        
        If Me.col_subEmails.Commit(oConn) > 0 Then
            
        End If
        
        If Me.col_subEmails.Commit(oConn) > 0 Then
            
        End If
        
        If Me.col_subAliases.Commit(oConn) > 0 Then
            
        End If
        
        If Me.col_subDSLLink.Commit(oConn) > 0 Then
            
        End If
        
        If Me.col_Domains.Commit(oConn) > 0 Then
            
        End If
                
        If Me.col_subEditLog.Commit(oConn) > 0 Then
            
        End If
        
        If Me.colPaymentSet.Commit(oConn) > 0 Then
            
        End If
    End If
    
    
End Function

Function LinkIDX(IDXCode As String)
    
    On Error Resume Next
    
    Select Case IDXCode
    Case "DefaultShip"
        Dim oSVR As cls_subServices
        For Each oSVR In Me.col_subServices
            If InStr(oSVR.Key, "NEW") > 0 Then oSVR.DefaultShippingID = Val(Mid(Me.colSnailMail(oSVR.DefaultShippingID).Key, 6))
        Next
    Case "Radius"
        Dim oRDS As cls_subServices
        For Each oRDS In Me.col_subServices
            If InStr(oRDS.Key, "NEW") > 0 And oRDS.RadiusID > 0 Then oRDS.RadiusID = Val(Mid(Me.col_subRadius(oRDS.RadiusID).Key, 6))
        Next
        Dim oDSL As cls_subDSLLink
        For Each oDSL In Me.col_subDSLLink
            If InStr(oDSL.Key, "NEW") > 0 Then oDSL.RadiusID = Val(Mid(Me.col_subRadius(oRDS.RadiusID).Key, 6))
        Next
    Case "Services"
        Dim oINV As cls_subTrans
        For Each oINV In Me.col_subDSLLink
            If InStr(oINV.Key, "NEW") > 0 Then oINV.PlanServiceID = Val(Mid(Me.col_subServices(oINV.PlanServiceID).Key, 6))
        Next
    Case "Referals"
        Dim oRef As cls_subReferals
        For Each oRef In Me.col_subReferals
            If InStr(oRef.Key, "NEW") > 0 Then oRef.acciServiceID = Val(Mid(Me.col_subServices(oRef.acciServiceID).Key, 6))
        Next
    End Select
    
    
    
End Function

Function Refesh_acci_RecID()


    Dim lx As Long
    
    For lx = 1 To Me.col_subEditLog.Count
        Me.col_subEditLog(lx).acci_RecID = Me.fRecID
    Next
            
    For lx = 1 To Me.col_Domains.Count
        Me.col_Domains(lx).acci_RecID = Me.fRecID
    Next
        
    For lx = 1 To Me.col_subAliases.Count
        Me.col_subAliases(lx).acci_RecID = Me.fRecID
    Next
        
    For lx = 1 To Me.col_subDSLLink.Count
        Me.col_subDSLLink(lx).acci_RecID = Me.fRecID
    Next
        
    For lx = 1 To Me.col_subEmails.Count
        Me.col_subEmails(lx).acci_RecID = Me.fRecID
    Next
        
    For lx = 1 To Me.col_subHardware.Count
        Me.col_subHardware(lx).acci_RecID = Me.fRecID
    Next
        
    For lx = 1 To Me.col_subPhoneNo.Count
        Me.col_subPhoneNo(lx).acci_RecID = Me.fRecID
    Next
        
    For lx = 1 To Me.col_subRadius.Count
        Me.col_subRadius(lx).acci_RecID = Me.fRecID
    Next
        
    For lx = 1 To Me.colPaymentSet.Count
        Me.colPaymentSet(lx).acci_RecID = Me.fRecID
    Next
        
    Dim oRef As cls_subReferals
    
    For Each oRef In col_subReferals
        oRef.acci_RecID = Me.fRecID
    Next
        
    For lx = 1 To Me.col_subServices.Count
        Me.col_subServices(lx).acci_RecID = Me.fRecID
    Next
        
    Dim subTrans As cls_subTrans
    For Each subTrans In Me.col_subTrans
        subTrans.acci_RecID = Me.fRecID
    Next
        
    For lx = 1 To Me.colSnailMail.Count
        Me.colSnailMail(lx).acci_RecID = Me.fRecID
    Next
        
End Function

Public Property Set col_subPhoneNo(ByVal vData As col_subPhoneNo)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.col_subPhoneNo = Form1
    Set mvarcol_subPhoneNo = vData
End Property


Public Property Get col_subPhoneNo() As col_subPhoneNo
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.col_subPhoneNo
    Set col_subPhoneNo = mvarcol_subPhoneNo
End Property



Public Property Set col_subHardware(ByVal vData As col_subHardware)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.col_subHardware = Form1
    Set mvarcol_subHardware = vData
End Property


Public Property Get col_subHardware() As col_subHardware
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.col_subHardware
    Set col_subHardware = mvarcol_subHardware
End Property



Public Property Set col_subEditLog(ByVal vData As col_subEditLog)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.col_subEditLog = Form1
    Set mvarcol_subEditLog = vData
End Property


Public Property Get col_subEditLog() As col_subEditLog
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.col_subEditLog
    Set col_subEditLog = mvarcol_subEditLog
End Property



Public Property Set col_subDSLLink(ByVal vData As col_subDSLLink)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.col_subDSLLink = Form1
    Set mvarcol_subDSLLink = vData
End Property


Public Property Get col_subDSLLink() As col_subDSLLink
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.col_subDSLLink
    Set col_subDSLLink = mvarcol_subDSLLink
End Property



Public Property Set col_subAliases(ByVal vData As col_subAliases)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.col_subAliases = Form1
    Set mvarcol_subAliases = vData
End Property


Public Property Get col_subAliases() As col_subAliases
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.col_subAliases
    Set col_subAliases = mvarcol_subAliases
End Property



Public Property Set colSnailMail(ByVal vData As colSnailMail)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colSnailMail = Form1
    Set mvarcolSnailMail = vData
End Property


Public Property Get colSnailMail() As colSnailMail
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colSnailMail
    Set colSnailMail = mvarcolSnailMail
End Property



Public Property Let bInitalised(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bInitalised = 5
    mvarbInitalised = vData
End Property


Public Property Get bInitalised() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bInitalised
    bInitalised = mvarbInitalised
End Property



Public Property Let fAboutUs(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fAboutUs = 5
    mvarfAboutUs = vData
End Property


Public Property Get fAboutUs() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fAboutUs
    fAboutUs = mvarfAboutUs
End Property



Public Property Let fgPassword(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fgPassword = 5
    mvarfgPassword = vData
End Property


Public Property Get fgPassword() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fgPassword
    fgPassword = mvarfgPassword
End Property



Public Property Let fgUsername(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fgUsername = 5
    mvarfgUsername = vData
End Property


Public Property Get fgUsername() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fgUsername
    fgUsername = mvarfgUsername
End Property



Public Property Let fBillingDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fBillingDate = 5
    mvarfBillingDate = vData
End Property


Public Property Get fBillingDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fBillingDate
    fBillingDate = mvarfBillingDate
End Property



Public Property Let fFlagBSet(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fFlagBSet = 5
    mvarfFlagBSet = vData
End Property


Public Property Get fFlagBSet() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fFlagBSet
    fFlagBSet = mvarfFlagBSet
End Property



Public Property Let fFlagASet(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fFlagASet = 5
    mvarfFlagASet = vData
End Property


Public Property Get fFlagASet() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fFlagASet
    fFlagASet = mvarfFlagASet
End Property



Public Property Let FCarnivorous(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FCarnivorous = 5
    mvarFCarnivorous = vData
End Property


Public Property Get FCarnivorous() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FCarnivorous
    FCarnivorous = mvarFCarnivorous
End Property



Public Property Let fOmnivorous(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fOmnivorous = 5
    mvarfOmnivorous = vData
End Property


Public Property Get fOmnivorous() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fOmnivorous
    fOmnivorous = mvarfOmnivorous
End Property



Public Property Let fCrystallise(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fCrystallise = 5
    mvarfCrystallise = vData
End Property


Public Property Get fCrystallise() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fCrystallise
    fCrystallise = mvarfCrystallise
End Property



Public Property Let fsfCycle_Mins(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fsfCycle_Mins = 5
    mvarfsfCycle_Mins = vData
End Property


Public Property Get fsfCycle_Mins() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fsfCycle_Mins
    fsfCycle_Mins = mvarfsfCycle_Mins
End Property



Public Property Let fsfCycle_Download(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fsfCycle_Download = 5
    mvarfsfCycle_Download = vData
End Property


Public Property Get fsfCycle_Download() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fsfCycle_Download
    fsfCycle_Download = mvarfsfCycle_Download
End Property



Public Property Let fsfCycle_Upload(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fsfCycle_Upload = 5
    mvarfsfCycle_Upload = vData
End Property


Public Property Get fsfCycle_Upload() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fsfCycle_Upload
    fsfCycle_Upload = mvarfsfCycle_Upload
End Property



Public Property Let fsfStartTime(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fsfStartTime = 5
    mvarfsfStartTime = vData
End Property


Public Property Get fsfStartTime() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fsfStartTime
    fsfStartTime = mvarfsfStartTime
End Property



Public Property Let fAgencyID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fAgencyID = 5
    mvarfAgencyID = vData
End Property


Public Property Get fAgencyID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fAgencyID
    fAgencyID = mvarfAgencyID
End Property



Public Property Let fVirtualID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fVirtualID = 5
    mvarfVirtualID = vData
End Property


Public Property Get fVirtualID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fVirtualID
    fVirtualID = mvarfVirtualID
End Property



Public Property Let fSysopID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fSysopID = 5
    mvarfSysopID = vData
End Property


Public Property Get fSysopID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fSysopID
    fSysopID = mvarfSysopID
End Property



Public Property Let fFlagB_RecID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fFlagB_RecID = 5
    mvarfFlagB_RecID = vData
End Property


Public Property Get fFlagB_RecID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fFlagB_RecID
    fFlagB_RecID = mvarfFlagB_RecID
End Property



Public Property Let fFlagA_RecID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fFlagA_RecID = 5
    mvarfFlagA_RecID = vData
End Property


Public Property Get fFlagA_RecID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fFlagA_RecID
    fFlagA_RecID = mvarfFlagA_RecID
End Property



Public Property Let fPayInterval(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fPayInterval = 5
    mvarfPayInterval = vData
End Property


Public Property Get fPayInterval() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fPayInterval
    fPayInterval = mvarfPayInterval
End Property



Public Property Let fPayIntervalType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fPayIntervalType = 5
    mvarfPayIntervalType = vData
End Property


Public Property Get fPayIntervalType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fPayIntervalType
    fPayIntervalType = mvarfPayIntervalType
End Property



Public Property Let fRealm(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fRealm = 5
    mvarfRealm = vData
End Property


Public Property Get fRealm() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fRealm
    fRealm = mvarfRealm
End Property



Public Property Let fClassification(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fClassification = 5
    mvarfClassification = vData
End Property


Public Property Get fClassification() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fClassification
    fClassification = mvarfClassification
End Property



Public Property Let fChecked(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fChecked = 5
    mvarfChecked = vData
End Property


Public Property Get fChecked() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fChecked
    fChecked = mvarfChecked
End Property



Public Property Let fCancelled(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fCancelled = 5
    mvarfCancelled = vData
End Property


Public Property Get fCancelled() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fCancelled
    fCancelled = mvarfCancelled
End Property



Public Property Let fProcessFlag(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fProcessFlag = 5
    mvarfProcessFlag = vData
End Property


Public Property Get fProcessFlag() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fProcessFlag
    fProcessFlag = mvarfProcessFlag
End Property



Public Property Let fExpiryDateSet(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fExpiryDateSet = 5
    mvarfExpiryDateSet = vData
End Property


Public Property Get fExpiryDateSet() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fExpiryDateSet
    fExpiryDateSet = mvarfExpiryDateSet
End Property



Public Property Let fActivateDateSet(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fActivateDateSet = 5
    mvarfActivateDateSet = vData
End Property


Public Property Get fActivateDateSet() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fActivateDateSet
    fActivateDateSet = mvarfActivateDateSet
End Property



Public Property Let fPaymentID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fPaymentID = 5
    mvarfPaymentID = vData
End Property


Public Property Get fPaymentID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fPaymentID
    fPaymentID = mvarfPaymentID
End Property



Public Property Let fDOB(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fDOB = 5
    mvarfDOB = vData
End Property


Public Property Get fDOB() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fDOB
    fDOB = mvarfDOB
End Property



Public Property Let fCreationDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fCreationDate = 5
    mvarfCreationDate = vData
End Property


Public Property Get fCreationDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fCreationDate
    fCreationDate = mvarfCreationDate
End Property



Public Property Let fExpiryDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fExpiryDate = 5
    mvarfExpiryDate = vData
End Property


Public Property Get fExpiryDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fExpiryDate
    fExpiryDate = mvarfExpiryDate
End Property



Public Property Let fActivationDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fActivationDate = 5
    mvarfActivationDate = vData
End Property


Public Property Get fActivationDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fActivationDate
    fActivationDate = mvarfActivationDate
End Property



Public Property Let fRecID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fRecID = 5
    mvarfRecID = vData
End Property


Public Property Get fRecID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fRecID
    fRecID = mvarfRecID
End Property



Public Property Let fAccountName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fAccountName = 5
    mvarfAccountName = vData
End Property


Public Property Get fAccountName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fAccountName
    fAccountName = mvarfAccountName
End Property



