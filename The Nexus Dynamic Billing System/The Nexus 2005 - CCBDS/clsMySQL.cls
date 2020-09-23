VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMySQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Function rGroupName(sHeader As String, RecID As Long, oConn As ADODB.Connection, dConn As DAO.Connection) As String

    Dim rsGroup As DAO.Recordset
    Dim rsLoad As ADODB.Recordset
    
    If Me.OpenTable(oConn, rsLoad, , "Select * from RadiusAccounts Where RecID = " & RecID) = True Then
        If rsLoad.RecordCount > 0 Then
            'MsgBox Me.ESC(sHeader & "-s" & rsLoad!SessionTimeout & "-i" & rsLoad!IdleTimeout & "-d" & IIf(IsNull(rsLoad!PrimaryDNS), "primary", rsLoad!PrimaryDNS) & "-" & rsLoad!SecondaryDNS) & "'"
            If Me.OpenTableDef(dConn, rsGroup, , "Select * from RadiusRadGroupCheck where GroupName = '" & Me.ESC(sHeader & "-s" & rsLoad!SessionTimeout & "-i" & rsLoad!IdleTimeout & "-d" & IIf(IsNull(rsLoad!PrimaryDNS), "primary", rsLoad!PrimaryDNS) & "-" & rsLoad!SecondaryDNS) & "'") = True Then
                If rsGroup.RecordCount > 0 Then
                    rGroupName = sHeader & "-s" & rsLoad!SessionTimeout & "-i" & rsLoad!IdleTimeout & "-d" & IIf(IsNull(rsLoad!PrimaryDNS), "primary", rsLoad!PrimaryDNS) & "-" & rsLoad!SecondaryDNS
                Else
                    rGroupName = sHeader & "-s" & rsLoad!SessionTimeout & "-i" & rsLoad!IdleTimeout & "-d" & IIf(IsNull(rsLoad!PrimaryDNS), "primary", rsLoad!PrimaryDNS) & "-" & rsLoad!SecondaryDNS
                    dConn.Execute "Insert Into RadiusGroupReply (groupname, attribute, value, op) VALUES ('" & Me.ESC(rGroupName) & "','Framed-Compression','Van-Jacobsen-TCP-IP',':=')"
                    If Not IsNull(rsLoad!FramedProtocol) Then dConn.Execute "Insert Into RadiusGroupReply (groupname, attribute, value, op) VALUES ('" & Me.ESC(rGroupName) & "','Framed-Protocol','" & rsLoad!FramedProtocol & "',':=')"
                    If Not IsNull(rsLoad!ServiceType) Then dConn.Execute "Insert Into RadiusGroupReply (groupname, attribute, value, op) VALUES ('" & Me.ESC(rGroupName) & "','Service-Type','" & rsLoad!ServiceType & "',':=')"
                    If Not IsNull(rsLoad!SecondaryDNS) Then dConn.Execute "Insert Into RadiusGroupReply (groupname, attribute, value, op) VALUES ('" & Me.ESC(rGroupName) & "','Ascend-Client-Secondary-DNS','" & rsLoad!SecondaryDNS & "',':=')"
                    If Not IsNull(rsLoad!PrimaryDNS) Then dConn.Execute "Insert Into RadiusGroupReply (groupname, attribute, value, op) VALUES ('" & Me.ESC(rGroupName) & "','Ascend-Client-Primary-DNS','" & rsLoad!PrimaryDNS & "',':=')"
                    If Not IsNull(rsLoad!IdleTimeout) Then If rsLoad!IdleTimeout <> 0 Then dConn.Execute "Insert Into RadiusGroupReply (groupname, attribute, value, op) VALUES ('" & Me.ESC(rGroupName) & "','Idle-Timeout','" & rsLoad!IdleTimeout & "',':=')"
                    If Not IsNull(rsLoad!SessionTimeout) Then If rsLoad!SessionTimeout <> 0 Then dConn.Execute "Insert Into RadiusGroupReply (groupname, attribute, value, op) VALUES ('" & Me.ESC(rGroupName) & "','Session-Limit','" & rsLoad!SessionTimeout & "',':=')"
                    dConn.Execute "Insert Into RadiusRadGroupCheck (GroupName, Attribute, Value, Op) VALUES ('" & Me.ESC(rGroupName) & "','Auth-Type','Local',':=')"
                End If
            Else
                rGroupName = "unpaid"
            End If
        Else
            rGroupName = "unpaid"
        End If
    Else
        rGroupName = "unpaid"
    End If
    
    
End Function

Function DAOConn(Optional db As String = "", Optional server As String = "", Optional UID As String = "", Optional PWD As String = "", Optional ByRef oConn As DAO.Connection, Optional ByRef wrkODBC As DAO.Workspace) As Boolean

    If db = "" Then db = GetSetting(App.ProductName, "DB", "DB", "no1_billing")
    If server = "" Then server = GetSetting(App.ProductName, "DB", "Server", "beast0.no1.com.au")
    If UID = "" Then UID = GetSetting(App.ProductName, "DB", "UID", "noosa")
    If PWD = "" Then PWD = GetSetting(App.ProductName, "DB", "PWD", "wizard")
        
    On Error GoTo ErrOcc

    Set wrkODBC = CreateWorkspace("noosa", UID, PWD, dbUseODBC)
    
    Select Case Login.bTestBench
    Case False
        Set oConn = wrkODBC.OpenConnection("Connection1", , , _
        "ODBC;DSN=NOOSA;" & _
            "Database=" + db + ";" & _
            "UID=" + UID + ";" & _
            "PWD=" + PWD + ";")
    
        cDebug "ODBC;DSN=NOOSA;" & _
                "Database=" + db + ";" & _
                "UID=" + UID + ";" & _
                "PWD=" + "Hidden" + ";"
    Case True
        
        Set oConn = wrkODBC.OpenConnection("Connection1", , , _
        "ODBC;DSN=Testbench;" & _
            "Database=" + db + ";" & _
            "UID=" + UID + ";" & _
            "PWD=" + PWD + ";")
        
        cDebug "ODBC;DSN=Testbench;" & _
                "Database=" + db + ";" & _
                "UID=" + UID + ";" & _
                "PWD=" + "Hidden" + ";"
    
    End Select
                            
    DAOConn = True

    'oConn.Execute "alter table InvoiceOut add column (SubRecID Bigint default 0)"
    
Exit Function
ErrOcc:
    
    Debug.Print Err.Description
    
    MsgBox Err.Description, vbCritical, "Could not connect to server " & server
    
    DAOConn = False


End Function


Function ReplaceString(sMessage As Variant, sSearch As String, sReplace As String) As String

    Dim iPos As Variant
    
    Do
        iPos = 0
        If InStr(sMessage, sSearch) > 0 Then
            iPos = InStr(iPos + 1, sMessage, sSearch)
            If iPos = 1 Then
                sMessage = sReplace & Mid(sMessage, Len(sSearch) + 1)
            Else
                sMessage = Left(sMessage, iPos - 1) & sReplace & Mid(sMessage, iPos + Len(sSearch))
            End If
            While InStr(iPos + Len(sReplace) + 1, sMessage, sSearch) > 0
                iPos = InStr(iPos + Len(sReplace) + 1, sMessage, sSearch)
                If iPos = 1 Then
                    sMessage = sReplace & Mid(sMessage, Len(sSearch))
                Else
                    sMessage = Left(sMessage, iPos - 1) & sReplace & Mid(sMessage, iPos + Len(sReplace))
                End If
            Wend
        Else
            Exit Do
        End If
    Loop While sReplace = ""
    ReplaceString = sMessage
    
End Function

' This function replaces all the characters that need escape character witht the escape symbol
Function ESC(txt As Variant) As Variant

    ESC = txt
        
    ESC = Me.ReplaceString(ESC, "\", "\\") ' A backslash (\) character.
    ESC = Me.ReplaceString(ESC, Chr$(0), "\0") ' An ASCII 0 (NUL) character
    ESC = Me.ReplaceString(ESC, "'", "\'") ' A single quote (') character.
    ESC = Me.ReplaceString(ESC, Chr$(34), "\" + Chr$(34)) 'A double quote (") character.
    ESC = Me.ReplaceString(ESC, Chr$(8), "\b")
    ESC = Me.ReplaceString(ESC, vbLf, "\n") ' A newline character.
    ESC = Me.ReplaceString(ESC, vbCr, "\r") ' A carriage return character.
    ESC = Me.ReplaceString(ESC, vbTab, "\t") ' A tab character
    ESC = Me.ReplaceString(ESC, Chr$(26), "\z") ' ASCII(26) (Control-Z).  This character can be encoded to allow you to work around the problem that ASCII(26) stands for END-OF-FILE on Windows.  (ASCII(26) will cause problems if you try to use mysql database < filename.)
    ESC = Me.ReplaceString(ESC, "%", "\%")
    ESC = Me.ReplaceString(ESC, "_", "\_") ' A _ character. This is used to search for literal instances of _ in contexts where _ would otherwise be interpreted as a wildcard character. See String comparison functions.


End Function
Function Connection(Optional db As String = "", Optional server As String = "", Optional UID As String = "", Optional PWD As String = "", Optional ByRef oConn As ADODB.Connection, Optional ByRef lblAction As Label = Nothing) As Boolean
    
    If db = "" Then db = GetSetting(App.ProductName, "DB", "DB", "no1_billing")
    If server = "" Then server = GetSetting(App.ProductName, "DB", "Server", "noosa.no1.com.au")
    If UID = "" Then UID = GetSetting(App.ProductName, "DB", "UID", "noosa")
    If PWD = "" Then PWD = GetSetting(App.ProductName, "DB", "PWD", "wizard")
        
    On Error GoTo ErrOcc
    
    Set oConn = New ADODB.Connection

    If Not lblAction Is Nothing Then lblAction.Caption = "Opening ADO Connection - [" & server & "]"
    
    oConn.Open "Driver={Mysql ODBC 3.51 Driver};" & _
            "Database=" + db + ";" & _
            "UID=" + UID + ";" & _
            "PWD=" + PWD + ";" & _
            "PORT=3306;" & _
            "OemToAnsi=No;" & _
            "Server=" + server + ";OPTION=2083;"
    
    cDebug "Driver={MySQL ODBC 3.51 Driver};" & _
            "Database=" + db + ";" & _
            "UID=" + UID + ";" & _
            "PORT=3306;" & _
            "OemToAnsi=No;" & _
            "Server=" + server + ";OPTION=2083;"
                        
    Connection = True

    'oConn.Execute "alter table InvoiceOut add column (SubRecID Bigint default 0)"
Exit Function
ErrOcc:
    
    Debug.Print Err.Description
    
    MsgBox Err.Description, vbCritical, "Could not connect to server " & server
    
    Connection = False

End Function

Function OpenTableDef(oConn As DAO.Connection, rsDAO As DAO.Recordset, Optional TableName As String, Optional SQLQuery As String) As Boolean

    On Error GoTo ErrOcc
    Dim MySQL  As String
    
    If TableName <> "" Then
        MySQL = "Select * from " & TableName
    ElseIf SQLQuery <> "" Then
        MySQL = SQLQuery
    Else
        GoTo ErrOcc
    End If
   
    cDebug MySQL
    'If Not rsDAO Is Nothing Then
    '
    '    Select Case rsDAO.State
    '    Case adStateClosed  'Default. Indicates that the object is closed.
    '
    '    Case adStateOpen    'Indicates that the object is open.
    '        If Not rsDAO.EOF Then
    '            Select Case rsDAO.EditMode
    '            Case adEditNone 'Indicates that no editing operation is in progress.
    '            Case adEditInProgress 'Indicates that data in the current record has been modified but not yet saved.
    '                rsDAO.Update
    '            Case adEditAdd 'Indicates that the AddNew method has been invoked, and the current record in the copy buffer is a new record that hasn't been saved in the database.
    '                rsDAO.Update
    '            End Select
    '        End If
    '        rsDAO.Close
    '    End Select
    '
    '    Set rsDAO = Nothing
    '
    'End If
    
    Set rsDAO = oConn.OpenRecordset(MySQL, dbOpenDynaset)
    'rsDAO.Source = MySQL
    'Set rsDAO.ActiveConnection = oConn
    'rsDAO.CursorLocation = adUseClient
    'rsDAO.LockType = adLockOptimistic
    'rsDAO.CursorType = adOpenKeyset
    'rsDAO.Open
    'DoEvents
    OpenTableDef = True
    
Exit Function
ErrOcc:

Debug.Print Err.Description
Debug.Print MySQL
OpenTableDef = False

End Function

Function OpenTable(oConn As ADODB.Connection, rsADO As ADODB.Recordset, Optional TableName As String, Optional SQLQuery As String) As Boolean

    On Error Resume Next

    On Error GoTo ErrOcc
        
    Dim sql As String
    
    If TableName <> "" Then
        sql = "Select * from " & TableName
    ElseIf SQLQuery <> "" Then
        sql = SQLQuery
    Else
        GoTo ErrOcc
    End If
   
    If Not rsADO Is Nothing Then
        
        Select Case rsADO.State
        Case adStateClosed  'Default. Indicates that the object is closed.
        
        Case adStateOpen    'Indicates that the object is open.
            If Not rsADO.EOF Then
                Select Case rsADO.EditMode
                Case adEditNone 'Indicates that no editing operation is in progress.
                Case adEditInProgress 'Indicates that data in the current record has been modified but not yet saved.
                    rsADO.Update
                Case adEditAdd 'Indicates that the AddNew method has been invoked, and the current record in the copy buffer is a new record that hasn't been saved in the database.
                    rsADO.Update
                End Select
            End If
            rsADO.Close
        End Select
        
        Set rsADO = Nothing
        
    End If
    
    Set rsADO = New ADODB.Recordset
    rsADO.Source = sql
    Set rsADO.ActiveConnection = oConn
    rsADO.CursorLocation = adUseClient
    rsADO.LockType = adLockOptimistic
    rsADO.CursorType = adOpenKeyset
    rsADO.Open
    DoEvents
    OpenTable = True
    
    cDebug sql
    
Exit Function
ErrOcc:

Debug.Print sql
Debug.Print Err.Description
OpenTable = False
    
End Function

Public Sub Execute(oConn As ADODB.Connection, sql As String)

    On Error Resume Next
    
           
    oConn.Execute sql
    cDebug sql
    
End Sub

Public Function SetColumnHeaders(lv As ListView, ServiceKey As String, Optional oConn As ADODB.Connection) As Boolean

    'On Error Resume Next
    
    lv.ColumnHeaders.Clear
    
    Dim rsLoad As ADODB.Recordset
    Dim bResult As Boolean
    Dim colHeader As ColumnHeader
    
    bResult = MySQL.OpenTable(oConn, rsLoad, , "Select * from ColumnLayout where ServiceKey = '" & ServiceKey & "'")
    
    If rsLoad.RecordCount > 0 Then
        While Not rsLoad.EOF
            Set colHeader = lv.ColumnHeaders.Add(, , rsLoad!Description, rsLoad!Width)
            colHeader.Tag = rsLoad!Fieldname & "<" & IIf(IsNull(rsLoad!sFormat), "No Format", rsLoad!sFormat)
            rsLoad.MoveNext
        Wend
    End If
    
    lv.Refresh
    DoEvents
    SetColumnHeaders = True
    
End Function

Public Function OP(ByRef oRS As ADODB.Recordset, Fieldname As String) As Variant

    If InStr(Fieldname, "iTax") > 0 Then
        Select Case Mid(Fieldname, 5, 1)
        Case "/"
            OP = oRS(Mid(Fieldname, 6)) * sTax
        Case "+"
            OP = oRS(Mid(Fieldname, 6)) * sTax + oRS(Mid(Fieldname, 6))
        Case Else
            Stop
        End Select
    Else
        OP = oRS(Fieldname)
    End If

End Function

Public Function VirtualISP(ByVal sSQL As String, Optional TblName As String) As String
    
    Select Case Login.lVirtualID
    Case 0
        VirtualISP = sSQL
    Case Else
        If InStr(LCase(sSQL), "order by") > 0 Then
            If InStr(LCase(sSQL), "where") > 0 Then
                VirtualISP = Trim(Left(sSQL, InStr(LCase(sSQL), "order by") - 1)) & " AND " & IIf(TblName <> "", TblName + ".", "") + "VirtualID = " & Login.lVirtualID & " " & Trim(Mid(sSQL, InStr(LCase(sSQL), "order by") - 1))
            Else
                VirtualISP = Trim(Left(sSQL, InStr(LCase(sSQL), "order by") - 1)) & " WHERE " & IIf(TblName <> "", TblName + ".", "") + "VirtualID = " & Login.lVirtualID & " " & Trim(Mid(sSQL, InStr(LCase(sSQL), "order by") - 1))
            End If
        ElseIf InStr(LCase(sSQL), "limit") > 0 Then
            If InStr(LCase(sSQL), "where") > 0 Then
                VirtualISP = Trim(Left(sSQL, InStr(LCase(sSQL), "limit") - 1)) & " AND " & IIf(TblName <> "", TblName + ".", "") + "VirtualID = " & Login.lVirtualID & " " & Trim(Mid(sSQL, InStr(LCase(sSQL), "limit") - 1))
            Else
                VirtualISP = Trim(Left(sSQL, InStr(LCase(sSQL), "limit") - 1)) & " WHERE " & IIf(TblName <> "", TblName + ".", "") + "VirtualID = " & Login.lVirtualID & " " & Trim(Mid(sSQL, InStr(LCase(sSQL), "limit") - 1))
            End If
        Else
            If InStr(LCase(sSQL), "where") > 0 Then
                VirtualISP = sSQL & " AND " & IIf(TblName <> "", TblName + ".", "") + "VirtualID = " & Login.lVirtualID
            Else
                VirtualISP = sSQL & " WHERE " & IIf(TblName <> "", TblName + ".", "") + "VirtualID = " & Login.lVirtualID
            End If
        End If
    End Select

End Function
Public Function ProcessBill2Now()

End Function

Public Function SetNextCycle(rs As ADODB.Recordset, sType As String, Interval As Integer)

    rs!PreviousCycle = rs!NextCycle
    rs!NextCycle = DateAdd(sType, Interval, Now)
    
End Function

Public Function SetRecID(rs As ADODB.Recordset, TblName As String, ADOConn As ADODB.Connection) As Variant
    
    '9.22337203685478E+18
    Randomize Now
    
    On Error Resume Next
    
    Dim rsCount As ADODB.Recordset
    Dim bResult As Boolean
           
    If IsNull(rs!RecID) Or TblName = "PlanTypes" Then
        
        SetRecID = Me.GetTMPRecID(TblName, ADOConn)
        rs!RecID = SetRecID
        rs.Update

        While Err.Number <> 0
            Err.Clear
            Randomize Now
            SetRecID = SetRecID + 1
            rs!RecID = SetRecID
            rs.Update
            cDebug Err.Description
            DoEvents
            If Err.Number > 0 Then cDebug Err.Description
        Wend
    Else
        SetRecID = rs!RecID
    End If
    
End Function

Public Function GetTMPRecID(TableName As String, ADOConn As ADODB.Connection) As Long

    On Error Resume Next
    
    Dim rsRandom As ADODB.Recordset
    Dim rsMAX As ADODB.Recordset
    
    Dim bResult As Boolean
    Do
    
        Randomize Now
        bResult = Me.OpenTable(ADOConn, rsMAX, , "Select MAX(RecID) as maxRecID from RecIDPlacement where TableName = '" + TableName + "'")
        
        
        If bResult = True Then
            If rsMAX.RecordCount > 0 Then
                GetTMPRecID = IIf(IsNull(rsMAX!maxrecid), 0, rsMAX!maxrecid)
            Else
                GetTMPRecID = 0
            End If
        Else
            GetTMPRecID = 0
        End If
        
        
        If GetTMPRecID = 0 Then
            bResult = Me.OpenTable(ADOConn, rsMAX, , "Select MAX(RecID) as maxRecID from " + TableName + "")
            If rsMAX.RecordCount > 0 Then
                If GetTMPRecID <= rsMAX!maxrecid Then GetTMPRecID = IIf(IsNull(rsMAX!maxrecid), 0, rsMAX!maxrecid)
            End If
        End If
        
            
        Do
            GetTMPRecID = GetTMPRecID + 1
            bResult = Me.OpenTable(ADOConn, rsRandom, , "Select * from RecIDPlacement where RecID = " & GetTMPRecID & " AND TableName = '" & TableName & "' Limit 1")
            DoEvents
        Loop Until rsRandom.RecordCount = 0
        
        rsRandom.AddNew
        rsRandom!RecID = GetTMPRecID
        rsRandom!TableName = TableName
        rsRandom.Update
        rsRandom.Close
            
        'mysql.execute adoconn,  "Delete from RecIDPlacement Where RecID < " & GetTMPRecID
        
        If TableName = "cctransactions" Then
            On Error Resume Next
            Err.Clear
            MySQL.Execute ADOConn, "Insert into " & TableName & " (RecID) Values (" & GetTMPRecID & ")"
            
            If Err.Number = 0 Then Exit Do
        Else
            Exit Do
        End If
    Loop
    
End Function

Public Function SetInvoiceSerial(Optional ADOConn As ADODB.Connection) As Variant
    
    Dim bResult As Boolean
    Dim rsLoad As ADODB.Recordset
    Dim InvSerialNumber As Variant
    
    bResult = MySQL.OpenTable(ADOConn, rsLoad, , "Select RecID from InvoiceTraxr order by RecID DESC Limit 1")
    If rsLoad.RecordCount > 0 Then
        InvSerialNumber = rsLoad!RecID
    Else
        InvSerialNumber = 10000
    End If
    
    Do
        InvSerialNumber = InvSerialNumber + 1
        bResult = MySQL.OpenTable(ADOConn, rsLoad, , "Select * from InvoiceTraxr Where InvoiceSerial = '" & Hex(InvSerialNumber) & "'")
    Loop Until rsLoad.RecordCount = 0
    
    SetInvoiceSerial = InvSerialNumber
    
End Function


Public Function Update(rs As ADODB.Recordset, TblName As String, PointerField As String) As Boolean

End Function
Public Function NumDecrypt(txt As String) As String

    Dim sNew As String
    
    For X = 1 To Len(txt)
        Select Case Mid(txt, X, 1)
        Case "9"
            sNew = sNew + "4"
        Case "7"
            sNew = sNew + "3"
        Case "5"
            sNew = sNew + "2"
        Case "3"
            sNew = sNew + "1"
        Case "1"
            sNew = sNew + "0"
        Case "8"
            sNew = sNew + "9"
        Case "6"
            sNew = sNew + "8"
        Case "4"
            sNew = sNew + "7"
        Case "2"
            sNew = sNew + "6"
        Case "0"
            sNew = sNew + "5"
        End Select
    Next
    
    NumDecrypt = sNew
    
End Function
Public Function NumCrypt(txt As String) As String

    Dim sNew As String
    
    For X = 1 To Len(txt)
        Select Case Mid(txt, X, 1)
        Case "0"
            sNew = sNew + "1"
        Case "1"
            sNew = sNew + "3"
        Case "2"
            sNew = sNew + "5"
        Case "3"
            sNew = sNew + "7"
        Case "4"
            sNew = sNew + "9"
        Case "5"
            sNew = sNew + "0"
        Case "6"
            sNew = sNew + "2"
        Case "7"
            sNew = sNew + "4"
        Case "8"
            sNew = sNew + "6"
        Case "9"
            sNew = sNew + "8"
        End Select
    Next
    
    NumCrypt = sNew
    
End Function

Public Function AddReceiptItem(oConn As ADODB.Connection, Optional acciRecID As Variant = 0, Optional InvoiceOutID As Variant = 0, Optional InvoiceInID As Variant = 0, Optional acciServicesID As Variant = 0, Optional RefundID As Variant = 0, Optional Paid As Currency = 0, Optional Refunded As Currency = 0, Optional PaymentType As String) As Variant

    Dim RecID As Variant
    
    Do
        Err.Clear
        RecID = Me.GetTMPRecID("Receipts", oConn)
        Me.Execute oConn, "insert into Receipts (RecID, acci_RecID, InvoiceOutID, InvoiceInID, RefundID, acciServicesID, Paid, Refunded, PaymentType) VALUES (" & RecID & "," & acciRecID & "," & InvoiceOutID & "," & InvoiceInID & "," & RefundID & "," & acciServicesID & "," & Paid & "," & Refunded & ",'" & PaymentType & "')"
        If Err.Number <> 0 Then cDebug Err.Description
    Loop Until Err.Number = 0
    
    AddReceiptItem = RecID
    
End Function
