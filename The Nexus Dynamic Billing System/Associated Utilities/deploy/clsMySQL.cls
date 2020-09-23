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
Attribute VB_Ext_KEY = "Member0" ,"colOPCache"
Private mvarcolOPCache As colOPCache


Public Property Get colOPCache() As colOPCache
    If mvarcolOPCache Is Nothing Then
        Set mvarcolOPCache = New colOPCache
    End If


    Set colOPCache = mvarcolOPCache
End Property


Public Property Set colOPCache(vData As colOPCache)
    Set mvarcolOPCache = vData
End Property
Private Sub Class_Terminate()
    Set mvarcolOPCache = Nothing
End Sub


Public Function fldConst(Fldtxt As String, tFlag As Byte) As String

    Dim iPos As Integer
    Dim iPosb As Integer
    Dim iField As String
    Dim iFormat As String

    iPos = 0
    iPosb = 0
    iPos = InStr(iPosb + 1, Fldtxt, "!")
    iPosb = InStr(iPos + 1, Fldtxt, "!")
    If iPos = 0 And iPosb = 0 Then
        iField = Fldtxt
        iFormat = Fldtxt
    Else
    
        iField = Mid(Fldtxt, iPos + 1, iPosb - iPos - 1)
        If Mid(Fldtxt, iPosb + 1, 1) = "^" Then
            iPos = InStr(iPosb + 1, Fldtxt, "^")
            iPosb = InStr(iPos + 1, Fldtxt, "^")
            iFormat = Mid(Fldtxt, iPos + 1, iPosb - iPos - 1)
        End If
    
    End If
    
    Select Case Flag
    Case 0
        fldConst = iField
    Case Else
        fldConst = iFormat
    End Select
    
End Function

Public Function fldSupportCipher(TypeID As Integer) As Boolean

    '*[ Error Checking Variables ]**********************************************************************************
    Const RoutineName = "Main"
    Const ContainerName = "Globals"
    '***************************************************************************************************************


    '
    '***********************************************************************************************
    '**  Project Alpha ® 2003, 2004 +                                                             **
    '***********************************************************************************************
    '**  This code is not to be distributed, reverse engineered or simulated in any way without   **
    '**  Premission from the author. The authors of this code is as follows: Simon Antony Roberts **                                                     **
    '**Jarrett Cliff Costi, these two are the only people you can communicate with about this code**
    '***********************************************************************************************
    '**  Project Alpha is a product of Exitstencil Press Australia                                **
    '***********************************************************************************************
    '**                                                                                           **
    '**  Routine:                                                                                 **
    '**  Arguments:                                                                               **
    '**  Description:    Subroutine, Function or Property of project alpha                        **
    '**  Author:         Simon Roberts                                                            **
    '**  Date Last mod:  19-01-2004                                                               **
    '**                                                                                           **
    '********************************************** Copyright © 2004 Exitstencil Press Australia ***
    '
    '
    '
    If bDebug = -1 Then
        On Error GoTo 0
    ElseIf bDebug = 1 Then
        On Error Resume Next
    Else
        On Error GoTo ErrorOccur
    End If
    fldSupportCipher = False
    
    Select Case TypeID
    Case adBSTR '8 This data type indicates a null-terminated Unicode character string (DBTYPE_BSTR). This data type is not supported by the OLE DB Provider.
        fldSupportCipher = True
    Case adVariant '12 This data type indicates an Automation variant (DBTYPE_VARIANT). This data type is not supported by the OLE DB Provider.
        fldSupportCipher = True
    Case adChar '129 This data type indicates a character string value (DBTYPE_STR).
        fldSupportCipher = True
    Case adWChar '130 This data type indicates a null-terminated Unicode character string (DBTYPE_WSTR). This data type is not supported by the OLE DB Provider.
        fldSupportCipher = True
    Case adVarChar '200 This data type indicates variable-length character data (DBTYPE_STR).
        fldSupportCipher = True
    Case adLongVarChar '201 This data type indicates a long string value.
        fldSupportCipher = True
    Case adVarWChar '202 This data type indicates a Unicode string value. This data type is not supported by the OLE DB Provider.
        fldSupportCipher = True
    Case adLongVarWChar '203 This data type indicates a long Unicode string value. This data type is not supported by the OLE DB Provider.
        fldSupportCipher = True
    End Select


Exit Function



ErrorOccur:
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function

Public Function fldType(TypeID As Integer, Optional ByRef bNumeric As Boolean) As String

    '*[ Error Checking Variables ]**********************************************************************************
    Const RoutineName = "Main"
    Const ContainerName = "Globals"
    '***************************************************************************************************************


    '
    '***********************************************************************************************
    '**  Project Alpha ® 2003, 2004 +                                                             **
    '***********************************************************************************************
    '**  This code is not to be distributed, reverse engineered or simulated in any way without   **
    '**  Premission from the author. The authors of this code is as follows: Simon Antony Roberts **                                                     **
    '**Jarrett Cliff Costi, these two are the only people you can communicate with about this code**
    '***********************************************************************************************
    '**  Project Alpha is a product of Exitstencil Press Australia                                **
    '***********************************************************************************************
    '**                                                                                           **
    '**  Routine:                                                                                 **
    '**  Arguments:                                                                               **
    '**  Description:    Subroutine, Function or Property of project alpha                        **
    '**  Author:         Simon Roberts                                                            **
    '**  Date Last mod:  19-01-2004                                                               **
    '**                                                                                           **
    '********************************************** Copyright © 2004 Exitstencil Press Australia ***
    '
    '
    '
    If bDebug = -1 Then
        On Error GoTo 0
    ElseIf bDebug = 1 Then
        On Error Resume Next
    Else
        On Error GoTo ErrorOccur
    End If

    Select Case TypeID
    Case adEmpty '0 This data type indicates that no value was specified (DBTYPE_EMPTY).
        fldType = "Empty"
    Case adSmallInt '2 This data type indicates a 2-byte (16-bit) signed integer (DBTYPE_I2).
        fldType = "Small Integer"
        bNumeric = True
    Case adInteger '3 This data type indicates a 4-byte (32-bit) signed integer (DBTYPE_I4).
        fldType = "Integer"
        bNumeric = True
    Case adSingle '4 This data type indicates a 4-byte (32-bit) single-precision IEEE floating-point number (DBTYPE_R4).
        fldType = "Single"
        bNumeric = True
    Case adDouble '5 This data type indicates an 8-byte (64-bit) double-precision IEEE floating-point number (DBTYPE_R8).
        fldType = "Double"
        bNumeric = True
    Case adCurrency '6 A data type indicates a currency value (DBTYPE_CY). Currency is a fixed-point number with 4 digits to the right of the decimal point. It is stored in an 8-byte signed integer scaled by 10,000. This data type is not supported by the OLE DB Provider.
        fldType = "Currency"
        bNumeric = True
    Case adDate '7 This data type indicates a date value stored as a Double, the whole part of which is the number of days since December 30, 1899, and the fractional part of which is the fraction of a day. This data type is not supported by the OLE DB Provider.
        fldType = "Date/Time"
        bNumeric = False
    Case adBSTR '8 This data type indicates a null-terminated Unicode character string (DBTYPE_BSTR). This data type is not supported by the OLE DB Provider.
        fldType = "String"
        bNumeric = False
    Case adIDispatch '9 This data type indicates a pointer to an IDispatch interface on an OLE object (DBTYPE_IDISPATCH). This data type is not supported by the OLE DB Provider.
        fldType = "Dispact"
        bNumeric = False
    Case adError '10 This data type indicates a 32-bit error code (DBTYPE_ERROR). This data type is not supported by the OLE DB Provider.
        fldType = "Error"
        bNumeric = False
    Case adBoolean '11 This data type indicates a Boolean value (DBTYPE_BOOL). This data type is not supported by the OLE DB Provider.
        fldType = "Boolean"
        bNumeric = False
    Case adVariant '12 This data type indicates an Automation variant (DBTYPE_VARIANT). This data type is not supported by the OLE DB Provider.
        fldType = "Variant"
        bNumeric = False
    Case adIUnknown '13 This data type indicates a pointer to an IUnknown interface on an OLE object (DBTYPE_IUNKNOWN). This data type is not supported by the OLE DB Provider.
        fldType = "Unknown"
        bNumeric = False
    Case adDecimal '14 This data type indicates numeric data with a fixed precision and scale (DBTYPE_DECIMAL).
        fldType = "Decimal"
        bNumeric = True
    Case adTinyInt '16 This data type indicates a single -byte (8-bit) signed integer (DBTYPE_I1). This data type is not supported by the OLE DB Provider.
        fldType = "Tiny Integer"
        bNumeric = True
    Case adUnsignedTinyInt '17 This data type indicates a single-byte (8-bit) unsigned integer (DBTYPE_UI1). This data type is not supported by the OLE DB Provider.
        fldType = "Unsigned Tiny Integer"
        bNumeric = True
    Case adUnsignedSmallInt '18 This data type indicates a 2-byte (16-bit) unsigned integer (DBTYPE_UI2). This data type is not supported by the OLE DB Provider.
        fldType = "Unsigned Small Integer"
        bNumeric = True
    Case adUnsignedInt '19 This data type indicates a 4-byte (32-bit) unsigned integer (DBTYPE_UI4). This data type is not supported by the OLE DB Provider.
        fldType = "Unsigned Integer"
        bNumeric = True
    Case adBigInt '20 This data type indicates an 8-byte (64-bit) signed integer (DBTYPE_I8). This data type is not supported by the OLE DB Provider.
        fldType = "Big Integer"
        bNumeric = True
    Case adUnsignedBigInt '21 This data type indicates an 8-byte (64-bit) unsigned integer (DBTYPE_UI8). This data type is not supported by the OLE DB Provider.
        fldType = "Unsigned Big Integer"
        bNumeric = True
    Case adGUID '72 This data type indicates a globally unique identifier or GUID (DBTYPE_GUID). This data type is not supported by the OLE DB Provider.
        fldType = "GUID"
        bNumeric = False
    Case adBinary '128 This data type indicates fixed-length binary data (DBTYPE_BYTES).
        fldType = "Binary"
        bNumeric = False
    Case adChar '129 This data type indicates a character string value (DBTYPE_STR).
        fldType = "Char"
        bNumeric = False
    Case adWChar '130 This data type indicates a null-terminated Unicode character string (DBTYPE_WSTR). This data type is not supported by the OLE DB Provider.
        fldType = "WChar"
        bNumeric = False
    Case adNumeric '131 This data type indicates numeric data where the precision and scale are exactly as specified (DBTYPE_NUMERIC). .
        fldType = "Numeric"
        bNumeric = True
    Case adUserDefined '132 This data type indicates user-defined data (DBTYPE_UDT). This data type is not supported by the OLE DB Provider.
        fldType = "User Defined"
        bNumeric = False
    Case adDBDate '133 This data type indicates a OLE DB date structure (DBTYPE_DATE).
        fldType = "Date"
        bNumeric = False
    Case adDBTime '134 This data type indicates a OLE DB time structure (DBTYPE_TIME).
        fldType = "Time"
        bNumeric = False
    Case adDBTimeStamp '135 This data type indicates a OLE DB timestamp structure (DBTYPE_TIMESTAMP).
        fldType = "Time Stamp"
        bNumeric = False
    Case adVarChar '200 This data type indicates variable-length character data (DBTYPE_STR).
        fldType = "Var Char"
        bNumeric = False
    Case adLongVarChar '201 This data type indicates a long string value.
        fldType = "Long Var Char"
        bNumeric = False
    Case adVarWChar '202 This data type indicates a Unicode string value. This data type is not supported by the OLE DB Provider.
        fldType = "VarWChar"
        bNumeric = False
    Case adLongVarWChar '203 This data type indicates a long Unicode string value. This data type is not supported by the OLE DB Provider.
        fldType = "Long VarWChar"
        bNumeric = False
    Case adVarBinary '204 This data type indicates variable-length binary data (DBTYPE_BYTES).
        fldType = "Var Binary"
        bNumeric = False
    Case adLongVarBinary '205 This data type indicates a long binary value.
        fldType = "Long Var Binary"
        bNumeric = False
    Case Else
        fldType = "Unknown"
        bNumeric = False
    End Select


Exit Function



ErrorOccur:
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function

Public Function ReplaceString(sMessage As Variant, sSearch As String, sReplace As String) As String


    '*[ Error Checking Variables ]**********************************************************************************
    Const RoutineName = "clsMySQL"
    Const ContainerName = "ReplaceString"
    '***************************************************************************************************************


    '
    '***********************************************************************************************
    '**  Project Alpha ® 2003, 2004 +                                                             **
    '***********************************************************************************************
    '**  This code is not to be distributed, reverse engineered or simulated in any way without   **
    '**  Premission from the author. The authors of this code is as follows: Simon Antony Roberts **                                                     **
    '**Jarrett Cliff Costi, these two are the only people you can communicate with about this code**
    '***********************************************************************************************
    '**  Project Alpha is a product of Exitstencil Press Australia                                **
    '***********************************************************************************************
    '**                                                                                           **
    '**  Routine:                                                                                 **
    '**  Arguments:                                                                               **
    '**  Description:    Subroutine, Function or Property of project alpha                        **
    '**  Author:         Simon Roberts                                                            **
    '**  Date Last mod:  19-01-2004                                                               **
    '**                                                                                           **
    '********************************************** Copyright © 2004 Exitstencil Press Australia ***
    '
    '
    '
    If bDebug = -1 Then
        On Error GoTo 0
    ElseIf bDebug = 1 Then
        On Error Resume Next
    Else
        On Error GoTo ErrorOccur
    End If
    
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
                    sMessage = Left(sMessage, iPos - 1) & sReplace & Mid(sMessage, iPos + 1)
                End If
            Wend
        Else
            Exit Do
        End If
    Loop While sReplace = ""
    ReplaceString = sMessage
    
Exit Function



ErrorOccur:
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

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
    ESC = Me.ReplaceString(ESC, Chr$(26), "\z") ' ASCII(26) (Control-Z).  This character can be AES_ENCRYPTd to allow you to work around the problem that ASCII(26) stands for END-OF-FILE on Windows.  (ASCII(26) will cause problems if you try to use mysql database < filename.)
    'ESC = Me.ReplaceString(ESC, "%", "\%")
    'ESC = Me.ReplaceString(ESC, "_", "\_") ' A _ character. This is used to search for literal instances of _ in contexts where _ would otherwise be interpreted as a wildcard character. See String comparison functions.


Exit Function



ErrorOccur:
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function
Function Connection(Optional db As String = "", Optional server As String = "", Optional UID As String = "", Optional PWD As String = "", Optional ByRef oConn As ADODB.Connection, Optional ByRef lblAction As Label = Nothing) As Boolean
    
    On Error Resume Next
    
    If db = "" Then db = GetSetting(App.ProductName, "DB", "DB", "projectalpha")
    If server = "" Then server = GetSetting(App.ProductName, "DB", "Server", "demon.comcen.com.au")
    If UID = "" Then UID = GetSetting(App.ProductName, "DB", "UID", "sroberts")
    If PWD = "" Then PWD = GetSetting(App.ProductName, "DB", "PWD", "kl7jb0lsf")
        
    If bDebug = True Then On Error GoTo 0 Else On Error Resume Next
    
ReDo:
    
    Set oConn = New ADODB.Connection

    If Not lblAction Is Nothing Then
        lblAction.Caption = "Opening ADO Connection - [" & server & "]"
        lblAction.Refresh
        DoEvents
    End If
    
    oConn.Open "Driver={Mysql ODBC 3.51 Driver};" & _
            "Database=" + db + ";" & _
            "UID=" + UID + ";" & _
            "PWD=" + PWD + ";" & _
            "PORT=3306;" & _
            "OemToAnsi=No;" & _
            "SSL=Yes;" & _
            "Server=" + server + ";OPTION=2084;"
    'Clipboard.Clear
    'Clipboard.SetText "Driver={Mysql ODBC 3.51 Driver};" & _
            "Database=" + db + ";" & _
            "UID=" + UID + ";" & _
            "PWD=" + PWD + ";" & _
            "PORT=3306;" & _
            "OemToAnsi=No;" & _
            "SSL=Yes;" & _
            "Server=" + server + ";OPTION=2084;"
    'Stop
    
    'cDebug "Driver={MySQL ODBC 3.51 Driver};" & _
            "Database=" + db + ";" & _
            "UID=" + UID + ";" & _
            "PORT=3306;" & _
            "OemToAnsi=No;" & _
            "Server=" + server + ";OPTION=2083;"
                        
    Connection = True

    'oConn.Execute "alter table invoiceout add column (SubRecID Bigint default 0)"
If Err.Number = 0 Then Exit Function
    
    Static Count As Long
    Count = Count + 1
    If Count < 25 Then GoTo ReDo
     
    cDebug Err.Description
    
    'MsgBox Err.Description, vbCritical, "Could not connect to server " & server
    
    Connection = False

Exit Function



ErrorOccur:
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function


Public Function OpenTable(oConn As ADODB.Connection, rsADO As ADODB.Recordset, Optional Tablename As String, Optional SQLQuery As Variant) As Boolean


    '*[ Error Checking Variables ]**********************************************************************************
    Const RoutineName = "Main"
    Const ContainerName = "Globals"
    '***************************************************************************************************************


    '
    '***********************************************************************************************
    '**  Project Alpha ® 2003, 2004 +                                                             **
    '***********************************************************************************************
    '**  This code is not to be distributed, reverse engineered or simulated in any way without   **
    '**  Premission from the author. The authors of this code is as follows: Simon Antony Roberts **                                                     **
    '**Jarrett Cliff Costi, these two are the only people you can communicate with about this code**
    '***********************************************************************************************
    '**  Project Alpha is a product of Exitstencil Press Australia                                **
    '***********************************************************************************************
    '**                                                                                           **
    '**  Routine:                                                                                 **
    '**  Arguments:                                                                               **
    '**  Description:    Subroutine, Function or Property of project alpha                        **
    '**  Author:         Simon Roberts                                                            **
    '**  Date Last mod:  19-01-2004                                                               **
    '**                                                                                           **
    '********************************************** Copyright © 2004 Exitstencil Press Australia ***
    '
    '
    '
    
    On Error Resume Next
    
    
        
    Dim sql As Variant
    
    If Tablename <> "" Then
        sql = "select * from " & Tablename
    ElseIf SQLQuery <> "" Then
        sql = SQLQuery
    Else
        GoTo ErrorOccur
    End If
   
'    Debug.Print sql
    
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
    
    If SQLQuery = "" Then Exit Function
    
    Set rsADO = New ADODB.Recordset
    rsADO.Source = SQLQuery
    Set rsADO.ActiveConnection = oConn
    rsADO.CursorLocation = adUseClient
    rsADO.LockType = adLockOptimistic
    rsADO.CursorType = adOpenDynamic
    rsADO.Open
    DoEvents
    OpenTable = True
    DoEvents
    OpenTable = True
'    MsgBox SQLQuery
    
    cDebug "ADO:// [" & Err.Number & "] - " & sql
    
    'Clipboard.Clear
    'Clipboard.SetText sql
    'Stop
    
Exit Function

Exit Function



ErrorOccur:
OpenTable = False
Dim floop As Integer

floop = floop + 1
If floop = 5 Then Exit Function
cDebug "ERR:// [" & Err.Number & "] - " & sql

Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit

Case vbResumeNext
    Resume Next
End Select
    
    
End Function

Public Function Execute(oConn As ADODB.Connection, sql As Variant, Optional bHidden As Boolean)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Execute"
    Const ContainerName = "clsMySQL"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
'**  This code is not to be distributed, reverse engineered or simulated in any way without   **
'**  Premission from the author. The authors of this code is as follows: Simon Antony Roberts **                                                     **
'**Jarrett Cliff Costi, these two are the only people you can communicate with about this code**
'***********************************************************************************************
'**  Project Alpha is a product of Exitstencil Press Australia                                **
'***********************************************************************************************
'**                                                                                           **
'**  Routine:                                                                                 **
'**  Arguments:                                                                               **
'**  Description:    Subroutine, Function or Property of project alpha                        **
'**  Author:         Simon Roberts                                                            **
'**  Date Last mod:  19-01-2004                                                               **
'**                                                                                           **
'********************************************** Copyright © 2004 Exitstencil Press Australia ***
'
'
'
    If bDebug = -1 Then
        On Error GoTo 0
    ElseIf bDebug = 1 Then
        On Error Resume Next
    Else
        On Error GoTo ErrorOccur
    End If


     On Error Resume Next
    
    
    
                
    
 
           
    oConn.Execute sql

    
    If bHidden = False Then cDebug "ADO:// [" & Err.Number & "] " & sql
    
'    frmAgent.oChar.Stop
    
   
Exit Function



ErrorOccur:

    frmAgent.oChar.Stop
    frmAgent.oChar.Speak "Oh an error has occure while executing a SQL Question inside the server." & vbCrLf & " Error Number [" & Val(Err.Number) & "] " & Err.Description & "has occured within the routine " & RoutineName & " inside container " & ContainerName
    
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function

Public Function NumDecrypt(txt As String) As String


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "NumDecrypt"
    Const ContainerName = "clsMySQL"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
'**  This code is not to be distributed, reverse engineered or simulated in any way without   **
'**  Premission from the author. The authors of this code is as follows: Simon Antony Roberts **                                                     **
'**Jarrett Cliff Costi, these two are the only people you can communicate with about this code**
'***********************************************************************************************
'**  Project Alpha is a product of Exitstencil Press Australia                                **
'***********************************************************************************************
'**                                                                                           **
'**  Routine:                                                                                 **
'**  Arguments:                                                                               **
'**  Description:    Subroutine, Function or Property of project alpha                        **
'**  Author:         Simon Roberts                                                            **
'**  Date Last mod:  19-01-2004                                                               **
'**                                                                                           **
'********************************************** Copyright © 2004 Exitstencil Press Australia ***
'
'
'
    If bDebug = -1 Then
        On Error GoTo 0
    ElseIf bDebug = 1 Then
        On Error Resume Next
    Else
        On Error GoTo ErrorOccur
    End If


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
    
Exit Function



ErrorOccur:
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function
Public Function NumCrypt(txt As String) As String


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "NumCrypt"
    Const ContainerName = "clsMySQL"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
'**  This code is not to be distributed, reverse engineered or simulated in any way without   **
'**  Premission from the author. The authors of this code is as follows: Simon Antony Roberts **                                                     **
'**Jarrett Cliff Costi, these two are the only people you can communicate with about this code**
'***********************************************************************************************
'**  Project Alpha is a product of Exitstencil Press Australia                                **
'***********************************************************************************************
'**                                                                                           **
'**  Routine:                                                                                 **
'**  Arguments:                                                                               **
'**  Description:    Subroutine, Function or Property of project alpha                        **
'**  Author:         Simon Roberts                                                            **
'**  Date Last mod:  19-01-2004                                                               **
'**                                                                                           **
'********************************************** Copyright © 2004 Exitstencil Press Australia ***
'
'
'
    If bDebug = -1 Then
        On Error GoTo 0
    ElseIf bDebug = 1 Then
        On Error Resume Next
    Else
        On Error GoTo ErrorOccur
    End If


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
    
Exit Function



ErrorOccur:
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function
Public Function CodeME(txt As String) As String

    If InStr(txt, " ") > 0 Or InStr(txt, "-") > 0 Then
        Dim iCnt As Integer
        Dim NextChar As Boolean
                    
            iCnt = 1
            Do
                
                
                If iCnt = 1 Then
                    CodeME = CodeME + Left(txt, 3)
                ElseIf Mid(txt, iCnt, 1) = " " Or Mid(txt, iCnt, 1) = "-" Then
                    NextChar = True
                End If

                iCnt = iCnt + 1
                
                If NextChar = True And Mid(txt, iCnt, 1) <> " " And Mid(txt, iCnt, 1) <> "-" Then
                    CodeME = CodeME + Mid(txt, iCnt, 1)
                End If
                
            Loop Until Len(CodeME) = 7 Or InStr(iCnt, txt, " ") = 0
    
    Else
        CodeME = UCase(Left(txt, 3))
    End If
    
    CodeME = UCase(Trim(CodeME))
    CodeME = Me.ReplaceString(CodeME, " ", "")
    CodeME = Me.ReplaceString(CodeME, "-", "")
    
End Function
