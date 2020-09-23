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

Public Function fillLV(oConn As ADODB.Connection, rsIn As ADODB.Recordset, lv As ListView, Optional DoNotClear As Boolean, Optional itmX As ListItem)



    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "fillLV"
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
    If DoNotClear = False Then lv.ListItems.Clear
    
    Dim rsLoad As ADODB.Recordset
    Dim SQL As String
    Dim sResult As String
    Dim bNumeric As Boolean
    
    Dim iPos As Integer
    Dim iPosb As Integer
        
    
        Dim iField As String
        Dim iFormat As String
        Dim bSetAll  As Boolean
        
        If itmX Is Nothing Then bSetAll = True
        
        Do
            If bSetAll = True Then Set itmX = lv.ListItems.Add(, "r" & rsIn!RecID, "")
            For X = 1 To lv.ColumnHeaders.Count
                If lv.ColumnHeaders(X).Tag <> "" Then
                    
                    If InStr(lv.ColumnHeaders(X).Tag, "!") > 0 Then
                        
                        iPos = 0
                        iPosb = 0
                        iPos = InStr(iPosb + 1, lv.ColumnHeaders(X).Tag, "!")
                        iPosb = InStr(iPos + 1, lv.ColumnHeaders(X).Tag, "!")
                        iField = Mid(lv.ColumnHeaders(X).Tag, iPos + 1, iPosb - iPos - 1)
                        If Mid(lv.ColumnHeaders(X).Tag, iPosb + 1, 1) = "^" Then
                            iPos = InStr(iPosb + 1, lv.ColumnHeaders(X).Tag, "^")
                            iPosb = InStr(iPos + 1, lv.ColumnHeaders(X).Tag, "^")
                            iFormat = Mid(lv.ColumnHeaders(X).Tag, iPos + 1, iPosb - iPos - 1)
                        End If
                        
                        If InStr(LCase(lv.ColumnHeaders(X).Tag), "^select") > 0 Then
                    
                            sResult = Me.fldType(rsIn(iField).Type, bNumeric)
                            Select Case bNumeric
                            Case True
                                If Not IsNull(rsIn(iField)) Then
                                    SQL = iFormat & "'" & Val(rsIn(iField)) & "'"
                                Else
                                    SQL = iFormat & "'0'"
                                End If
                            Case False
                                If Not IsNull(rsIn(iField)) Then
                                    SQL = iFormat & "'" & Me.ESC(rsIn(iField)) & "'"
                                Else
                                    SQL = iFormat & "''"
                                End If
                            End Select
                                                   
                                                
                            Call Me.OpenTable(oConn, rsLoad, , SQL)
                            If rsLoad.State = adStateOpen Then
                                If rsLoad.RecordCount > 0 Then
                                   sResult = Me.fldType(rsLoad("nResult").Type, bNumeric)
                                   Select Case bNumeric
                                   Case True
                                        
                                        If X = 1 Then
                                            itmX.Text = "" & Val(IIf(IsNull(rsLoad("nResult")), 0, rsLoad("nResult")))
                                        Else
                                            itmX.SubItems(X - 1) = "" & Val(IIf(IsNull(rsLoad("nResult")), 0, rsLoad("nResult")))
                                        End If
                                        
                                   Case False
    
                                        If X = 1 Then
                                            itmX.Text = "" & IIf(IsNull(rsLoad("nResult")), 0, rsLoad("nResult"))
                                            If iField = "VendorID" Then itmX.Text = Me.CodeME(itmX.Text)
                                        Else
                                            itmX.SubItems(X - 1) = "" & IIf(IsNull(rsLoad("nResult")), "", rsLoad("nResult"))
                                            If iField = "VendorID" Then itmX.SubItems(X - 1) = Me.CodeME(itmX.SubItems(X - 1))
                                        End If
                                        
                                   End Select
                                End If
                            End If
                            
                        Else
                        
                            If X = 1 Then
                                itmX.Text = "" & Me.OP(rsIn, lv.ColumnHeaders(X).Tag)
                            Else
                                itmX.SubItems(X - 1) = "" & Me.OP(rsIn, lv.ColumnHeaders(X).Tag)
                            End If
                            
                        End If
                    End If
                    
                End If
                gSleep
            Next X
        rsIn.MoveNext
        If bSetAll = False Then Exit Do
    Loop Until rsIn.EOF
    
    Exit Function
        
ErrorOccur:
Select Case oErr.chkError(Err)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
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
Select Case oErr.chkError(Err)
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
Select Case oErr.chkError(Err)
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
Select Case oErr.chkError(Err)
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
    ESC = Me.ReplaceString(ESC, Chr$(26), "\z") ' ASCII(26) (Control-Z).  This character can be AES_ENCRYPTd to allow you to work around the problem that ASCII(26) stands for END-OF-FILE on Windows.  (ASCII(26) will cause problems if you try to use oMySQL database < filename.)
    'ESC = Me.ReplaceString(ESC, "%", "\%")
    'ESC = Me.ReplaceString(ESC, "_", "\_") ' A _ character. This is used to search for literal instances of _ in contexts where _ would otherwise be interpreted as a wildcard character. See String comparison functions.


Exit Function



ErrorOccur:
Select Case oErr.chkError(Err)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function
Function Connection(Optional db As String = "", Optional server As String = "", Optional UID As String = "", Optional PWD As String = "", Optional ByRef oConn As ADODB.Connection, Optional ByRef lblAction As Label = Nothing) As Boolean
    
    On Error Resume Next
           
    If bDebug = True Then On Error GoTo 0 Else On Error Resume Next
    
ReDo:
    
    Set oConn = New ADODB.Connection

    If Not lblAction Is Nothing Then
        lblAction.Caption = "Opening ADO 2.8 Connection"
        lblAction.Refresh
        gSleep
    End If
    
    If db <> "" Then
        oConn.Open "Driver={Mysql ODBC 3.51 Driver};" & _
                "Database=" + db + ";" & _
                "UID=" + UID + ";" & _
                "PWD=" + PWD + ";" & _
                "PORT=3306;" & _
                "OemToAnsi=No;" & _
                "SSL=Yes;" & _
                "Server=" + server + ";OPTION=2084;"
    
        cDebug "Driver={MySQL ODBC 3.51 Driver};" & _
                "Database=" + db + ";" & _
                "UID=" + UID + ";" & _
                "PORT=3306;" & _
                "OemToAnsi=No;" & _
                "Server=" + server + ";OPTION=2083;"
    Else
        
        Dim tmp As String
        
        tmp = GetSetting("projectalpha", "db", "ConnectionString", "")
        
        If tmp <> "" Then ConnStr = tmp
        
        If ConnStr = "" Then
            
            ShellLauncher
            
            End
            End
            End
            
        Else
            Dim commstr As String
            
            'MsgBox Crypt(ConnStr, False, "PublicKey", "projectalpha")
            'Stop
            oConn.Open Crypt(ConnStr, False, "PublicKey", "projectalpha")
            If oConn.State = adStateClosed Then
                MsgBox "Could not connect to server.", vbCritical, "ADO Error"
                End
            End If
            oConn.Execute "use projectalpha"
            SaveSetting "projectalpha", "db", "ConnectionString", ""
            
        End If
        
    End If
    
    Connection = True

    'oConn.Execute "alter table invoiceout add column (SubRecID Bigint default 0)"
If Err.Number = 0 Then
    oConn.Execute "use projectalpha"
    Exit Function
End If
    
    Static Count As Long
    Count = Count + 1
    If Count < 25 Then GoTo ReDo
     
    cDebug Err.Description
    
    'MsgBox Err.Description, vbCritical, "Could not connect to server " & server
    
    Connection = False
      
Exit Function




End Function

Public Function OpenTable(oConn As ADODB.Connection, rsADO As ADODB.Recordset, Optional Tablename As String, Optional SQLQuery As String) As Boolean


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
    
    If oConn.State = adStateClosed Then
        
        Select Case Login.bTestBench
        Case False
            If oMySQL.Connection(, sServer, , , oConn) = False Then
                MsgBox "Unable to Connect to oMySQL Server, Please check your internet connection and attempt to restart the program.", vbCritical, "MySQL Server Not Found"
                End
            End If
        Case True
            
            sServer = "169.254.60.62"
            sUID = "pa2004"
            sPWD = "projectselector"
            If oMySQL.Connection(, sServer, sUID, sPWD, oConn) = False Then
                MsgBox "Unable to Connect to oMySQL Test Bench Server, Please check your LAN connection and attempt to restart the program.", vbCritical, "MySQL Test Bench Not Found"
                End
            End If
        
        End Select
        
    End If

        
    Dim SQL As String
    
    If Tablename <> "" Then
        SQL = "Select * from " & Tablename
    ElseIf SQLQuery <> "" Then
        SQL = SQLQuery
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
    gSleep
    OpenTable = True
    gSleep
    OpenTable = True
    
    
    cDebug "ADO:// [" & Err.Number & "] - " & SQL
    
Exit Function

Exit Function



ErrorOccur:
OpenTable = False
Dim floop As Integer

floop = floop + 1
If floop = 5 Then Exit Function
cDebug "ERR:// [" & Err.Number & "] - " & SQL

Select Case oErr.chkError(Err)
Case vbResume
    Resume
Case vbExit

Case vbResumeNext
    Resume Next
End Select
    
    
End Function

Public Sub Execute(oConn As ADODB.Connection, SQL As String, Optional bHidden As Boolean)


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
    
    If oConn.State = adStateClosed Then
        
        Select Case Login.bTestBench
        Case False
            If oMySQL.Connection(, sServer, , , oConn) = False Then
                MsgBox "Unable to Connect to oMySQL Server, Please check your internet connection and attempt to restart the program.", vbCritical, "MySQL Server Not Found"
                End
            End If
        Case True
            
            sServer = "169.254.60.62"
            sUID = "pa2004"
            sPWD = "projectselector"
            If oMySQL.Connection(, sServer, sUID, sPWD, oConn) = False Then
                MsgBox "Unable to Connect to oMySQL Test Bench Server, Please check your LAN connection and attempt to restart the program.", vbCritical, "MySQL Test Bench Not Found"
                End
            End If
        
        End Select
        
    End If
           
    oConn.Execute SQL

    
    If bHidden = False Then cDebug "ADO:// [" & Err.Number & "] " & SQL
    
Exit Sub



ErrorOccur:
Select Case oErr.chkError(Err)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Sub

Public Function SetColumnHeaders(formcode As String, lv As ListView, ServiceKey As String, Optional oConn As ADODB.Connection) As Boolean


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "SetColumnHeaders"
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


    'On Error Resume Next
    
    lv.ColumnHeaders.Clear
    
    Dim rsLoad As ADODB.Recordset
    Dim bResult As Boolean
    Dim colHeader As ColumnHeader
    
    bResult = oMySQL.OpenTable(oConn, rsLoad, , "Select * from columnlayout where ServiceKey = '" & ServiceKey & "' and formcode = '" & formcode & "' order by cOrder")
    
    If rsLoad.RecordCount > 0 Then
        While Not rsLoad.EOF
            Set colHeader = lv.ColumnHeaders.Add(, , rsLoad!Description, rsLoad!Width)
            colHeader.Tag = "!" & rsLoad!FieldName & "!" & IIf(IsNull(rsLoad!sFormat), "", "^" & rsLoad!sFormat & "^")
            rsLoad.MoveNext
        Wend
    End If
    
    lv.Refresh
    gSleep
    SetColumnHeaders = True
    
Exit Function



ErrorOccur:
Select Case oErr.chkError(Err)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function

Public Function OP(ByRef oRS As ADODB.Recordset, FieldName As String) As String


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "OP"
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


    If oRS.State = adStateOpen Then
    
    FieldName = Me.ReplaceString(FieldName, "^^", "")
    
    If InStr(FieldName, "!") > 0 Then
        OP = FieldName
        Dim iPos As Integer
        Dim iPosb As Integer
        
        Dim iField As String
        Dim iFormat As String
        Do
            iPos = InStr(iPosb + 1, FieldName, "!")
            iPosb = InStr(iPos + 1, FieldName, "!")
            iField = Mid(FieldName, iPos + 1, iPosb - iPos - 1)
            If Mid(FieldName, iPosb + 1, 1) = "^" Then
                iPos = InStr(iPosb + 1, FieldName, "^")
                iPosb = InStr(iPos + 1, FieldName, "^")
                iFormat = Mid(FieldName, iPos + 1, iPosb - iPos - 1)
            Else
                iFormat = ""
            End If
            
            If LCase(iFormat) = "yesno" Then
                If Val(IIf(IsNull(oRS(iField)), 0, oRS(iField))) = 0 Then
                    OP = "No"
                Else
                    OP = "Yes"
                End If
            ElseIf LCase(iFormat) = "unlimitedmb" Then
                If Val(IIf(IsNull(oRS(iField)), 0, oRS(iField))) = -1 Then
                    OP = "Unlimited MB's"
                Else
                    OP = IIf(IsNull(oRS(iField)), "0 ", oRS(iField)) & " MB's"
                End If
            ElseIf LCase(iFormat) = "unlimitedhr" Then
                If Val(IIf(IsNull(oRS(iField)), 0, oRS(iField))) = -1 Then
                    OP = "Unlimited Hrs"
                Else
                    OP = IIf(IsNull(oRS(iField)), "0 ", oRS(iField)) & " Hrs"
                End If
            
            ElseIf LCase(iFieldName) = "vendorid" Then
                If iField <> "" Then OP = oMySQL.ReplaceString(OP, "!" & iField & "!" & IIf(iFormat <> "", "^" & iFormat & "^", ""), Me.CodeME(Format(IIf(IsNull(oRS(iField)), "", oRS(iField)), iFormat)))
            ElseIf InStr(FieldName, "iTax") > 0 Then
                Select Case Mid(iField, 5, 1)
                Case "/"
                    'If Not oRS.EOF And Not oRS.BOF Then OP = oMySQL.ReplaceString(OP, "!" & iField & "!" & IIf(iFormat <> "", "^" & iFormat & "^", ""), Format(oRS(Mid(iField, 6)) * oTax("G", "AUS001"), iFormat))
                Case "+"
                    'If Not oRS.EOF And Not oRS.BOF Then OP = oMySQL.ReplaceString(OP, "!" & iField & "!" & IIf(iFormat <> "", "^" & iFormat & "^", ""), Format(oRS(Mid(iField, 6)) * oTax("G", "AUS001") + oRS(Mid(iField, 6)), iFormat))
                Case Else
                    Stop
                End Select
            Else
                If iField <> "" Then OP = oMySQL.ReplaceString(OP, "!" & iField & "!" & IIf(iFormat <> "", "^" & iFormat & "^", ""), IIf(iFormat = "", IIf(IsNull(oRS(iField)), "", oRS(iField)), Format(IIf(IsNull(oRS(iField)), "", oRS(iField)), iFormat)))
            End If
            
        Loop Until InStr(iPos, "!", OP) <= 0
    
    ElseIf InStr(FieldName, "VendorID") > 0 Then
        If Not oRS.EOF And Not oRS.BOF Then OP = Me.CodeME(oRS(FieldName))
    Else
        If Not oRS.EOF And Not oRS.BOF Then OP = oRS(FieldName)
    End If
    End If
Exit Function



ErrorOccur:
Select Case oErr.chkError(Err)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function

Public Function virtualisp(ByVal sSQL As String, Optional TblName As String, Optional AllTiers As Boolean, Optional OverrideVID0 As Boolean) As String


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "virtualisp"
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

    
    If Login.lVirtualID = 0 Then
        If OverrideVID0 = False Then AllTiers = True
    End If
    

    sSQL = oMySQL.ReplaceString(LCase(sSQL), "from", "from virtualisp inner join " + TblName + " on virtualisp.RecID = " + TblName + ".VirtualID or virtualisp.VirtualID  = " + TblName + ".VirtualID,")
    sSQL = oMySQL.ReplaceString(sSQL, TblName + ", ", "")
    sSQL = oMySQL.ReplaceString(sSQL, ", " + TblName + "", "")
    
    
    'Clipboard.Clear
    'Clipboard.SetText sSQL
    

    Dim tmpSQL As String
    
    If ViSPMAP.Count > 0 And AllTiers = True Then
        Dim iCount As Long
        tmpSQL = "("
        For iCount = 1 To ViSPMAP.Count
            tmpSQL = tmpSQL + "'" & ViSPMAP(iCount).RecIDb & "',"
        Next
        tmpSQL = Left(tmpSQL, Len(tmpSQL) - 1) + ")"
    Else
        tmpSQL = "('" & Login.lVirtualID & "')"
    End If
    
    If InStr(LCase(sSQL), "select distinct") > 0 Then
        sSQL = oMySQL.ReplaceString(sSQL, "select distinct ", "select distinct " & TblName & "." & "VirtualID IN " + tmpSQL & ", ")
    Else
        sSQL = oMySQL.ReplaceString(sSQL, "select ", "select " & TblName & "." & "VirtualID IN " + tmpSQL & ", ")
    End If
    
    sSQL = oMySQL.ReplaceString(sSQL, ", where", " where")
    sSQL = oMySQL.ReplaceString(sSQL, ",  where", " where")
    sSQL = oMySQL.ReplaceString(sSQL, ",   where", " where")
    sSQL = oMySQL.ReplaceString(sSQL, ",    where", " where")
    
    'Debug.Print sSQL
    virtualisp = sSQL
    
    Exit Function
    '-------------------------------------------------------------------------------------------------------------------------------------------------
    '-- Old Routine
    '-------------------------------------------------------------------------------------------------------------------------------------------------
    Select Case Login.lVirtualID
    Case 0
        
    Case Else
        If InStr(LCase(sSQL), "order by") > 0 Then
            If InStr(LCase(sSQL), "where") > 0 Then
                'virtualisp = Trim(Left(sSQL, InStr(LCase(sSQL), "order by") - 1)) & " AND " & IIf(TblName <> "", TblName + ".", "") + "VirtualID = " & Login.lVirtualID & " and virtualisp.RecID = " & & Trim(Mid(sSQL, InStr(LCase(sSQL), "order by") - 1))
            Else
                virtualisp = Trim(Left(sSQL, InStr(LCase(sSQL), "order by") - 1)) & " WHERE " & IIf(TblName <> "", TblName + ".", "") + "VirtualID = " & Login.lVirtualID & " " & Trim(Mid(sSQL, InStr(LCase(sSQL), "order by") - 1))
            End If
        ElseIf InStr(LCase(sSQL), "limit") > 0 Then
            If InStr(LCase(sSQL), "where") > 0 Then
                virtualisp = Trim(Left(sSQL, InStr(LCase(sSQL), "limit") - 1)) & " AND " & IIf(TblName <> "", TblName + ".", "") + "VirtualID = " & Login.lVirtualID & " " & Trim(Mid(sSQL, InStr(LCase(sSQL), "limit") - 1))
            Else
                virtualisp = Trim(Left(sSQL, InStr(LCase(sSQL), "limit") - 1)) & " WHERE " & IIf(TblName <> "", TblName + ".", "") + "VirtualID = " & Login.lVirtualID & " " & Trim(Mid(sSQL, InStr(LCase(sSQL), "limit") - 1))
            End If
        Else
            If InStr(LCase(sSQL), "where") > 0 Then
                virtualisp = sSQL & " AND " & IIf(TblName <> "", TblName + ".", "") + "VirtualID = " & Login.lVirtualID
            Else
                virtualisp = sSQL & " WHERE " & IIf(TblName <> "", TblName + ".", "") + "VirtualID = " & Login.lVirtualID
            End If
        End If
    End Select

Exit Function



ErrorOccur:
Select Case oErr.chkError(Err)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function


Public Function SetNextCycle(rs As ADODB.Recordset, sType As String, Interval As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "SetNextCycle"
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


    rs!PreviousCycle = rs!NextCycle
    rs!NextCycle = DateAdd(sType, Interval, Now)
    
Exit Function



ErrorOccur:
Select Case oErr.chkError(Err)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function

Public Function SetRecID(rs As ADODB.Recordset, TblName As String, oConn As ADODB.Connection) As Variant


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "SetRecID"
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

    
    '9.22337203685478E+18
    Randomize Now
    
    On Error Resume Next
    
    Dim rsCount As ADODB.Recordset
    Dim bResult As Boolean
           
    If IsNull(rs!RecID) Or TblName = "plantypes" Then

        While Err.Number <> 0
            SetRecID = Me.GetTMPRecID(TblName, oConn)
            rs!RecID = SetRecID
            rs.Update
        Wend
    Else
        SetRecID = rs!RecID
    End If
    
Exit Function



ErrorOccur:
Select Case oErr.chkError(Err)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function

Public Function GetTMPRecID(Tablename As String, oConn As ADODB.Connection) As Long


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "GetTMPRecID"
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


    Static Direction As Boolean
    
    Dim rsRandom As ADODB.Recordset
    Dim rsMAX As ADODB.Recordset
    
    Dim bResult As Boolean
    Do
    
        Randomize Now
        
        Select Case Direction
        Case False
            
            bResult = Me.OpenTable(oConn, rsMAX, , "Select MAX(RecID) as maxRecID from recidplacement where TableName = '" + Tablename + "'")
            
            
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
                bResult = Me.OpenTable(oConn, rsMAX, , "Select MAX(RecID) as maxRecID from " + Tablename + "")
                On Error Resume Next
                If rsMAX.RecordCount > 0 Then
                    If GetTMPRecID <= Val(rsMAX!maxrecid) Then GetTMPRecID = IIf(IsNull(rsMAX!maxrecid), 0, rsMAX!maxrecid)
                End If
            End If
            
            Do
                GetTMPRecID = GetTMPRecID + 1
                bResult = Me.OpenTable(oConn, rsRandom, , "Select * from recidplacement where RecID = " & GetTMPRecID & " AND TableName = '" & Tablename & "' Limit 1")
                gSleep
            Loop Until rsRandom.RecordCount = 0


            rsRandom.AddNew
            rsRandom!RecID = GetTMPRecID
            rsRandom!Tablename = Tablename
            rsRandom.Update
            rsRandom.Close
                       
             Direction = True
             
            'mysql.execute oConn,  "Delete from recidplacement Where RecID < " & GetTMPRecID
            
            Me.Execute oConn, "DELETE FROM recidplacement where TableName = '" & Tablename & "' and RecID " & IIf(Direction = False, "^", ">") & " " & GetTMPRecID
            If Tablename = "cctransactions" Then
                On Error Resume Next
                Err.Clear
                oMySQL.Execute oConn, "Insert into " & Tablename & " (RecID) Values (" & GetTMPRecID & ")"
                
                If Err.Number = 0 Then Exit Do
            ElseIf Tablename = "accountinfo" Then
                On Error Resume Next
                Err.Clear
                oMySQL.Execute oConn, "Insert into " & Tablename & " (RecID) Values (" & GetTMPRecID & ")"
                
                If Err.Number = 0 Then Exit Do
            Else
                Exit Do
            End If
                        
                  
                  
        Case True
        
        
            bResult = Me.OpenTable(oConn, rsMAX, , "Select Min(RecID) as minRecID from recidplacement where TableName = '" + Tablename + "'")
            
            
            If bResult = True Then
                If rsMAX.RecordCount > 0 Then
                    GetTMPRecID = IIf(IsNull(rsMAX!minRecID), 0, rsMAX!minRecID)
                Else
                    GetTMPRecID = 0
                End If
            Else
                GetTMPRecID = 0
            End If
            
            
            If GetTMPRecID = 0 Then
                bResult = Me.OpenTable(oConn, rsMAX, , "Select min(RecID) as minRecID from " + Tablename + "")
                On Error Resume Next
                If rsMAX.RecordCount > 0 Then
                    If IsNull(rsMAX!minRecID) Then
                        GetTMPRecID = -1
                    Else
                        If GetTMPRecID <= Val(rsMAX!minRecID) Then GetTMPRecID = IIf(IsNull(rsMAX!minRecID), -1, rsMAX!minRecID)
                        If GetTMPRecID = 0 Or GetTMPRecID > 0 Then GetTMPRecID = -1
                    End If
                End If
            End If
            
            Do
                GetTMPRecID = GetTMPRecID - 1
                bResult = Me.OpenTable(oConn, rsRandom, , "Select * from recidplacement where RecID = " & GetTMPRecID & " AND TableName = '" & Tablename & "' Limit 1")
                gSleep
            Loop Until rsRandom.RecordCount = 0
        
            
            rsRandom.AddNew
            rsRandom!RecID = GetTMPRecID
            rsRandom!Tablename = Tablename
            rsRandom.Update
            rsRandom.Close
                       
            'mysql.execute oConn,  "Delete from recidplacement Where RecID < " & GetTMPRecID
            Direction = False
            
            Me.Execute oConn, "DELETE FROM recidplacement where TableName = '" & Tablename & "' and RecID " & IIf(Direction = False, "^", ">") & " " & GetTMPRecID
            If Tablename = "cctransactions" Then
                On Error Resume Next
                Err.Clear
                oMySQL.Execute oConn, "Insert into " & Tablename & " (RecID) Values (" & GetTMPRecID & ")"
                
                If Err.Number = 0 Then Exit Do
            ElseIf Tablename = "accountinfo" Then
                On Error Resume Next
                Err.Clear
                oMySQL.Execute oConn, "Insert into " & Tablename & " (RecID) Values (" & GetTMPRecID & ")"
                
                If Err.Number = 0 Then Exit Do
            Else
                Exit Do
            End If
            
            
        End Select
                
    Loop
    
Exit Function



ErrorOccur:
Select Case oErr.chkError(Err)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function

Public Function SetInvoiceSerial(Optional oConn As ADODB.Connection) As Variant


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "SetInvoiceSerial"
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

    
    Dim bResult As Boolean
    Dim rsLoad As ADODB.Recordset
    Dim InvSerialNumber As Variant
    
    bResult = oMySQL.OpenTable(oConn, rsLoad, , "Select RecID from invoicetraxr order by RecID DESC Limit 1")
    If rsLoad.RecordCount > 0 Then
        InvSerialNumber = rsLoad!RecID
    Else
        InvSerialNumber = 10000
    End If
    
    Do
        InvSerialNumber = InvSerialNumber + 1
        bResult = oMySQL.OpenTable(oConn, rsLoad, , "Select * from invoicetraxr Where InvoiceSerial = '" & Hex(InvSerialNumber) & "'")
    Loop Until rsLoad.RecordCount = 0
    
    SetInvoiceSerial = InvSerialNumber
    
Exit Function



ErrorOccur:
Select Case oErr.chkError(Err)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function


Public Function Update(rs As ADODB.Recordset, TblName As String, PointerField As String) As Boolean


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Update"
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


Exit Function



ErrorOccur:
Select Case oErr.chkError(Err)
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
Select Case oErr.chkError(Err)
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
Select Case oErr.chkError(Err)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function

Public Function AddReceiptItem(oConn As ADODB.Connection, Optional acciRecID As Variant = 0, Optional invoiceoutID As Variant = 0, Optional invoiceinID As Variant = 0, Optional acciServicesID As Variant = 0, Optional RefundID As Variant = 0, Optional Paid As Currency = 0, Optional Refunded As Currency = 0, Optional paymenttype As String, Optional Num As String, Optional TraxrID As Long) As Variant


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "AddReceiptItem"
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


    Dim RecID As Variant
    
    If (InStr(paymenttype, "Cheque") > 0 Or InStr(paymenttype, "Order") > 0) And (Num = "") Then
        frmSN.Show 1
        Num = frmSN.Num
    End If
    
    Do
        Err.Clear
        RecID = Me.GetTMPRecID("receipts", oConn)
        Me.Execute oConn, "insert into receipts (RecID, acci_RecID, invoiceoutID, invoiceinID, RefundID, acciServicesID, Paid, Refunded, paymenttype, SerialNumber, TraxrID) VALUES (" & RecID & "," & acciRecID & "," & invoiceoutID & "," & invoiceinID & "," & RefundID & "," & acciServicesID & "," & Paid & "," & Refunded & ",'" & paymenttype & "','" & Num & "','" & TraxrID & "')"
        If Err.Number <> 0 Then cDebug Err.Description
    Loop Until Err.Number = 0
    
    AddReceiptItem = RecID
    
Exit Function



ErrorOccur:
Select Case oErr.chkError(Err)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function


Public Function AddBonus(oConn As ADODB.Connection, sField As String, Text As String, Credit As Long, AwardTo As Integer, Optional SysopID As Long, Optional AgencyID As Long, Optional VirtualID As Long, Optional DevID As Long, Optional ptRecID As Long, Optional acci_RecID As Long, Optional PlnSrvID As Long)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "AddBonus"
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


    Me.Execute oConn, "INSERT INTO bonus_matrix (" & sField & ", TransactionText, SysopID, AgencyID, VirtualID, DeveloperID, AwardTo, ptRecID, acci_RecID, PlanServiceID) VALUES (" & Credit & ",'" & Me.ESC(Text) & "'," & SysopID & "," & AgencyID & "," & VirtualID & "," & DevID & "," & AwardTo & "," & ptRecID & "," & acci_RecID & "," & PlnSrvID & ")"
    

Exit Function



ErrorOccur:
Select Case oErr.chkError(Err)
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
