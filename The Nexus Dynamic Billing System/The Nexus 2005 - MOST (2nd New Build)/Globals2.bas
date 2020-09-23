Attribute VB_Name = "Globals"
Type type_pbBar
    Value As Single
    Max As Single
    Min As Single
    pb As ProgressBar
End Type


Global INClause As String

Global iStop As Long
Global iCont As Long

Global oFlags As New clsFlags

' ---------------------------------------------------------------------------------------------------
' | Class Fetch Status Codes
' ---------------------------------------------------------------------------------------------------

Global Const fs_NoChanges = 1
Global Const fs_Edited = 2
Global Const fs_NewLine_Insert = 3
Global Const fs_DeleteRecord = 4
Global Const fs_LoadingData = 5
Global Const fs_Saving = 6
Global Const fs_Deleting = 7
Global Const fs_Idle = 8
Global Const fs_CreateNewViSP = 9
Global Const fs_AccountHeader = 10
Global Const fs_Addresses = 11
Global Const fs_EmailAddresses = 12
Global Const fs_PhoneNumbers = 13
Global Const fs_InvoiceItems = 14
Global Const fs_PaymentHistory = 15
Global Const fs_LoadHeader = 16
Global Const fs_LoadInvoice = 17
Global Const fs_LoadAllContactDetails = 18
Global Const fs_LoadPaymentHistory = 19
Global Const fs_LoadAll = 20
Global Const fs_LoadMinimum = 21
Global Const fs_LoadEmail = 22
Global Const fs_LoadPhone = 23
Global Const fs_LoadAddress = 24

Public Enum enumFetchStatus
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

Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long

Global oAudit As New clsAudit
Global oReseller As New colReseller
Global oResell As New colReseller

Enum enumAuditType
    Client
    Bot
End Enum

Enum enumVarType
    vURL
    vUNC
    vHTML
    vXML
    vString
    vDouble
    vLong
    vCurrency
    vInteger
    vByte
    vBoolean
    vText
    vVarChar
    vVariant
End Enum

 Enum enumFormState
    Saved = 1
    Cancelled = 2
    Deleted = 3
    Loading = 4
    Waiting = 5
    Finished = 6
End Enum

Global Const BIF_RETURNONLYFSDIRS = 1
Global Const MAX_PATH = 260
Global Const MAXDWORD = &HFFFF
Global Const INVALID_HANDLE_VALUE = -1
Global Const FILE_ATTRIBUTE_ARCHIVE = &H20
Global Const FILE_ATTRIBUTE_DIRECTORY = &H10
Global Const FILE_ATTRIBUTE_HIDDEN = &H2
Global Const FILE_ATTRIBUTE_NORMAL = &H80
Global Const FILE_ATTRIBUTE_READONLY = &H1
Global Const FILE_ATTRIBUTE_SYSTEM = &H4
Global Const FILE_ATTRIBUTE_TEMPORARY = &H100

Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

'APIs for the folder selection
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lprogressbar1i As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lprogressbar1uffer As String) As Long

'APIs used to find files.
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long


Global bDebug As Integer

Global fIcon As New frmCompiledIcons

'Global Const EMAILSalt = "emailsalt"

Global Const clrPaid = 2015232
Global Const clrPartial = 3297408
Global Const clrNoPay = 983295

Global oErr As New clsErrors
Global oService As colPackages
Global oAccInf As New clsAccInf
Global oIP As New clsIP
Global oViSP As New clsVirtualISP

Global ViSPMAP As New mapVirtualISP
Global NDECHRT As New mapAccHoldings
Global odb As New clsDB

Global ConnStr As String

Global DirName As String
Global PrimaryIDName As String

Global sysnow As Date

Enum NumberorString
    nString = 1
    nNumber = 2
End Enum
    


Type odbFields
    FieldName As String
    CipherCode As String
    Encrypted As Boolean
    Salt As String
    ControlName As String
    ControlIndx As Integer
    TypeID As Long
End Type


Type invoiceout
    RecID As Long
    SubRecID As Long
    Description As String
    PaidWhen As Date
    StartCycle As Date
    EndCycle As Date
    AmountDue As Currency
    GSTCharged As Currency
    TotalDue As Currency
    AmountRefunded As Currency
    GSTRefunded As Currency
    AmountPaid As Currency
    acci_RecID As Long
    SysopID As Long
    PlanServiceID As Long
    PaymentDue As Date
    VirtualID  As Long
    StatementID As Long
    sfCycle_Download As Variant
    sfCycle_Upload As Variant
    sfCycle_Mins As Variant
    FlagID As Byte
End Type

Public Enum spiffSource
    tosysop
    toSite
    toAgency
    toDiv
End Enum

Global bBigFont As Boolean


Type Login_Details
    
    IconsSet As Boolean
    ViSPDesc As String
    
    TaxMode As Byte
    TaxCode As String
    TaxCountry As String
    
    lSysopID As Long
    sUsername As String
    lVirtualID As Long
    lLevel As Byte
    bTestBench As Boolean
    bTestBench2 As Boolean
    bSkipupgrade As Boolean
    sVISPDomain As String
    bVISPPrimary As Boolean
    sVISPLogo As String
    sVISPName As String
    
    lAgencyID As Long
    MISCFee As Variant
    
    bAgency As Boolean
    bCreateSysop As Boolean
    bPrimary As Boolean
    bTemplates As Boolean
    bRunMaintenance As Boolean

    bRecievables        As Boolean
    bInvoice            As Boolean
    bExpenditure        As Boolean
    bHoldings           As Boolean
    bComm               As Boolean
    bRefund             As Boolean
    bAddCust            As Boolean
    bOwnership          As Boolean
    bAccSettings        As Boolean
    bVendors            As Boolean
    bMaster As Boolean
    bVISP As Boolean
    bVISPFiscal As Boolean

    PublicKey As String
    
End Type

Global Login As Login_Details

'Global oDAOConn As DAO.Connection


Type tyAddress
    UnitNumber As String
    StreetNo As String
    StreetName As String
    StreetType As String
    Suburb As String
    State As String
    Country As String
    PostCode As String
End Type


Type IPs_Type
    First As Integer
    Second As Integer
    Third As Integer
    Fourth As Integer
End Type

Type IPINFO
     dwAddr As Long   ' IP address
    dwIndex As Long '  interface index
    dwMask As Long ' subnet mask
    dwBCastAddr As Long ' broadcast address
    dwReasmSize  As Long ' assembly size
    unused1 As Integer ' not currently used
    unused2 As Integer '; not currently used
End Type

Const MAX_IP = 10  'To make a buffer... i dont think you have more than 5 ip on your pc..

Type MIB_IPADDRTABLE
    dEntrys As Long   'number of entrues in the table
    mIPInfo(MAX_IP) As IPINFO  'array of IP address entrues
End Type

Type IP_Array
    mBuffer As MIB_IPADDRTABLE
    BufferLen As Long
End Type

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long

Global frmMDIMain_Loaded As Boolean

Global Const vbResumeNext = 1
Global Const vbResume = 2
Global Const vbExit = 3

Global Const EFTPOS = 0
Global Const Visa = 1
Global Const Mastercard = 2
Global Const AmericanExpress = 3
Global Const DinnersClub = 4
Global Const Discover = 5
Global Const JCB = 6


Global fLine As frmLines
Global Radius As New colRadius

Enum enum_FlagID
    bNotSet
    bProcessed1
    bProcessed2
    bProcessed3
    bProcessed4
    bProcessed5
    bProcessed6
    bProcessed7
    bProcessed8
    bProcessed9
    bProcessed0
End Enum


Public Const MAX_HOSTNAME_LEN = 132
Public Const MAX_DOMAIN_NAME_LEN = 132
Public Const MAX_SCOPE_ID_LEN = 260
Public Const MAX_ADAPTER_NAME_LENGTH = 260
Public Const MAX_ADAPTER_ADDRESS_LENGTH = 8
Public Const MAX_ADAPTER_DESCRIPTION_LENGTH = 132
Public Const ERROR_BUFFER_OVERFLOW = 111
Public Const MIB_IF_TYPE_ETHERNET = 1
Public Const MIB_IF_TYPE_TOKENRING = 2
Public Const MIB_IF_TYPE_FDDI = 3
Public Const MIB_IF_TYPE_PPP = 4
Public Const MIB_IF_TYPE_LOOPBACK = 5
Public Const MIB_IF_TYPE_SLIP = 6

Type IP_ADDR_STRING
    Next As Long
    IPAddress As String * 16
    IpMask As String * 16
    Context As Long
End Type

Type IP_ADAPTER_INFO
    Next As Long
    ComboIndex As Long
    AdapterName As String * MAX_ADAPTER_NAME_LENGTH
    Description As String * MAX_ADAPTER_DESCRIPTION_LENGTH
    AddressLength As Long
    Address(MAX_ADAPTER_ADDRESS_LENGTH - 1) As Byte
    Index As Long
    Type As Long
    DhcpEnabled As Long
    CurrentIpAddress As Long
    IpAddressList As IP_ADDR_STRING
    GatewayList As IP_ADDR_STRING
    DhcpServer As IP_ADDR_STRING
    HaveWins As Boolean
    PrimaryWinsServer As IP_ADDR_STRING
    SecondaryWinsServer As IP_ADDR_STRING
    LeaseObtained As Long
    LeaseExpires As Long
End Type

Type FIXED_INFO
    HostName As String * MAX_HOSTNAME_LEN
    DomainName As String * MAX_DOMAIN_NAME_LEN
    CurrentDnsServer As Long
    DnsServerList As IP_ADDR_STRING
    NodeType As Long
    ScopeId  As String * MAX_SCOPE_ID_LEN
    EnableRouting As Long
    EnableProxy As Long
    EnableDns As Long
End Type

Public Declare Function GetNetworkParams Lib "IPHlpApi" (FixedInfo As Any, pOutBufLen As Long) As Long
Public Declare Function GetAdaptersInfo Lib "IPHlpApi" (IpAdapterInfo As Any, pOutBufLen As Long) As Long
'Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const Flags = SWP_NOMOVE Or SWP_NOSIZE

Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2

Global MySQL As New clsMySQL
Global GUI As New clsInterface
Global sSTR As New clsStrings
Global cRadius As New colRadius
Global cErr As New clsErrors

Global sServer As String
Global sUID As String
Global sPWD As String


Type tySchedule
    iNumUserFiles As Integer
    iNumRadius As Integer
End Type

Global Schedule As tySchedule

Type Settings

    SleepForMS As Long
    Sleep2Doevents As Long
    
    sRadiusFTPPort As String
    sRadiusFTPServer As String
    sRadiusFTPUsername As String
    sRadiusFTPPassword As String
    sTargetDir As String
    sTargetFilename As String
    
    iRadiusHistory As Long
    iUpkeep As Long
    iUnpaid As Long
    iSendPO As Long
    iUpdate As Long
    
    s2RadiusFTPPort As String
    s2RadiusFTPServer As String
    s2RadiusFTPUsername As String
    s2RadiusFTPPassword As String
    s2TargetDir As String
    s2TargetFilename As String
    
    smtpSetYet As Boolean
    smtpServer As String
    smtpPort As String
    smtpDomain As String
    smtpUsername As String
    smtpPassword As String
    
    ReplyAddress As String
    bManual(2) As Boolean
    
End Type

Type Column_Type
    ColumnWidth As Long
    ColumnTitle As String
End Type

Public Enum enumSchedula
    UpdateUserFile
    DownloadRadius
End Enum

Public Enum frm_CloseStates
    frmCloseCancel
    frmCloseSave
End Enum

Global eScedula As enumSchedula

'Global Const frmCloseCancel = 0
'Global Const frmCloseSave = 1

Global reg As Settings

Type POINTAPI
    X As Long
    Y As Long
End Type

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Global Point As POINTAPI

Public Declare Function GetSystemMenu Lib "user32.dll" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32.dll" (ByVal hMenu As Long) As Long
Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    ftype As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type
Public Const MIIM_STATE = &H1
Public Const MIIM_ID = &H2
Public Const MIIM_TYPE = &H10
Public Const MFT_SEPARATOR = &H800
Public Const MFT_STRING = &H0
Public Const MFS_ENABLED = &H0
Public Const MFS_CHECKED = &H8
Public Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Public Declare Function SetMenuItemInfo Lib "user32.dll" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_WNDPROC = -4
Public Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_SYSCOMMAND = &H112
Public Const WM_INITMENU = &H116
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Public Const MF_BITMAP = 4
Public Const MF_CHECKED = 8
' Add an option to make window Form1 "Always On Top" to the bottom of its system
' menu.  A check mark appears next to this option when active.  The menu item acts as a toggle.
' Note how subclassing the window is necessary to process the two messages needed
' to give the added system menu item its full functionality.

' *** Place the following code in a module. ***

Public pOldProc As Long  ' pointer to Form1's previous window procedure
Public ontop As Boolean  ' identifies if Form1 is always on top or not

' The following function acts as Form1's window procedure to process messages.
Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim hSysMenu As Long     ' handle to Form1's system menu
    Dim mii As MENUITEMINFO  ' menu item information for Always On Top
    Dim retval As Long       ' return value
    
    Select Case uMsg
    Case WM_INITMENU
        ' Before displaying the system menu, make sure that the Always On Top
        ' option is properly checked.
        hSysMenu = GetSystemMenu(hwnd, 0)
        With mii
            ' Size of the structure.
            .cbSize = Len(mii)
            ' Only use what needs to be changed.
            .fMask = MIIM_STATE
            ' If Form1 is now always on top, check the item.
            .fState = MFS_ENABLED Or IIf(ontop, MFS_CHECKED, 0)
        End With
        retval = SetMenuItemInfo(hSysMenu, 1, 0, mii)
        WindowProc = 0
    Case WM_SYSCOMMAND
        ' If Always On Top (ID = 1) was selected, change the on top/not on top
        ' setting of Form1 to match.
        If wParam = 1 Then
            ' Reverse the setting and make it the current one.
            ontop = Not ontop
            retval = SetWindowPos(hwnd, IIf(ontop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
            WindowProc = 0
        Else
            ' Some other item was selected.  Let the previous window procedure
            ' process it.
            WindowProc = CallWindowProc(pOldProc, hwnd, uMsg, wParam, lParam)
        End If
    Case Else
        ' If this is some other message, let the previous procedure handle it.
        WindowProc = CallWindowProc(pOldProc, hwnd, uMsg, wParam, lParam)
    End Select
End Function

Public Sub AddAudit(aType As enumAuditType, vVarType As enumVarType, vVarName As String, frm As Form, Optional newcur As Single, Optional oldcur As Single, Optional newvalue As String, Optional oldvalue As String, Optional newpointer As Long, Optional oldpointer As Long, Optional Description As String, Optional acci_RecID As Long, Optional RefundID As Long, Optional invTraxrID As Long, Optional FlagID As Long)

    Dim vartype As String
    
    Select Case vVarType
    Case vURL
        vartype = "URL"
    Case vUNC
        vartype = "UNC"
    Case vHTML
        vartype = "HTML"
    Case vXML
        vartype = "XML"
    Case vString
        vartype = "String"
    Case vDouble
        vartype = "Double"
    Case vLong
        vartype = "Long"
    Case vCurrency
        vartype = "Currency"
    Case vInteger
        vartype = "Integer"
    Case vByte
        vartype = "Byte"
    Case vBoolean
        vartype = "Boolean"
    Case vText
        vartype = "Text"
    Case vVarChar
        vartype = "VarChar"
    Case vVariant
        vartype = "Variant"
    End Select

    oAudit.colAuditTrail.Add "", Now, sysnow, Now, App.EXEName, App.Major & "." & App.Minor & ".0." & App.Revision, App.hInstance, frm.Name, frm.hwnd, vartype, vVarName, oldcur, newcur, oldvalue, newvalue, CDbl(oldpointer), CDbl(newpointer), 0, Login.lSysopID, Login.lVirtualID, Login.lAgencyID, acci_RecID, RefundID, invTraxrID, FlagID, Description, False, ""
    
End Sub

Public Sub ShellLauncher()

    Dim Path As String, Extension As String
    
    Path = IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\")
    
            SaveSetting "projectalpha", "db", "ConnectionString", ""
    
    If Dir(Path + "dblauncher.exe", vbNormal) <> "" Then
            
            
            ShellExecute frmAgent.hwnd, vbNullString, Path + "dblauncher.exe " + Extension, vbNullString, "C:\", SW_SHOWNORMAL
            
    End If
    
        
End Sub

Public Sub ShellUgrade()

    Dim Path As String, Extension As String
    
    Path = IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\")
    
            SaveSetting "projectalpha", "db", "ConnectionString", Crypt(ConnStr, True, "PublicKey")
    
    If Dir(Path + "pa_deployment.exe", vbNormal) <> "" Then
            
            
            ShellExecute frmAgent.hwnd, vbNullString, Path + "pa_deployment.exe", vbNullString, App.Path, SW_SHOWNORMAL
            
    End If
    
        
End Sub
Sub Main()


    '*[ Error Checking Variables ]**********************************************************************************
    Const RoutineName = "Main"
    Const ContainerName = "Globals"
    '***************************************************************************************************************


    '
    '***********************************************************************************************
    '**  Project Alpha Æ 2003, 2004 +                                                             **
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
    '**  Description:    Subroutine, Function or Property of The Nexus                        **
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

    PrimaryIDName = GetSetting(App.ProductName, "db", "PrimaryID", "RecID")
    
        
    Dim msgType As enumMessages
    
    If InStr(1, LCase(Command), "/debug") > 0 Then bDebug = True
    If InStr(1, LCase(Command), "/upgrade") = 0 Then Login.bSkipupgrade = True
    If InStr(1, LCase(Command), "/bigfont") > 0 Then bBigFont = True
    
    If InStr(1, LCase(Command), "/testbench") > 0 Then
    
    Else
    
        tmp = GetSetting(App.ProductName, "db", "ConnectionString", "")
            
        If tmp <> "" Then ConnStr = tmp
        
        If ConnStr = "" Then
            
            ShellLauncher
            
            End
            End
            End
            
        End If
    End If
    
    SaveSetting App.ProductName, "Main", "Major", App.Major
    SaveSetting App.ProductName, "Main", "Minor", App.Minor
    SaveSetting App.ProductName, "Main", "Revision", App.Revision
    
    
    If GetSetting(App.ProductName, "PublicKey", "B-0", -1) = -1 Then
        
        Dim fKey As New frmPublicKey
        
        fKey.KeyName = "PublicKey"
        
        fKey.Show 1
    
    End If
    
    Load frmAgent
    
    frmSplash.Show
    
Exit Sub



ErrorOccur:
Select Case oErr.chkError(directConn, Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Sub

Public Function RunOpenSequence(pb As ProgressBar, lblAction As Label)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "RunOpenSequence"
    Const ContainerName = "PublicSubroutines"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha Æ 2003, 2004 +                                                             **
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
'**  Description:    Subroutine, Function or Property of The Nexus                        **
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
    
    Const TotalAction = 5
    
    Dim bCloseDown As Boolean
        
    pb.Max = TotalAction
    
    lblAction.Caption = "Opening ADO Connection"
    gSleep
    'Opens ADO Connection To Server
    
    If InStr(LCase(Command), "/testbench") > 0 Or InStr(LCase(Command), "/testbench2") > 0 Then Login.bTestBench = True
    
    sServer = gte
    
    Select Case Login.bTestBench
    Case False
            If MySQL.Connection(, , , , directConn, lblAction) = False Then
                MsgBox "Unable to Connect to MySQL Server, Please check your internet connection and attempt to restart the program. It could be that you have not installed the MyODBC Drivers, these are located in the start bar under 'The Nexus'. There also may be limited bandwidth not allowing for a secure connection to be established with the server.", vbCritical, "MySQL Server Not Found"
                End
            End If
    Case True
        
        If InStr(LCase(Command), "/testbench2") > 0 Then
            sServer = "localhost"
            sUID = Crypt("øªë00ä", False, "None")
            sPWD = Crypt("¡≈¥Ÿ°‘œ∂±‘∆Ÿ‘ó¥Ÿ±", False, "None")
        Else
            sServer = Crypt("ë0ë.¢Õë.¢ëè.ë≤", False, "None")
            sUID = Crypt("øªë00ä", False, "None")
            sPWD = Crypt("¡≈¥Ÿ°∂±œ∆¶", False, "None")
        End If
        
        If MySQL.Connection("projectalpha", sServer, sUID, sPWD, directConn, lblAction) = False Then
            MsgBox "Unable to Connect to MySQL Test Bench Server [ " & sServer & " ], Please check your LAN connection and attempt to restart the program.", vbCritical, "MySQL Test Bench Not Found"
            End
        End If
    

    End Select
    
            
        Dim rsKeys As adodb.Recordset
        Dim iCnt2 As Long
        
        lblAction.Caption = "Loading Encryption Systems..... 0%"
        lblAction.Refresh
        
        Call MySQL.OpenTable(directConn, rsKeys, , "select RecID, sName, AES_DECRYPT(SuperKey, RecUnlock) as sKey from superkeysindex")
                   
        
        If rsKeys.State = adStateOpen Then
            pb.Max = pb.Max + rsKeys.RecordCount

            iCnt2 = 0
            odb.colSalts.Clear
            
            While Not rsKeys.EOF And Err.Number = 0
            
                iCnt2 = iCnt2 + 1
                odb.colSalts.Add "s" & rsKeys!RecID, IIf(IsNull(rsKeys!sName), "", rsKeys!sName), IIf(IsNull(rsKeys!sKey), "", rsKeys!sKey), "s" & rsKeys!RecID
                lblAction.Caption = "Loading Encryption Systems..... " & Round((iCnt2 / rsKeys.RecordCount) * 100) & "%"
                rsKeys.MoveNext
                
            Wend
        End If
        
        rsKeys.Close
            
        If bDebug = True Then frmDebug.Show
        
        
        Dim ffrmLogin As New frmLogin
        If bCloseDown = False Then ffrmLogin.Show 1
        
        Unload ffrmLogin
        
        Set fLogin = ffrmLogin
        
        lblAction.Caption = "Checking for Version Update"
        lblAction.Refresh
        
        Dim rsload As adodb.Recordset
        bResult = MySQL.OpenTable(directConn, rsload, , "select RecID, AES_DECRYPT(Password,'" & odb.colSalts.ReturnSalt("md5Password") & "') as Password, Port, Username, Server, Filename, Version, revision, MSI from upgrade Where revision > " & App.Revision & " Order By revision DESC Limit 1")
        If rsload.RecordCount >= 1 Then
            Select Case MsgBox("There is a newer version of The Nexus available, this is version " & IIf(IsNull(rsload!Version), "0.0.0.0", rsload!Version) & ". Would you like to upgrade now?", vbCritical + vbYesNo, "Newer Version Available")
            Case vbYes
                ShellUgrade
                End
            End Select
        End If
        PopulateResellers = PopulateResellers + 1: pb.Value = pb.Value + 1
        
        lblAction.Caption = "Getting Ready to Populating Icon Set"
        lblAction.Refresh
        Dim fIcon As New frmIcons
        gSleep
        LoadRegistry
        PopulateResellers = PopulateResellers + 1: pb.Value = pb.Value + 1
        fIcon.Show
        
        
        'lblAction.Caption = "Opening Jet Database"
        'gSleep
        'Call OpenJetConn
        lblAction.Caption = "Building Tax Profile"
        lblAction.Refresh
        Dim tCont As Long
        
        If MySQL.OpenTable(directConn, rsload, , "select * from tax") = True Then
            If rsload.State = adStateOpen Then
                
                If rsload.RecordCount > 0 Then
                    GUI.mapTax.Clear
                    pb.Max = pb.Max + rsload.RecordCount + 1
                    Do
                        GUI.mapTax.Add "r" & rsload!RecID, rsload!Code, rsload!Country, rsload!Percentage, rsload!iFlag, rsload!Description, rsload!lGroup, Val(rsload!RangeMin), Val(rsload!RangeMax), Val(rsload!FlatRate), "r" & rsload!RecID
                        
                        PopulateResellers = PopulateResellers + 1: pb.Value = pb.Value + 1
                        rsload.MoveNext
                        gSleep
                        tCont = tCont + 1
                        lblAction.Caption = "Building Tax Profile - " & (tCont / rsload.RecordCount) * 100 & "%"
                    Loop Until rsload.EOF Or Err.Number <> 0
                    rsload.Close
                End If
                
            End If
        End If
        Pause
        PopulateResellers = PopulateResellers + 1: pb.Value = pb.Value + 1
        
        
    
        PopulateResellers = PopulateResellers + 1: pb.Value = pb.Value + 1
        
        'frmSplash.picBg.Picture = frmSplash.picBg.Image
        
        If reg.smtpSetYet = False Then
            frmSMTP.Show 1
        End If
        
        pb.Value = pb.Max
                       

        If bCloseDown = False Then
            'Load frmMDIMain
            frmMDIMain.Show
            Unload frmSplash
        End If
        
        pb.Value = pb.Max
        
    
Exit Function



ErrorOccur:

If Err.Number = 3709 Then
    Msg = "There appears to be no connection with the WAN source server. This can occur from TCP/IP Faults and Firewalls." & vbCrLf
    Msg = Msg + vbCrLf + "Also it means that you cannot access project alphas main system or log onto this server at this time."
    Msg = Msg + "Please be advised to contact your system administrator reguarding this and allow for time to check permissions on firewalls."
    Msg = Msg + vbCrLf & vbclr & "You will have to allow port 3306 on your hardware firewall for MySQL and install the latest MyODBC Drivers."
    Msg = Msg + "These are available from http://www.mysql.com/, under 'binary downloads'."
    
    MsgBox Msg, vbCritical, "No WAN/LAN/INTERNET Access At this Time"
    
    End

End If

Select Case oErr.chkError(directConn, Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Function

Public Function Crypt(txt As String, bEncrypt As Boolean, RegKeyType As String) As String

    Select Case bEncrypt
    Case False
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-86", 187)), Chr(GetSetting(App.ProductName, RegKeyType, "A-86", 97))) ' 187[ª] = 97[a]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-85", 208)), Chr(GetSetting(App.ProductName, RegKeyType, "A-85", 70))) ' 208[–] = 70[F]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-84", 176)), Chr(GetSetting(App.ProductName, RegKeyType, "A-84", 86))) ' 176[∞] = 86[V]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-83", 179)), Chr(GetSetting(App.ProductName, RegKeyType, "A-83", 60))) ' 179[≥] = 60[<]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-82", 190)), Chr(GetSetting(App.ProductName, RegKeyType, "A-82", 106))) ' 190[æ] = 106[j]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-81", 161)), Chr(GetSetting(App.ProductName, RegKeyType, "A-81", 103))) ' 161[°] = 103[g]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-80", 145)), Chr(GetSetting(App.ProductName, RegKeyType, "A-80", 50))) ' 145[ë] = 50[2]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-79", 207)), Chr(GetSetting(App.ProductName, RegKeyType, "A-79", 111))) ' 207[œ] = 111[o]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-78", 130)), Chr(GetSetting(App.ProductName, RegKeyType, "A-78", 108))) ' 130[Ç] = 108[l]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-77", 184)), Chr(GetSetting(App.ProductName, RegKeyType, "A-77", 129))) ' 184[∏] = 129[Å]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-76", 209)), Chr(GetSetting(App.ProductName, RegKeyType, "A-76", 127))) ' 209[—] = 127[]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-75", 156)), Chr(GetSetting(App.ProductName, RegKeyType, "A-75", 85))) ' 156[ú] = 85[U]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-74", 136)), Chr(GetSetting(App.ProductName, RegKeyType, "A-74", 99))) ' 136[à] = 99[c]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-73", 191)), Chr(GetSetting(App.ProductName, RegKeyType, "A-73", 112))) ' 191[ø] = 112[p]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-72", 182)), Chr(GetSetting(App.ProductName, RegKeyType, "A-72", 109))) ' 182[∂] = 109[m]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-71", 146)), Chr(GetSetting(App.ProductName, RegKeyType, "A-71", 69))) ' 146[í] = 69[E]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-70", 201)), Chr(GetSetting(App.ProductName, RegKeyType, "A-70", 107))) ' 201[…] = 107[k]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-69", 134)), Chr(GetSetting(App.ProductName, RegKeyType, "A-69", 124))) ' 134[Ü] = 124[|]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-68", 204)), Chr(GetSetting(App.ProductName, RegKeyType, "A-68", 54))) ' 204[Ã] = 54[6]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-67", 172)), Chr(GetSetting(App.ProductName, RegKeyType, "A-67", 58))) ' 172[¨] = 58[:]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-66", 185)), Chr(GetSetting(App.ProductName, RegKeyType, "A-66", 131))) ' 185[π] = 131[É]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-65", 217)), Chr(GetSetting(App.ProductName, RegKeyType, "A-65", 110))) ' 217[Ÿ] = 110[n]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-64", 170)), Chr(GetSetting(App.ProductName, RegKeyType, "A-64", 64))) ' 170[™] = 64[@]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-63", 137)), Chr(GetSetting(App.ProductName, RegKeyType, "A-63", 118))) ' 137[â] = 118[v]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-62", 183)), Chr(GetSetting(App.ProductName, RegKeyType, "A-62", 78))) ' 183[∑] = 78[N]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-61", 212)), Chr(GetSetting(App.ProductName, RegKeyType, "A-61", 115))) ' 212[‘] = 115[s]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-60", 144)), Chr(GetSetting(App.ProductName, RegKeyType, "A-60", 135))) ' 144[ê] = 135[á]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-59", 165)), Chr(GetSetting(App.ProductName, RegKeyType, "A-59", 75))) ' 165[•] = 75[K]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-58", 200)), Chr(GetSetting(App.ProductName, RegKeyType, "A-58", 89))) ' 200[»] = 89[Y]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-57", 216)), Chr(GetSetting(App.ProductName, RegKeyType, "A-57", 133))) ' 216[ÿ] = 133[Ö]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-56", 215)), Chr(GetSetting(App.ProductName, RegKeyType, "A-56", 66))) ' 215[◊] = 66[B]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-55", 155)), Chr(GetSetting(App.ProductName, RegKeyType, "A-55", 76))) ' 155[õ] = 76[L]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-54", 157)), Chr(GetSetting(App.ProductName, RegKeyType, "A-54", 96))) ' 157[ù] = 96[`]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-53", 180)), Chr(GetSetting(App.ProductName, RegKeyType, "A-53", 105))) ' 180[¥] = 105[i]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-52", 151)), Chr(GetSetting(App.ProductName, RegKeyType, "A-52", 104))) ' 151[ó] = 104[h]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-51", 205)), Chr(GetSetting(App.ProductName, RegKeyType, "A-51", 55))) ' 205[Õ] = 55[7]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-50", 186)), Chr(GetSetting(App.ProductName, RegKeyType, "A-50", 81))) ' 186[∫] = 81[Q]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-49", 214)), Chr(GetSetting(App.ProductName, RegKeyType, "A-49", 88))) ' 214[÷] = 88[X]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-48", 169)), Chr(GetSetting(App.ProductName, RegKeyType, "A-48", 57))) ' 169[©] = 57[9]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-47", 177)), Chr(GetSetting(App.ProductName, RegKeyType, "A-47", 101))) ' 177[±] = 101[e]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-46", 139)), Chr(GetSetting(App.ProductName, RegKeyType, "A-46", 79))) ' 139[ã] = 79[O]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-45", 181)), Chr(GetSetting(App.ProductName, RegKeyType, "A-45", 95))) ' 181[µ] = 95[_]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-44", 173)), Chr(GetSetting(App.ProductName, RegKeyType, "A-44", 121))) ' 173[≠] = 121[y]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-43", 189)), Chr(GetSetting(App.ProductName, RegKeyType, "A-43", 82))) ' 189[Ω] = 82[R]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-42", 152)), Chr(GetSetting(App.ProductName, RegKeyType, "A-42", 90))) ' 152[ò] = 90[Z]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-41", 166)), Chr(GetSetting(App.ProductName, RegKeyType, "A-41", 116))) ' 166[¶] = 116[t]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-40", 193)), Chr(GetSetting(App.ProductName, RegKeyType, "A-40", 98))) ' 193[¡] = 98[b]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-39", 135)), Chr(GetSetting(App.ProductName, RegKeyType, "A-39", 113))) ' 135[á] = 113[q]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-38", 199)), Chr(GetSetting(App.ProductName, RegKeyType, "A-38", 71))) ' 199[«] = 71[G]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-37", 192)), Chr(GetSetting(App.ProductName, RegKeyType, "A-37", 62))) ' 192[¿] = 62[>]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-36", 213)), Chr(GetSetting(App.ProductName, RegKeyType, "A-36", 119))) ' 213[’] = 119[w]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-35", 158)), Chr(GetSetting(App.ProductName, RegKeyType, "A-35", 102))) ' 158[û] = 102[f]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-34", 203)), Chr(GetSetting(App.ProductName, RegKeyType, "A-34", 92))) ' 203[À] = 92[\]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-33", 188)), Chr(GetSetting(App.ProductName, RegKeyType, "A-33", 77))) ' 188[º] = 77[M]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-32", 194)), Chr(GetSetting(App.ProductName, RegKeyType, "A-32", 128))) ' 194[¬] = 128[Ä]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-31", 150)), Chr(GetSetting(App.ProductName, RegKeyType, "A-31", 61))) ' 150[ñ] = 61[=]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-30", 162)), Chr(GetSetting(App.ProductName, RegKeyType, "A-30", 49))) ' 162[¢] = 49[1]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-29", 141)), Chr(GetSetting(App.ProductName, RegKeyType, "A-29", 132))) ' 141[ç] = 132[Ñ]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-28", 175)), Chr(GetSetting(App.ProductName, RegKeyType, "A-28", 130))) ' 175[Ø] = 130[Ç]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-27", 147)), Chr(GetSetting(App.ProductName, RegKeyType, "A-27", 120))) ' 147[ì] = 120[x]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-26", 196)), Chr(GetSetting(App.ProductName, RegKeyType, "A-26", 125))) ' 196[ƒ] = 125[}]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-25", 154)), Chr(GetSetting(App.ProductName, RegKeyType, "A-25", 87))) ' 154[ö] = 87[W]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-24", 140)), Chr(GetSetting(App.ProductName, RegKeyType, "A-24", 73))) ' 140[å] = 73[I]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-23", 178)), Chr(GetSetting(App.ProductName, RegKeyType, "A-23", 53))) ' 178[≤] = 53[5]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-22", 143)), Chr(GetSetting(App.ProductName, RegKeyType, "A-22", 51))) ' 143[è] = 51[3]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-21", 210)), Chr(GetSetting(App.ProductName, RegKeyType, "A-21", 83))) ' 210[“] = 83[S]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-20", 211)), Chr(GetSetting(App.ProductName, RegKeyType, "A-20", 93))) ' 211[”] = 93[]]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-19", 160)), Chr(GetSetting(App.ProductName, RegKeyType, "A-19", 122))) ' 160[†] = 122[z]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-18", 159)), Chr(GetSetting(App.ProductName, RegKeyType, "A-18", 84))) ' 159[ü] = 84[T]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-17", 168)), Chr(GetSetting(App.ProductName, RegKeyType, "A-17", 126))) ' 168[®] = 126[~]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-16", 198)), Chr(GetSetting(App.ProductName, RegKeyType, "A-16", 117))) ' 198[∆] = 117[u]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-15", 202)), Chr(GetSetting(App.ProductName, RegKeyType, "A-15", 72))) ' 202[ ] = 72[H]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-14", 138)), Chr(GetSetting(App.ProductName, RegKeyType, "A-14", 52))) ' 138[ä] = 52[4]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-13", 174)), Chr(GetSetting(App.ProductName, RegKeyType, "A-13", 56))) ' 174[Æ] = 56[8]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-12", 131)), Chr(GetSetting(App.ProductName, RegKeyType, "A-12", 63))) ' 131[É] = 63[?]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-11", 206)), Chr(GetSetting(App.ProductName, RegKeyType, "A-11", 67))) ' 206[Œ] = 67[C]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-10", 132)), Chr(GetSetting(App.ProductName, RegKeyType, "A-10", 123))) ' 132[Ñ] = 123[{]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-9", 195)), Chr(GetSetting(App.ProductName, RegKeyType, "A-9", 80))) ' 195[√] = 80[P]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-8", 133)), Chr(GetSetting(App.ProductName, RegKeyType, "A-8", 68))) ' 133[Ö] = 68[D]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-7", 164)), Chr(GetSetting(App.ProductName, RegKeyType, "A-7", 94))) ' 164[§] = 94[^]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-6", 167)), Chr(GetSetting(App.ProductName, RegKeyType, "A-6", 65))) ' 167[ß] = 65[A]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-5", 142)), Chr(GetSetting(App.ProductName, RegKeyType, "A-5", 59))) ' 142[é] = 59[;]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-4", 163)), Chr(GetSetting(App.ProductName, RegKeyType, "A-4", 74))) ' 163[£] = 74[J]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-3", 153)), Chr(GetSetting(App.ProductName, RegKeyType, "A-3", 91))) ' 153[ô] = 91[[]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-2", 149)), Chr(GetSetting(App.ProductName, RegKeyType, "A-2", 100))) ' 149[ï] = 100[d]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-1", 171)), Chr(GetSetting(App.ProductName, RegKeyType, "A-1", 134))) ' 171[´] = 134[Ü]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "B-0", 197)), Chr(GetSetting(App.ProductName, RegKeyType, "A-0", 114))) ' 197[≈] = 114[r]
    Case True
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-86", 97)), Chr(GetSetting(App.ProductName, RegKeyType, "B-86", 187))) ' 97[a] = 187[ª]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-85", 70)), Chr(GetSetting(App.ProductName, RegKeyType, "B-85", 208))) ' 70[F] = 208[–]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-84", 86)), Chr(GetSetting(App.ProductName, RegKeyType, "B-84", 176))) ' 86[V] = 176[∞]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-83", 60)), Chr(GetSetting(App.ProductName, RegKeyType, "B-83", 179))) ' 60[<] = 179[≥]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-82", 106)), Chr(GetSetting(App.ProductName, RegKeyType, "B-82", 190))) ' 106[j] = 190[æ]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-81", 103)), Chr(GetSetting(App.ProductName, RegKeyType, "B-81", 161))) ' 103[g] = 161[°]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-80", 50)), Chr(GetSetting(App.ProductName, RegKeyType, "B-80", 145))) ' 50[2] = 145[ë]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-79", 111)), Chr(GetSetting(App.ProductName, RegKeyType, "B-79", 207))) ' 111[o] = 207[œ]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-78", 108)), Chr(GetSetting(App.ProductName, RegKeyType, "B-78", 130))) ' 108[l] = 130[Ç]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-77", 129)), Chr(GetSetting(App.ProductName, RegKeyType, "B-77", 184))) ' 129[Å] = 184[∏]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-76", 127)), Chr(GetSetting(App.ProductName, RegKeyType, "B-76", 209))) ' 127[] = 209[—]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-75", 85)), Chr(GetSetting(App.ProductName, RegKeyType, "B-75", 156))) ' 85[U] = 156[ú]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-74", 99)), Chr(GetSetting(App.ProductName, RegKeyType, "B-74", 136))) ' 99[c] = 136[à]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-73", 112)), Chr(GetSetting(App.ProductName, RegKeyType, "B-73", 191))) ' 112[p] = 191[ø]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-72", 109)), Chr(GetSetting(App.ProductName, RegKeyType, "B-72", 182))) ' 109[m] = 182[∂]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-71", 69)), Chr(GetSetting(App.ProductName, RegKeyType, "B-71", 146))) ' 69[E] = 146[í]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-70", 107)), Chr(GetSetting(App.ProductName, RegKeyType, "B-70", 201))) ' 107[k] = 201[…]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-69", 124)), Chr(GetSetting(App.ProductName, RegKeyType, "B-69", 134))) ' 124[|] = 134[Ü]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-68", 54)), Chr(GetSetting(App.ProductName, RegKeyType, "B-68", 204))) ' 54[6] = 204[Ã]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-67", 58)), Chr(GetSetting(App.ProductName, RegKeyType, "B-67", 172))) ' 58[:] = 172[¨]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-66", 131)), Chr(GetSetting(App.ProductName, RegKeyType, "B-66", 185))) ' 131[É] = 185[π]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-65", 110)), Chr(GetSetting(App.ProductName, RegKeyType, "B-65", 217))) ' 110[n] = 217[Ÿ]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-64", 64)), Chr(GetSetting(App.ProductName, RegKeyType, "B-64", 170))) ' 64[@] = 170[™]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-63", 118)), Chr(GetSetting(App.ProductName, RegKeyType, "B-63", 137))) ' 118[v] = 137[â]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-62", 78)), Chr(GetSetting(App.ProductName, RegKeyType, "B-62", 183))) ' 78[N] = 183[∑]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-61", 115)), Chr(GetSetting(App.ProductName, RegKeyType, "B-61", 212))) ' 115[s] = 212[‘]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-60", 135)), Chr(GetSetting(App.ProductName, RegKeyType, "B-60", 144))) ' 135[á] = 144[ê]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-59", 75)), Chr(GetSetting(App.ProductName, RegKeyType, "B-59", 165))) ' 75[K] = 165[•]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-58", 89)), Chr(GetSetting(App.ProductName, RegKeyType, "B-58", 200))) ' 89[Y] = 200[»]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-57", 133)), Chr(GetSetting(App.ProductName, RegKeyType, "B-57", 216))) ' 133[Ö] = 216[ÿ]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-56", 66)), Chr(GetSetting(App.ProductName, RegKeyType, "B-56", 215))) ' 66[B] = 215[◊]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-55", 76)), Chr(GetSetting(App.ProductName, RegKeyType, "B-55", 155))) ' 76[L] = 155[õ]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-54", 96)), Chr(GetSetting(App.ProductName, RegKeyType, "B-54", 157))) ' 96[`] = 157[ù]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-53", 105)), Chr(GetSetting(App.ProductName, RegKeyType, "B-53", 180))) ' 105[i] = 180[¥]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-52", 104)), Chr(GetSetting(App.ProductName, RegKeyType, "B-52", 151))) ' 104[h] = 151[ó]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-51", 55)), Chr(GetSetting(App.ProductName, RegKeyType, "B-51", 205))) ' 55[7] = 205[Õ]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-50", 81)), Chr(GetSetting(App.ProductName, RegKeyType, "B-50", 186))) ' 81[Q] = 186[∫]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-49", 88)), Chr(GetSetting(App.ProductName, RegKeyType, "B-49", 214))) ' 88[X] = 214[÷]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-48", 57)), Chr(GetSetting(App.ProductName, RegKeyType, "B-48", 169))) ' 57[9] = 169[©]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-47", 101)), Chr(GetSetting(App.ProductName, RegKeyType, "B-47", 177))) ' 101[e] = 177[±]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-46", 79)), Chr(GetSetting(App.ProductName, RegKeyType, "B-46", 139))) ' 79[O] = 139[ã]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-45", 95)), Chr(GetSetting(App.ProductName, RegKeyType, "B-45", 181))) ' 95[_] = 181[µ]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-44", 121)), Chr(GetSetting(App.ProductName, RegKeyType, "B-44", 173))) ' 121[y] = 173[≠]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-43", 82)), Chr(GetSetting(App.ProductName, RegKeyType, "B-43", 189))) ' 82[R] = 189[Ω]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-42", 90)), Chr(GetSetting(App.ProductName, RegKeyType, "B-42", 152))) ' 90[Z] = 152[ò]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-41", 116)), Chr(GetSetting(App.ProductName, RegKeyType, "B-41", 166))) ' 116[t] = 166[¶]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-40", 98)), Chr(GetSetting(App.ProductName, RegKeyType, "B-40", 193))) ' 98[b] = 193[¡]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-39", 113)), Chr(GetSetting(App.ProductName, RegKeyType, "B-39", 135))) ' 113[q] = 135[á]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-38", 71)), Chr(GetSetting(App.ProductName, RegKeyType, "B-38", 199))) ' 71[G] = 199[«]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-37", 62)), Chr(GetSetting(App.ProductName, RegKeyType, "B-37", 192))) ' 62[>] = 192[¿]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-36", 119)), Chr(GetSetting(App.ProductName, RegKeyType, "B-36", 213))) ' 119[w] = 213[’]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-35", 102)), Chr(GetSetting(App.ProductName, RegKeyType, "B-35", 158))) ' 102[f] = 158[û]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-34", 92)), Chr(GetSetting(App.ProductName, RegKeyType, "B-34", 203))) ' 92[\] = 203[À]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-33", 77)), Chr(GetSetting(App.ProductName, RegKeyType, "B-33", 188))) ' 77[M] = 188[º]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-32", 128)), Chr(GetSetting(App.ProductName, RegKeyType, "B-32", 194))) ' 128[Ä] = 194[¬]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-31", 61)), Chr(GetSetting(App.ProductName, RegKeyType, "B-31", 150))) ' 61[=] = 150[ñ]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-30", 49)), Chr(GetSetting(App.ProductName, RegKeyType, "B-30", 162))) ' 49[1] = 162[¢]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-29", 132)), Chr(GetSetting(App.ProductName, RegKeyType, "B-29", 141))) ' 132[Ñ] = 141[ç]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-28", 130)), Chr(GetSetting(App.ProductName, RegKeyType, "B-28", 175))) ' 130[Ç] = 175[Ø]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-27", 120)), Chr(GetSetting(App.ProductName, RegKeyType, "B-27", 147))) ' 120[x] = 147[ì]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-26", 125)), Chr(GetSetting(App.ProductName, RegKeyType, "B-26", 196))) ' 125[}] = 196[ƒ]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-25", 87)), Chr(GetSetting(App.ProductName, RegKeyType, "B-25", 154))) ' 87[W] = 154[ö]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-24", 73)), Chr(GetSetting(App.ProductName, RegKeyType, "B-24", 140))) ' 73[I] = 140[å]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-23", 53)), Chr(GetSetting(App.ProductName, RegKeyType, "B-23", 178))) ' 53[5] = 178[≤]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-22", 51)), Chr(GetSetting(App.ProductName, RegKeyType, "B-22", 143))) ' 51[3] = 143[è]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-21", 83)), Chr(GetSetting(App.ProductName, RegKeyType, "B-21", 210))) ' 83[S] = 210[“]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-20", 93)), Chr(GetSetting(App.ProductName, RegKeyType, "B-20", 211))) ' 93[]] = 211[”]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-19", 122)), Chr(GetSetting(App.ProductName, RegKeyType, "B-19", 160))) ' 122[z] = 160[†]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-18", 84)), Chr(GetSetting(App.ProductName, RegKeyType, "B-18", 159))) ' 84[T] = 159[ü]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-17", 126)), Chr(GetSetting(App.ProductName, RegKeyType, "B-17", 168))) ' 126[~] = 168[®]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-16", 117)), Chr(GetSetting(App.ProductName, RegKeyType, "B-16", 198))) ' 117[u] = 198[∆]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-15", 72)), Chr(GetSetting(App.ProductName, RegKeyType, "B-15", 202))) ' 72[H] = 202[ ]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-14", 52)), Chr(GetSetting(App.ProductName, RegKeyType, "B-14", 138))) ' 52[4] = 138[ä]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-13", 56)), Chr(GetSetting(App.ProductName, RegKeyType, "B-13", 174))) ' 56[8] = 174[Æ]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-12", 63)), Chr(GetSetting(App.ProductName, RegKeyType, "B-12", 131))) ' 63[?] = 131[É]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-11", 67)), Chr(GetSetting(App.ProductName, RegKeyType, "B-11", 206))) ' 67[C] = 206[Œ]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-10", 123)), Chr(GetSetting(App.ProductName, RegKeyType, "B-10", 132))) ' 123[{] = 132[Ñ]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-9", 80)), Chr(GetSetting(App.ProductName, RegKeyType, "B-9", 195))) ' 80[P] = 195[√]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-8", 68)), Chr(GetSetting(App.ProductName, RegKeyType, "B-8", 133))) ' 68[D] = 133[Ö]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-7", 94)), Chr(GetSetting(App.ProductName, RegKeyType, "B-7", 164))) ' 94[^] = 164[§]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-6", 65)), Chr(GetSetting(App.ProductName, RegKeyType, "B-6", 167))) ' 65[A] = 167[ß]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-5", 59)), Chr(GetSetting(App.ProductName, RegKeyType, "B-5", 142))) ' 59[;] = 142[é]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-4", 74)), Chr(GetSetting(App.ProductName, RegKeyType, "B-4", 163))) ' 74[J] = 163[£]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-3", 91)), Chr(GetSetting(App.ProductName, RegKeyType, "B-3", 153))) ' 91[[] = 153[ô]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-2", 100)), Chr(GetSetting(App.ProductName, RegKeyType, "B-2", 149))) ' 100[d] = 149[ï]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-1", 134)), Chr(GetSetting(App.ProductName, RegKeyType, "B-1", 171))) ' 134[Ü] = 171[´]
        txt = MySQL.ReplaceString(txt, Chr(GetSetting(App.ProductName, RegKeyType, "A-0", 114)), Chr(GetSetting(App.ProductName, RegKeyType, "B-0", 197))) ' 114[r] = 197[≈]
    End Select

    Crypt = txt
End Function

Function XMLTag(start As Integer, txt As String, Tag As String) As String

    Dim iPos As Integer
    Dim iPosb As Integer

    iPos = InStr(start, txt, "<" & Tag & ">")
    iPosb = InStr(start, txt, "</" & Tag & ">")
    If iPos > 0 And iPosb > 0 And iPos < iPosb Then
    
        XMLTag = Mid(txt, iPos + Len("<" & Tag & ">"), iPosb - iPos - Len("<" & Tag & ">"))
    
    End If
    
End Function

Function pass(OldPassword As String, Optional Capt As String) As String

    If Capt = "" Then
        If OldPassword = "" Then
            Capt = "Please enter in a new password."
        Else
            Capt = "Verify And Change the Password"
        End If
    End If
    
    Dim fPas As New frmPWDChanger
    fPas.PassWRD = OldPassword
    fPas.Capt = Capt
    fPas.Show 1
    If fPas.PassWRD <> OldPassword Then
        pass = fPas.PassWRD
    Else
        pass = OldPassword
    End If
    
End Function

Function GetSessionChar(ByRef SESSIONCHAR As String, fhWnd As Long, Optional ilen As Byte = 14) As String

    Dim iPos As Byte
    Dim buf As String
    
    Randomize fhWnd
        
    For ilen = 1 To ilen
        buf = buf & Chr(Round(Rnd * 200) + 55)
    Next
    
    SESSIONCHAR = buf
    GetSessionChar = buf
End Function

Function NullStr(ByVal obj As Variant, Optional NumStr As NumberorString = nString) As Variant

    If IsNull(obj) = True Then
        NullStr = IIf(NumStr = nNumber, 0, "")
    Else
        Set NullStr = obj
    End If
    
End Function

Public Sub gSleep(Optional lMss As Integer = 0)

    ' This routine replaces the gSleep.
    ' The Doevents will run away/pause when a seperate thread exists.
    Static NumOfCalls As Byte
    
    NumOfCalls = NumOfCalls + 1
    
    If lMss = 0 Then
        lMss = Round(Rnd * reg.SleepForMS)
    End If
    
    If NumOfCalls < reg.Sleep2Doevents Then
        NumOfCalls = NumOfCalls + 1
        Call SleepEx(lMss, True)
    ElseIf NumOfCalls < reg.Sleep2Doevents + 1 Then
        NumOfCalls = NumOfCalls + 1
        DoEvents
    Else
        NumOfCalls = 0
        DoEvents
    End If
        
    
End Sub
Public Function IsIcon(Optional ByVal lIconNum As Long = -1, Optional ByVal sIconKey As String = "blank", Optional ByRef pIcon16x16 As IPictureDisp, Optional ByRef pIcon32x32 As IPictureDisp) As Boolean

    IsThere = False
    
    If fIcon.il16x16.ListImages.Count > 1 Then
        If lIconNum > 0 Then
            If fIcon.il16x16.ListImages.Count >= lIconNum And fIcon.il16x16.ListImages.Count <= lIconNum Then
                lIconNum = lIconNum
                IsThere = True
            Else
                lIconNum = 1
            End If
            
            Set pIcon16x16 = fIcon.il16x16.ListImages(lIconNum).ExtractIcon
            Set pIcon32x32 = fIcon.il32x32.ListImages(lIconNum).ExtractIcon
            
        ElseIf Len(sIconKey) > 0 Then
            
            Dim hImage As ListImage
            Dim bFound As Boolean
            
            bFound = False
            
            For Each hImage In fIcon.il16x16.ListImages
                If hImage.Key = sIconKey Then
                    bFound = True
                    Exit For
                End If
            Next
            
            If bFound = True Then
                sIconKey = sIconKey
                IsThere = True
            Else
                sIconKey = "blank"
            End If
            
            Set pIcon16x16 = fIcon.il16x16.ListImages(sIconKey).ExtractIcon
            Set pIcon32x32 = fIcon.il32x32.ListImages(sIconKey).ExtractIcon
        Else
            lIconNum = 1
            sIconKey = "blank"
            Set pIcon16x16 = fIcon.il16x16.ListImages(sIconKey).ExtractIcon
            Set pIcon32x32 = fIcon.il32x32.ListImages(sIconKey).ExtractIcon
        End If
    Else
        lIconNum = 1
        sIconKey = "blank"
        Set pIcon16x16 = fIcon.il16x16.ListImages(sIconKey).Picture
        Set pIcon32x32 = fIcon.il32x32.ListImages(sIconKey).Picture
    End If
    
        
    
End Function
