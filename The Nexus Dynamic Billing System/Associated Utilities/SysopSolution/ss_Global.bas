Attribute VB_Name = "Module1"
Global oConn As ADODB.Connection
Global ViSPMAP As New mapVirtualISP

Type Login_Details
    
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

Global oErr As New clsErrors


Global sTax As Single

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


'Global fLine As frmLines
'Global Radius As New colRadius

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
    IpAddress As String * 16
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

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const Flags = SWP_NOMOVE Or SWP_NOSIZE

Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2

Global oMySQL As New clsMySQL
Global oGUI As New clsInterface
Global cErr As New clsErrors

Global sServer As String
Global sUID As String
Global sPWD As String


Type tySchedule
    iNumUserFiles As Integer
    iNumRadius As Integer
End Type

Global Schedule As tySchedule


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

'Global reg As Settings

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

Global DirName As String

Global Const PasswordSalt = "dr34mt1me"

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Global Point As POINTAPI
Global ConnStr As String

Private Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub ShellLauncher()

    Dim path As String, extension As String
    
    path = IIf(Right(App.path, 1) = "\", App.path, App.path + "\")
    
            SaveSetting "projectalpha", "db", "ConnectionString", ""
    
    If Dir(path + "dblauncher.exe", vbNormal) <> "" Then
            
            
            Shell path + "dblauncher.exe " + extension, vbNormalFocus
            
    End If
    
        
End Sub
Public Function Crypt(txt As String, bEncrypt As Boolean, RegKeyType As String, PrimaryKey As String) As String

    Select Case bEncrypt
    Case False
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-86", 187)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-86", 97))) ' 187[ª] = 97[a]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-85", 208)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-85", 70))) ' 208[–] = 70[F]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-84", 176)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-84", 86))) ' 176[∞] = 86[V]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-83", 179)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-83", 60))) ' 179[≥] = 60[<]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-82", 190)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-82", 106))) ' 190[æ] = 106[j]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-81", 161)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-81", 103))) ' 161[°] = 103[g]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-80", 145)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-80", 50))) ' 145[ë] = 50[2]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-79", 207)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-79", 111))) ' 207[œ] = 111[o]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-78", 130)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-78", 108))) ' 130[Ç] = 108[l]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-77", 184)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-77", 129))) ' 184[∏] = 129[Å]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-76", 209)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-76", 127))) ' 209[—] = 127[]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-75", 156)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-75", 85))) ' 156[ú] = 85[U]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-74", 136)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-74", 99))) ' 136[à] = 99[c]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-73", 191)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-73", 112))) ' 191[ø] = 112[p]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-72", 182)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-72", 109))) ' 182[∂] = 109[m]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-71", 146)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-71", 69))) ' 146[í] = 69[E]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-70", 201)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-70", 107))) ' 201[…] = 107[k]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-69", 134)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-69", 124))) ' 134[Ü] = 124[|]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-68", 204)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-68", 54))) ' 204[Ã] = 54[6]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-67", 172)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-67", 58))) ' 172[¨] = 58[:]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-66", 185)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-66", 131))) ' 185[π] = 131[É]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-65", 217)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-65", 110))) ' 217[Ÿ] = 110[n]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-64", 170)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-64", 64))) ' 170[™] = 64[@]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-63", 137)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-63", 118))) ' 137[â] = 118[v]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-62", 183)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-62", 78))) ' 183[∑] = 78[N]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-61", 212)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-61", 115))) ' 212[‘] = 115[s]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-60", 144)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-60", 135))) ' 144[ê] = 135[á]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-59", 165)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-59", 75))) ' 165[•] = 75[K]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-58", 200)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-58", 89))) ' 200[»] = 89[Y]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-57", 216)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-57", 133))) ' 216[ÿ] = 133[Ö]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-56", 215)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-56", 66))) ' 215[◊] = 66[B]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-55", 155)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-55", 76))) ' 155[õ] = 76[L]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-54", 157)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-54", 96))) ' 157[ù] = 96[`]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-53", 180)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-53", 105))) ' 180[¥] = 105[i]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-52", 151)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-52", 104))) ' 151[ó] = 104[h]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-51", 205)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-51", 55))) ' 205[Õ] = 55[7]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-50", 186)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-50", 81))) ' 186[∫] = 81[Q]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-49", 214)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-49", 88))) ' 214[÷] = 88[X]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-48", 169)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-48", 57))) ' 169[©] = 57[9]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-47", 177)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-47", 101))) ' 177[±] = 101[e]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-46", 139)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-46", 79))) ' 139[ã] = 79[O]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-45", 181)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-45", 95))) ' 181[µ] = 95[_]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-44", 173)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-44", 121))) ' 173[≠] = 121[y]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-43", 189)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-43", 82))) ' 189[Ω] = 82[R]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-42", 152)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-42", 90))) ' 152[ò] = 90[Z]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-41", 166)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-41", 116))) ' 166[¶] = 116[t]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-40", 193)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-40", 98))) ' 193[¡] = 98[b]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-39", 135)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-39", 113))) ' 135[á] = 113[q]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-38", 199)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-38", 71))) ' 199[«] = 71[G]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-37", 192)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-37", 62))) ' 192[¿] = 62[>]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-36", 213)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-36", 119))) ' 213[’] = 119[w]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-35", 158)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-35", 102))) ' 158[û] = 102[f]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-34", 203)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-34", 92))) ' 203[À] = 92[\]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-33", 188)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-33", 77))) ' 188[º] = 77[M]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-32", 194)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-32", 128))) ' 194[¬] = 128[Ä]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-31", 150)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-31", 61))) ' 150[ñ] = 61[=]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-30", 162)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-30", 49))) ' 162[¢] = 49[1]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-29", 141)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-29", 132))) ' 141[ç] = 132[Ñ]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-28", 175)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-28", 130))) ' 175[Ø] = 130[Ç]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-27", 147)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-27", 120))) ' 147[ì] = 120[x]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-26", 196)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-26", 125))) ' 196[ƒ] = 125[}]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-25", 154)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-25", 87))) ' 154[ö] = 87[W]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-24", 140)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-24", 73))) ' 140[å] = 73[I]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-23", 178)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-23", 53))) ' 178[≤] = 53[5]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-22", 143)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-22", 51))) ' 143[è] = 51[3]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-21", 210)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-21", 83))) ' 210[“] = 83[S]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-20", 211)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-20", 93))) ' 211[”] = 93[]]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-19", 160)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-19", 122))) ' 160[†] = 122[z]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-18", 159)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-18", 84))) ' 159[ü] = 84[T]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-17", 168)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-17", 126))) ' 168[®] = 126[~]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-16", 198)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-16", 117))) ' 198[∆] = 117[u]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-15", 202)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-15", 72))) ' 202[ ] = 72[H]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-14", 138)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-14", 52))) ' 138[ä] = 52[4]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-13", 174)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-13", 56))) ' 174[Æ] = 56[8]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-12", 131)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-12", 63))) ' 131[É] = 63[?]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-11", 206)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-11", 67))) ' 206[Œ] = 67[C]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-10", 132)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-10", 123))) ' 132[Ñ] = 123[{]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-9", 195)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-9", 80))) ' 195[√] = 80[P]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-8", 133)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-8", 68))) ' 133[Ö] = 68[D]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-7", 164)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-7", 94))) ' 164[§] = 94[^]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-6", 167)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-6", 65))) ' 167[ß] = 65[A]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-5", 142)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-5", 59))) ' 142[é] = 59[;]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-4", 163)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-4", 74))) ' 163[£] = 74[J]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-3", 153)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-3", 91))) ' 153[ô] = 91[[]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-2", 149)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-2", 100))) ' 149[ï] = 100[d]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-1", 171)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-1", 134))) ' 171[´] = 134[Ü]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "B-0", 197)), Chr(GetSetting(PrimaryKey, RegKeyType, "A-0", 114))) ' 197[≈] = 114[r]
    Case True
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-86", 97)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-86", 187))) ' 97[a] = 187[ª]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-85", 70)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-85", 208))) ' 70[F] = 208[–]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-84", 86)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-84", 176))) ' 86[V] = 176[∞]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-83", 60)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-83", 179))) ' 60[<] = 179[≥]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-82", 106)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-82", 190))) ' 106[j] = 190[æ]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-81", 103)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-81", 161))) ' 103[g] = 161[°]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-80", 50)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-80", 145))) ' 50[2] = 145[ë]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-79", 111)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-79", 207))) ' 111[o] = 207[œ]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-78", 108)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-78", 130))) ' 108[l] = 130[Ç]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-77", 129)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-77", 184))) ' 129[Å] = 184[∏]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-76", 127)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-76", 209))) ' 127[] = 209[—]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-75", 85)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-75", 156))) ' 85[U] = 156[ú]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-74", 99)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-74", 136))) ' 99[c] = 136[à]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-73", 112)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-73", 191))) ' 112[p] = 191[ø]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-72", 109)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-72", 182))) ' 109[m] = 182[∂]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-71", 69)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-71", 146))) ' 69[E] = 146[í]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-70", 107)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-70", 201))) ' 107[k] = 201[…]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-69", 124)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-69", 134))) ' 124[|] = 134[Ü]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-68", 54)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-68", 204))) ' 54[6] = 204[Ã]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-67", 58)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-67", 172))) ' 58[:] = 172[¨]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-66", 131)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-66", 185))) ' 131[É] = 185[π]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-65", 110)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-65", 217))) ' 110[n] = 217[Ÿ]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-64", 64)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-64", 170))) ' 64[@] = 170[™]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-63", 118)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-63", 137))) ' 118[v] = 137[â]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-62", 78)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-62", 183))) ' 78[N] = 183[∑]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-61", 115)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-61", 212))) ' 115[s] = 212[‘]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-60", 135)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-60", 144))) ' 135[á] = 144[ê]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-59", 75)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-59", 165))) ' 75[K] = 165[•]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-58", 89)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-58", 200))) ' 89[Y] = 200[»]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-57", 133)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-57", 216))) ' 133[Ö] = 216[ÿ]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-56", 66)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-56", 215))) ' 66[B] = 215[◊]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-55", 76)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-55", 155))) ' 76[L] = 155[õ]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-54", 96)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-54", 157))) ' 96[`] = 157[ù]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-53", 105)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-53", 180))) ' 105[i] = 180[¥]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-52", 104)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-52", 151))) ' 104[h] = 151[ó]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-51", 55)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-51", 205))) ' 55[7] = 205[Õ]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-50", 81)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-50", 186))) ' 81[Q] = 186[∫]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-49", 88)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-49", 214))) ' 88[X] = 214[÷]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-48", 57)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-48", 169))) ' 57[9] = 169[©]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-47", 101)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-47", 177))) ' 101[e] = 177[±]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-46", 79)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-46", 139))) ' 79[O] = 139[ã]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-45", 95)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-45", 181))) ' 95[_] = 181[µ]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-44", 121)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-44", 173))) ' 121[y] = 173[≠]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-43", 82)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-43", 189))) ' 82[R] = 189[Ω]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-42", 90)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-42", 152))) ' 90[Z] = 152[ò]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-41", 116)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-41", 166))) ' 116[t] = 166[¶]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-40", 98)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-40", 193))) ' 98[b] = 193[¡]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-39", 113)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-39", 135))) ' 113[q] = 135[á]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-38", 71)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-38", 199))) ' 71[G] = 199[«]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-37", 62)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-37", 192))) ' 62[>] = 192[¿]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-36", 119)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-36", 213))) ' 119[w] = 213[’]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-35", 102)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-35", 158))) ' 102[f] = 158[û]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-34", 92)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-34", 203))) ' 92[\] = 203[À]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-33", 77)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-33", 188))) ' 77[M] = 188[º]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-32", 128)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-32", 194))) ' 128[Ä] = 194[¬]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-31", 61)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-31", 150))) ' 61[=] = 150[ñ]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-30", 49)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-30", 162))) ' 49[1] = 162[¢]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-29", 132)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-29", 141))) ' 132[Ñ] = 141[ç]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-28", 130)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-28", 175))) ' 130[Ç] = 175[Ø]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-27", 120)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-27", 147))) ' 120[x] = 147[ì]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-26", 125)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-26", 196))) ' 125[}] = 196[ƒ]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-25", 87)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-25", 154))) ' 87[W] = 154[ö]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-24", 73)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-24", 140))) ' 73[I] = 140[å]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-23", 53)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-23", 178))) ' 53[5] = 178[≤]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-22", 51)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-22", 143))) ' 51[3] = 143[è]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-21", 83)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-21", 210))) ' 83[S] = 210[“]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-20", 93)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-20", 211))) ' 93[]] = 211[”]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-19", 122)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-19", 160))) ' 122[z] = 160[†]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-18", 84)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-18", 159))) ' 84[T] = 159[ü]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-17", 126)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-17", 168))) ' 126[~] = 168[®]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-16", 117)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-16", 198))) ' 117[u] = 198[∆]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-15", 72)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-15", 202))) ' 72[H] = 202[ ]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-14", 52)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-14", 138))) ' 52[4] = 138[ä]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-13", 56)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-13", 174))) ' 56[8] = 174[Æ]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-12", 63)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-12", 131))) ' 63[?] = 131[É]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-11", 67)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-11", 206))) ' 67[C] = 206[Œ]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-10", 123)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-10", 132))) ' 123[{] = 132[Ñ]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-9", 80)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-9", 195))) ' 80[P] = 195[√]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-8", 68)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-8", 133))) ' 68[D] = 133[Ö]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-7", 94)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-7", 164))) ' 94[^] = 164[§]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-6", 65)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-6", 167))) ' 65[A] = 167[ß]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-5", 59)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-5", 142))) ' 59[;] = 142[é]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-4", 74)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-4", 163))) ' 74[J] = 163[£]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-3", 91)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-3", 153))) ' 91[[] = 153[ô]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-2", 100)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-2", 149))) ' 100[d] = 149[ï]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-1", 134)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-1", 171))) ' 134[Ü] = 171[´]
        txt = oMySQL.ReplaceString(txt, Chr(GetSetting(PrimaryKey, RegKeyType, "A-0", 114)), Chr(GetSetting(PrimaryKey, RegKeyType, "B-0", 197))) ' 114[r] = 197[≈]
    End Select

    Crypt = txt
End Function


Public Function cDebug(ByVal txt As String)

    Static iLineCount As Variant
    If iLineCount = 2147483600 Then iLineCount = 0
    iLineCount = iLineCount + 1
    Dim I As Long
    
    If InStr(txt, "AES_ENCRYPT") > 0 Then
        txt = Left(txt, InStr(txt, "AES_ENCRYPT")) + "AES_ENCRYPT()" + Mid(txt, InStr(InStr(txt, "AES_ENCRYPT") + 1, txt, ")"))
    End If
    
    'txt = oMySQL.ReplaceString(txt, CCSalt, "salt")
    'txt = oMySQL.ReplaceString(txt, md5Password, "salt")
    txt = oMySQL.ReplaceString(txt, PasswordSalt, "salt")
    'txt = oMySQL.ReplaceString(txt, "386bba5a5dc4fac215c9cf0b9a29b352", "salt")
    
    'frmDebug.txtDebug.SelStart = Len(frmDebug.txtDebug)
    I = Len(frmDebug.txtDebug)
    frmDebug.txtDebug = frmDebug.txtDebug + "&H" & String(8 - Len(Hex(iLineCount)), "0") + Hex(iLineCount) + "       " + txt + vbCrLf
    Debug.Print "&H" & String(8 - Len(Hex(iLineCount)), "0") + Hex(iLineCount) + "       " + txt
    frmDebug.txtDebug.SelStart = I
    
    If iLineCount / 200 = Round(iLineCount / 200) Then
        If frmMDIMain_Loaded = True Then frmMDIMain.txtDebug.Text = Mid(frmMDIMain.txtDebug.Text, InStr(frmMDIMain.txtDebug.Text, "&H" & String(8 - Len(Hex(iLineCount)), "0") + Hex(iLineCount)))
        frmDebug.txtDebug = frmMDIMain.txtDebug.Text
    End If
    
End Function

Sub gSleep(Optional ByVal mssnooze As Long = 0)
    
    mssnooze = Rnd * IIf(mssmooth = 0, 30, mssmooth)
    Call SleepEx(mssnooze, False)
    
End Sub


Public Function XMLVal(XML As String, XMLTag As String) As String

    If InStr(XML, XMLTag) = 0 Then Exit Function
    
    Dim iPosA As Integer
    Dim iPosb As Integer
    
    iPosA = InStr(XML, "<" & XMLTag & ">") + Len("<" & XMLTag & ">")
    iPosb = InStr(XML, IIf(InStr(XML, "</" & XMLTag & ">") > 0, "</" & XMLTag & ">", "<" & XMLTag & "/>"))
    
    XMLVal = Mid(XML, iPosA, iPosb - iPosA)
    
    
End Function
Function CloneSysop(oConn As ADODB.Connection, dSysopID As Long, NewUsername As String, NewPassword As String, NewDescription As String) As Long


    Dim rsIn As ADODB.Recordset
    Dim SQL As String
    Dim GoOnLogin As Boolean
    
    On Error Resume Next
    Do
        Err.Clear
        CloneSysop = MySQL.GetTMPRecID("sysops", oConn, "RecID")
        oConn.Execute "insert into sysops (RecID) VALUES ('" & CloneSysop & "')"
        gSleep
    Loop Until Err.Number = 0
    
    GoOnLogin = True
    
    If MySQL.OpenTable(oConn, rsIn, , "select * from sysops where RecID = '" & dSysopID & "'", adOpenStatic, adLockOptimistic) = True Then
        If rsIn.State = adStateOpen Then
            If rsIn.BOF And rsIn.EOF Then
                GoOnLogin = True
            Else
                If rsIn.RecordCount > 0 Then
                    SQL = SQL & "update sysops set " & _
                                ",Username = '" & NewUsername & "'" & _
                                ",Password = AES_ENCYPT('" & MySQL.ESC(NewPassword) & "','" & odb.colSalts.ReturnSalt("PasswordSalt") & "')" & _
                                ",Description = '" & MySQL.ESC(NewDescription) & "'" & _
                                ",Checked = '" & Val(rsIn!Checked) & "'" & _
                                ",SecurityLevel = '" & Val(rsIn!SecurityLevel) & "'" & _
                                ",VirtualID = '" & Val(rsIn!VirtualID) & "'" & _
                                ",AgencyID = '" & Val(rsIn!AgencyID) & "'" & _
                                ",bMaintain = '" & Val(rsIn!bMaintain) & "'" & _
                                ",bVISP = '" & Val(rsIn!bVISP) & "'" & _
                                ",IncomeTax = '" & Val(rsIn!IncomeTax) & "'" & _
                                ",SuperRate = '" & Val(rsIn!SuperRate) & "'" & _
                                ",bVISPFiscal = '" & Val(rsIn!bVISPFiscal) & "'" & _
                                ",PerVISP = '" & Val(rsIn!PerVISP) & "'" & _
                                ",bAgency = '" & Val(rsIn!bAgency) & "'" & _
                                ",bCreateSysop = '" & Val(rsIn!bCreateSysop) & "'" & _
                                ",bPrimary = '" & Val(rsIn!bPrimary) & "'" & _
                                ",bTemplates = '" & Val(rsIn!bTemplates) & "'" & _
                                ",NextCycle = '" & Format(DateAdd("d", 2, sysnow), "yyyy-mm-dd ttttt") & "'" & _
                                ",PreviousCycle = '" & Format(DateAdd("m", -1, sysnow), "yyyy-mm-dd ttttt") & "'" & _
                                ",bInvoice = '" & Val(rsIn!bInvoice) & "'"
                    SQL = SQL & ",bRecievables = '" & Val(rsIn!bRecievables) & "'" & _
                                ",bExpenditure = '" & Val(rsIn!bExpenditure) & "'" & _
                                ",bHoldings = '" & Val(rsIn!bHoldings) & "'" & _
                                ",bComm = '" & Val(rsIn!bComm) & "'" & _
                                ",bRefund = '" & Val(rsIn!bRefund) & "'" & _
                                ",bAddCust = '" & Val(rsIn!bAddCust) & "'" & _
                                ",bOwnership = '" & Val(rsIn!bOwnership) & "'" & _
                                ",bAccSettings = '" & Val(rsIn!bAccSettings) & "'" & _
                                ",bVendors = '" & Val(rsIn!bVendors) & "'" & _
                                ",PublicKey = '" & MySQL.ESC(rsIn!PublicKey) & "'" & _
                                ",bWEBAccount = '" & Val(rsIn!bWEBAccount) & "'"

                    GoOnLogin = False
                Else
                    GoOnLogin = True
                End If
            End If
        Else
            GoOnLogin = True
        End If
    
    End If
    
    If GoOnLogin = True Then
        SQL = SQL & "update sysops set " & _
                     ",Username = '" & NewUsername & "'" & _
                     ",Password = AES_ENCYPT('" & MySQL.ESC(NewPassword) & "','" & odb.colSalts.ReturnSalt("PasswordSalt") & "')" & _
                     ",Description = '" & MySQL.ESC(NewDescription) & "'" & _
                    ",Checked = '" & "-1" & "'" & _
                    ",SecurityLevel = '" & Login.lLevel & "'" & _
                    ",VirtualID = '" & Login.lVirtualID & "'" & _
                    ",AgencyID = '" & Login.lAgencyID & "'" & _
                    ",bMaintain = '" & Login.bRunMaintenance & "'" & _
                    ",bVISP = '" & Login.bVISP & "'" & _
                    ",bVISPFiscal = '" & Login.bVISPFiscal & "'" & _
                    ",bAgency = '" & Login.bAgency & "'" & _
                    ",bCreateSysop = '" & Login.bCreateSysop & "'" & _
                    ",bPrimary = '" & Login.bPrimary & "'" & _
                    ",bTemplates = '" & Login.bTemplates & "'" & _
                    ",NextCycle = '" & Format(DateAdd("d", 2, sysnow), "yyyy-mm-dd ttttt") & "'" & _
                    ",PreviousCycle = '" & Format(DateAdd("m", -1, sysnow), "yyyy-mm-dd ttttt") & "'" & _
                    ",bInvoice = '" & Login.bInvoice & "'"
        SQL = SQL & ",bRecievables = '" & Login.bRecievables & "'" & _
                    ",bExpenditure = '" & Login.bExpenditure & "'" & _
                    ",bHoldings = '" & Login.bHoldings & "'" & _
                    ",bComm = '" & Login.bComm & "'" & _
                    ",bRefund = '" & Login.bRefund & "'" & _
                    ",bAddCust = '" & Login.bAddCust & "'" & _
                    ",bOwnership = '" & Login.bOwnership & "'" & _
                    ",bAccSettings = '" & Login.bAccSettings & "'" & _
                    ",bVendors = '" & Login.bVendors & "'" & _
                    ",PublicKey = '" & Login.PublicKey & "'" & _
                    ",bWEBAccount = '0'"

        End If
        
    Call MySQL.Execute(oConn, SQL & " where RecID = " & CloneSysop)
    
End Function


