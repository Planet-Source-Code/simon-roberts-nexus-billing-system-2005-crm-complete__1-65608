VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   3435
   ClientLeft      =   7440
   ClientTop       =   4725
   ClientWidth     =   3900
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   3900
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1740
      Top             =   1500
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close and Exit"
      Height          =   315
      Left            =   2370
      TabIndex        =   7
      Top             =   3030
      Width           =   1395
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      Default         =   -1  'True
      Height          =   315
      Left            =   150
      TabIndex        =   6
      Top             =   3060
      Width           =   1395
   End
   Begin VB.Frame Frame2 
      Caption         =   "&Password"
      Height          =   1065
      Left            =   150
      TabIndex        =   3
      Top             =   1860
      Width           =   3615
      Begin VB.CheckBox chkRemember 
         Caption         =   "&Remember Password"
         Height          =   285
         Left            =   150
         TabIndex        =   5
         Top             =   690
         Width           =   3315
      End
      Begin VB.TextBox txtPassword 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   150
         PasswordChar    =   "#"
         TabIndex        =   4
         Top             =   300
         Width           =   3345
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Username"
      Height          =   795
      Left            =   150
      TabIndex        =   1
      Top             =   990
      Width           =   3615
      Begin VB.ComboBox cmbUsername 
         Height          =   360
         Left            =   150
         TabIndex        =   2
         Top             =   300
         Width           =   3375
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sysop Login"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1080
      TabIndex        =   0
      Top             =   30
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1050
      Left            =   -60
      Picture         =   "frmLogin.frx":0442
      Stretch         =   -1  'True
      Top             =   -180
      Width           =   1080
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   -90
      X2              =   4890
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H001394F2&
      FillStyle       =   0  'Solid
      Height          =   915
      Left            =   -30
      Top             =   -30
      Width           =   4095
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'converts a Long  to a string
Public Function ConvertAddressToString(longAddr As Long) As String


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "ConvertAddressToString"
    Const ContainerName = "frmLogin"
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
bDebug = 1
    If bDebug = -1 Then
        On Error GoTo 0
    ElseIf bDebug = 1 Then
        On Error Resume Next
    Else
        
    End If

    Dim myByte(3) As Byte
    Dim Cnt As Long
    CopyMemory myByte(0), longAddr, 4
    For Cnt = 0 To 3
        ConvertAddressToString = ConvertAddressToString + CStr(myByte(Cnt)) + "."
    Next Cnt
    ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
Exit Function




End Function

Private Sub cmdClose_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdClose_Click"
    Const ContainerName = "frmLogin"
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
bDebug = 1
    If bDebug = -1 Then
        On Error GoTo 0
    ElseIf bDebug = 1 Then
        On Error Resume Next
    Else
       
    End If


    End
    
Exit Sub




End Sub

Private Sub cmdLogin_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdLogin_Click"
    Const ContainerName = "frmLogin"
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
bDebug = 1
    If bDebug = -1 Then
        On Error GoTo 0
    ElseIf bDebug = 1 Then
        On Error Resume Next
    Else
        'On Error GoTo ErrorOccur
    End If
   
    Dim ix As Integer
    Dim bFound As Boolean
    Dim bResult As Boolean
    Dim bPassed As Boolean
    Dim rsVISP As ADODB.Recordset
    Dim rsLoad As ADODB.Recordset
        
    If InStr(cmbUsername.Text, "%") > 0 Then Exit Sub
    
    Dim sql As String
    
    sql = "Select distinct sysops.* ,decode(sysops.Password,'" + PasswordSalt + "') as decPassword, AES_DECRYPT(sysops.PublicKey,'" & PublicKey & "') as PubKey From sysops where Username = '" & cmbUsername.Text & "'"
    
    bResult = oMySQL.OpenTable(oConn, rsLoad, , sql) 'MySQL.virtualisp(sql, "sysops", True, True))
        
    If rsLoad.State = adStateClosed Then bPassed = False
    
    If rsLoad.State = adStateOpen Then
        
        If rsLoad.RecordCount > 0 Then
        
            While Not rsLoad.EOF
                If rsLoad!DecPassword = txtPassword.Text Then
                    Login.PublicKey = IIf(IsNull(rsLoad!PubKey) = True, "", rsLoad!PubKey)
                    Login.lSysopID = rsLoad!RecID
                    Login.sUsername = cmbUsername.Text
                    Login.lVirtualID = rsLoad!VirtualID
                    Login.lLevel = rsLoad!SecurityLevel
                    Login.bCreateSysop = IIf(Val(rsLoad!bCreateSysop) = 0, False, True)
                    Login.bAgency = Val(rsLoad!bAgency)
                    Login.lAgencyID = rsLoad!AgencyID
                    Login.bPrimary = Val(rsLoad!bPrimary)
                    Login.bTemplates = Val(rsLoad!bTemplates)
                    Login.bMaster = IIf(IsNull(rsLoad!Master), 0, Val(rsLoad!Master))
                    Login.bVISP = IIf(IsNull(rsLoad!bVISP), 0, Val(rsLoad!bVISP))
                    Login.bVISPFiscal = IIf(IsNull(rsLoad!bVISPFiscal), 0, Val(rsLoad!bVISPFiscal))
                    Login.bRunMaintenance = IIf(IsNull(rsLoad!bMaintain), 0, Val(rsLoad!bMaintain))
                    Login.bRecievables = IIf(IsNull(rsLoad!bRecievables), 0, Val(rsLoad!bRecievables))
                    Login.bInvoice = IIf(IsNull(rsLoad!bInvoice), 0, Val(rsLoad!bInvoice))
                    Login.bExpenditure = IIf(IsNull(rsLoad!bExpenditure), 0, Val(rsLoad!bExpenditure))
                    Login.bHoldings = IIf(IsNull(rsLoad!bHoldings), 0, Val(rsLoad!bHoldings))
                    Login.bComm = IIf(IsNull(rsLoad!bComm), 0, Val(rsLoad!bComm))
                    Login.bRefund = IIf(IsNull(rsLoad!bRefund), 0, Val(rsLoad!bRefund))
                    Login.bAddCust = IIf(IsNull(rsLoad!bAddCust), 0, Val(rsLoad!bAddCust))
                    Login.bOwnership = IIf(IsNull(rsLoad!bOwnership), 0, Val(rsLoad!bOwnership))
                    Login.bAccSettings = IIf(IsNull(rsLoad!bAccSettings), 0, Val(rsLoad!bAccSettings))
                    Login.bVendors = IIf(IsNull(rsLoad!bVendors), 0, Val(rsLoad!bVendors))
                    
                    If Login.lVirtualID <> 0 Then
                        bResult = oMySQL.OpenTable(oConn, rsVISP, , "Select RecID, Realm, MISCFee, SysopID from virtualisp Where RecID = " & Login.lVirtualID)
                        If rsVISP.State = adStateOpen Then
                        
                            Login.sVISPDomain = rsVISP!Realm
                            Login.MISCFee = rsVISP!MISCFee
                            If Login.lSysopID = rsVISP!SysopID Then Login.bVISPPrimary = True
                        End If
                    Else
                        Login.sVISPDomain = "ep.net.au"
                        If Login.lSysopID = 3 Then Login.bVISPPrimary = True
                    End If
                    
                    App.LogEvent ":: Sysop Login :: " & cmbUsername.Text & " ::", vbLogEventTypeInformation
                    bPassed = True
                    rsLoad.MoveLast
                End If
                rsLoad.MoveNext
            Wend
        Else
            bPassed = False
        End If
        
        If bPassed = True Then
            
            On Error Resume Next
            
            For ix = 0 To cmbUsername.ListCount - 1
                If LCase(cmbUsername.Text) = LCase(cmbUsername.List(ix)) Then bFound = True
            Next
            
            If bFound = False Then cmbUsername.AddItem cmbUsername.Text
                
            If bFound = True Then
                SaveSetting "projectalpha", "Login", "Default", cmbUsername.ListIndex
            Else
                SaveSetting "projectalpha", "Login", "Default", cmbUsername.ListCount - 1
            End If
            
            SaveSetting "projectalpha", "Login", "NumberOf", cmbUsername.ListCount
            For ix = 0 To cmbUsername.ListCount - 1
                SaveSetting "projectalpha", "Login", "User" & ix, cmbUsername.List(ix)
            Next
            
            If chkRemember.Value = 1 Then
                Dim rsEnc As ADODB.Recordset
                Call oMySQL.OpenTable(oConn, rsEnc, , "select AES_ENCRYPT('" & txtPassword.Text & "','" & PublicKey & "') as nResult")
                SaveSetting "projectalpha", "Login", "Password", rsEnc!nResult
            Else
                SaveSetting "projectalpha", "Login", "Password", ""
            End If
            
           'This example was created by George Bernier (bernig@dinomail.qc.ca)
            Dim error As Long
            Dim FixedInfoSize As Long
            Dim AdapterInfoSize As Long
            Dim I As Integer
            Dim PhysicalAddress  As String
            Dim NewTime As Date
            Dim AdapterInfo As IP_ADAPTER_INFO
            Dim Adapt As IP_ADAPTER_INFO
            Dim AddrStr As IP_ADDR_STRING
            Dim FixedInfo As FIXED_INFO
            Dim Buffer As IP_ADDR_STRING
            Dim pAddrStr As Long
            Dim pAdapt As Long
            Dim Buffer2 As IP_ADAPTER_INFO
            Dim FixedInfoBuffer() As Byte
            Dim AdapterInfoBuffer() As Byte
            Dim Info As String
            'Get the main IP configuration information for this machine using a FIXED_INFO structure
            FixedInfoSize = 0
            error = GetNetworkParams(ByVal 0&, FixedInfoSize)
            If error <> 0 Then
                If error <> ERROR_BUFFER_OVERFLOW Then
                   Info = Info + Trim(sSTR.ReplaceString("GetNetworkParams sizing failed with error " & error, Chr$(0), " "))
                   Exit Sub
                End If
            End If
            ReDim FixedInfoBuffer(FixedInfoSize - 1)
        
            error = GetNetworkParams(FixedInfoBuffer(0), FixedInfoSize)
            If error = 0 Then
                    CopyMemory FixedInfo, FixedInfoBuffer(0), Len(FixedInfo)
                    Info = Info + Trim(sSTR.ReplaceString("Host Name:  " & FixedInfo.HostName, Chr$(0), " "))
                    Info = Info + vbCrLf + Trim(sSTR.ReplaceString("DNS Servers:  " & FixedInfo.DnsServerList.IpAddress, Chr$(0), " "))
                    pAddrStr = FixedInfo.DnsServerList.Next
                    Do While pAddrStr <> 0
                          CopyMemory Buffer, ByVal pAddrStr, Len(Buffer)
                          Info = Info + vbCrLf + Trim(sSTR.ReplaceString("DNS Servers:  " & Buffer.IpAddress, Chr$(0), " ")) 'dns server IP
                          pAddrStr = Buffer.Next
                    Loop
                    
                    Select Case FixedInfo.NodeType 'node type
                               Case 1
                                          Info = Info + vbCrLf + Trim(sSTR.ReplaceString("Node type: Broadcast", Chr$(0), " "))
                               Case 2
                                           Info = Info + vbCrLf + Trim(sSTR.ReplaceString("Node type: Peer to peer", Chr$(0), " "))
                               Case 4
                                            Info = Info + vbCrLf + Trim(sSTR.ReplaceString("Node type: Mixed", Chr$(0), " "))
                               Case 8
                                            Info = Info + vbCrLf + Trim(sSTR.ReplaceString("Node type: Hybrid", Chr$(0), " "))
                               Case Else
                                            Info = Info + vbCrLf + Trim(sSTR.ReplaceString("Unknown node type", Chr$(0), " "))
                    End Select
                    
                    Info = Info + vbCrLf + Trim(sSTR.ReplaceString("NetBIOS Scope ID:  " & FixedInfo.ScopeId, Chr$(0), " ")) 'scope ID
                    'routing
                    If FixedInfo.EnableRouting Then
                               Info = Info + vbCrLf + Trim(sSTR.ReplaceString("IP Routing Enabled ", Chr$(0), " "))
                    Else
                               Info = Info + vbCrLf + Trim(sSTR.ReplaceString("IP Routing not enabled", Chr$(0), " "))
                    End If
                    ' proxy
                    If FixedInfo.EnableProxy Then
                               Info = Info + vbCrLf + Trim(sSTR.ReplaceString("WINS Proxy Enabled ", Chr$(0), " "))
                    Else
                               Info = Info + vbCrLf + Trim(sSTR.ReplaceString("WINS Proxy not Enabled ", Chr$(0), " "))
                    End If
                    ' netbios
                    If FixedInfo.EnableDns Then
                              Info = Info + vbCrLf + Trim(sSTR.ReplaceString("NetBIOS Resolution Uses DNS ", Chr$(0), " "))
                    Else
                              Info = Info + vbCrLf + Trim(sSTR.ReplaceString("NetBIOS Resolution Does not use DNS  ", Chr$(0), " "))
                    End If
            Else
                    Info = Info + vbCrLf + Trim(sSTR.ReplaceString("GetNetworkParams failed with error " & error, Chr$(0), " "))
                    Exit Sub
            End If
            
            'Enumerate all of the adapter specific information using the IP_ADAPTER_INFO structure.
            'Note:  IP_ADAPTER_INFO contains a linked list of adapter entrues.
            
            AdapterInfoSize = 0
            error = GetAdaptersInfo(ByVal 0&, AdapterInfoSize)
            If error <> 0 Then
                If error <> ERROR_BUFFER_OVERFLOW Then
                   Info = Info + vbCrLf + Trim(sSTR.ReplaceString("GetAdaptersInfo sizing failed with error " & error, Chr$(0), " "))
                   Exit Sub
                End If
            End If
           ReDim AdapterInfoBuffer(AdapterInfoSize - 1)
         
         ' Get actual adapter information
           error = GetAdaptersInfo(AdapterInfoBuffer(0), AdapterInfoSize)
           If error <> 0 Then
              Info = Info + vbCrLf + Trim(sSTR.ReplaceString("GetAdaptersInfo failed with error " & error, Chr$(0), " "))
              Exit Sub
           End If
           CopyMemory AdapterInfo, AdapterInfoBuffer(0), Len(AdapterInfo)
           pAdapt = AdapterInfo.Next
        
           Do While pAdapt <> 0
                CopyMemory Buffer2, AdapterInfo, Len(Buffer2)
                  Select Case Buffer2.Type
                        Case MIB_IF_TYPE_ETHERNET
                            Info = Info + vbCrLf + Trim(sSTR.ReplaceString("Ethernet adapter ", Chr$(0), " "))
                        Case MIB_IF_TYPE_TOKENRING
                            Info = Info + vbCrLf + Trim(sSTR.ReplaceString("Token Ring adapter ", Chr$(0), " "))
                        Case MIB_IF_TYPE_FDDI
                            Info = Info + vbCrLf + Trim(sSTR.ReplaceString("FDDI adapter ", Chr$(0), " "))
                        Case MIB_IF_TYPE_PPP
                            Info = Info + vbCrLf + Trim(sSTR.ReplaceString("PPP adapter", Chr$(0), " "))
                        Case MIB_IF_TYPE_LOOPBACK
                            Info = Info + vbCrLf + Trim(sSTR.ReplaceString("Loopback adapter ", Chr$(0), " "))
                        Case MIB_IF_TYPE_SLIP
                            Info = Info + vbCrLf + Trim(sSTR.ReplaceString("Slip adapter ", Chr$(0), " "))
                        Case Else
                            Info = Info + vbCrLf + Trim(sSTR.ReplaceString("Other adapter ", Chr$(0), " "))
                End Select
            Info = Info + vbCrLf + Trim(sSTR.ReplaceString(" AdapterName: " & Buffer2.AdapterName, Chr$(0), " "))
            Info = Info + vbCrLf + Trim(sSTR.ReplaceString("AdapterDescription: " & Buffer2.Description, Chr$(0), " ")) 'adatpter name
        
            For I = 0 To Buffer2.AddressLength - 1
                   PhysicalAddress = PhysicalAddress & Hex(Buffer2.Address(I))
                    If I < Buffer2.AddressLength - 1 Then
                     PhysicalAddress = PhysicalAddress & "-"
                    End If
        
            Next
            Info = Info + vbCrLf + Trim(sSTR.ReplaceString("Physical Address: " & PhysicalAddress, Chr$(0), " ")) 'mac address
            If Buffer2.DhcpEnabled Then
                    Info = Info + vbCrLf + Trim(sSTR.ReplaceString("DHCP Enabled ", Chr$(0), " "))
            Else
                    Info = Info + vbCrLf + Trim(sSTR.ReplaceString("DHCP disabled", Chr$(0), " "))
            End If
        
            pAddrStr = Buffer2.IpAddressList.Next
            Do While pAddrStr <> 0
                   CopyMemory Buffer, Buffer2.IpAddressList, LenB(Buffer)
                   Info = Info + vbCrLf + Trim(sSTR.ReplaceString("IP Address: " & Buffer.IpAddress, Chr$(0), " "))
                   Info = Info + vbCrLf + Trim(sSTR.ReplaceString("Subnet Mask: " & Buffer.IpMask, Chr$(0), " "))
                   pAddrStr = Buffer.Next
                   If pAddrStr <> 0 Then
                    CopyMemory Buffer2.IpAddressList, ByVal pAddrStr, Len(Buffer2.IpAddressList)
                   End If
           Loop
            Info = Info + vbCrLf + Trim(sSTR.ReplaceString("Default Gateway: " & Buffer2.GatewayList.IpAddress, Chr$(0), " "))
            pAddrStr = Buffer2.GatewayList.Next
            Do While pAddrStr <> 0
                    CopyMemory Buffer, Buffer2.GatewayList, Len(Buffer)
                    Info = Info + vbCrLf + Trim(sSTR.ReplaceString("IP Address: " & Buffer.IpAddress, Chr$(0), " "))
                    pAddrStr = Buffer.Next
                    If pAddrStr <> 0 Then
                    CopyMemory Buffer2.GatewayList, ByVal pAddrStr, Len(Buffer2.GatewayList)
                    End If
            Loop
        
            Info = Info + vbCrLf + Trim(sSTR.ReplaceString("DHCP Server: " & Buffer2.DhcpServer.IpAddress, Chr$(0), " "))
            Info = Info + vbCrLf + Trim(sSTR.ReplaceString("Primary WINS Server: " & Buffer2.PrimaryWinsServer.IpAddress, Chr$(0), " "))
            Info = Info + vbCrLf + Trim(sSTR.ReplaceString("Secondary WINS Server: " & Buffer2.SecondaryWinsServer.IpAddress, Chr$(0), " "))
        
            ' Display time
            NewTime = CDate(Adapt.LeaseObtained)
            Info = Info + vbCrLf + Trim(sSTR.ReplaceString("Lease Obtained: " & CStr(NewTime), Chr$(0), " "))
        
            NewTime = CDate(Adapt.LeaseExpires)
            Info = Info + vbCrLf + Trim(sSTR.ReplaceString("Lease Expires :  " & CStr(NewTime), Chr$(0), " "))
            pAdapt = Buffer2.Next
            If pAdapt <> 0 Then
                CopyMemory AdapterInfo, ByVal pAdapt, Len(AdapterInfo)
            End If
        
           Loop
            
            Dim Ret As Long, Tel As Long
            Dim bBytes() As Byte
            Dim Listing As MIB_IPADDRTABLE
            
            GetIpAddrTable ByVal 0&, Ret, True
            
            If Ret <= 0 Then Exit Sub
            ReDim bBytes(0 To Ret - 1) As Byte
            'retrueve the data
            GetIpAddrTable bBytes(0), Ret, False
              
            'Get the first 4 bytes to get the entry's.. ip installed
            CopyMemory Listing.dEntrys, bBytes(0), 4
            'MsgBox "IP's found : " & Listing.dEntrys    => Founded ip installed on your PC..
            For Tel = 0 To Listing.dEntrys - 1
                'Copy whole structure to Listing..
               ' MsgBox bBytes(tel) & "."
                CopyMemory Listing.mIPInfo(Tel), bBytes(4 + (Tel * Len(Listing.mIPInfo(0)))), Len(Listing.mIPInfo(Tel))
                Info = Info + vbCrLf + ConvertAddressToString(Listing.mIPInfo(Tel).dwAddr)
            Next
            
            
            Info = sSTR.ReplaceString(Info, Chr(0), "")
            'MsgBox oMySQL.ESC(Info)
            'mysql.execute oConn,  "update sysops set IP='" & Winsock1.LocalIP & ", LastNetworkAdd = '" & oMySQL.ESC(Info) & "', project alphaVersion = " '" & App.Major & "." & App.Minor & "." & App.Revision & "' where RecID = " & rsload!RecID
            
            Dim rsSave As ADODB.Recordset
            
            
                oMySQL.Execute oConn, "update sysops set IP = '" & Winsock1.LocalIP & "' where RecID = " & rsLoad!RecID, True
                oMySQL.Execute oConn, "update sysops set LastNetworkAdd = '" & Info & "' where RecID = " & rsLoad!RecID, True
                oMySQL.Execute oConn, "update sysops set prjAlphaVersion = '" & App.Major & "." & App.Minor & "." & App.Revision & "' where RecID = " & rsLoad!RecID, True
            
            
            frmAgent.oChar.StopAll
            Unload Me
        Else
            frmAgent.oChar.Play "Decline"
            frmAgent.oChar.Speak "Password or Username was incorrect!"
        End If
    Else
        If rsLoad.RecordCount = 0 Then
            frmAgent.oChar.Play "Decline"
            frmAgent.oChar.Speak "Username not found!"
        Else
            frmAgent.oChar.Play "Decline"
            frmAgent.oChar.Speak "Password was Incorrect!"
        End If
    End If
            
    
   
        
Exit Sub




End Sub

Private Sub Form_Load()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Load"
    Const ContainerName = "frmLogin"
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
bDebug = 1
    If bDebug = -1 Then
        On Error GoTo 0
    ElseIf bDebug = 1 Then
        On Error Resume Next
    Else
        'On Error GoTo ErrorOccur
    End If


            
    Dim iUsernames As Integer
    
    iUsernames = GetSetting("projectalpha", "Login", "NumberOf", 0)
    
    If iUsernames > 0 Then
        For ix = 0 To iUsernames - 1
            cmbUsername.AddItem GetSetting("projectalpha", "Login", "User" & ix, "")
        Next
    End If
    
    cmbUsername.ListIndex = GetSetting("projectalpha", "Login", "Default", -1)
    
    Dim pass As String
    
    Dim rsEnc As ADODB.Recordset
    
    Call oMySQL.OpenTable(oConn, rsEnc, , "select AES_DECRYPT('" & GetSetting("projectalpha", "Login", "Password", "") & "','" & PublicKey & "') as nResult")
    
    If IsNull(rsEnc!nResult) Then txtPassword.Text = "" Else txtPassword.Text = rsEnc!nResult
    
    If txtPassword.Text <> "" Then chkRemember.Value = 1
            
Exit Sub



End Sub

Private Sub Form_Paint()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Paint"
    Const ContainerName = "frmLogin"
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
    bDebug = 1
    
    If bDebug = -1 Then
        On Error GoTo 0
    ElseIf bDebug = 1 Then
        On Error Resume Next
    Else
        
    End If


    Static bAgent As Boolean
    
    If bAgent = False Then
        'frmAgent.LoadChar "merlin.asc", 200, 200
        If frmAgent.oChar Is Nothing Then
            frmAgent.LoadChar DirName + "msagent\chars\merlin.acs", (Me.Left + Me.Width + 120) / Screen.TwipsPerPixelX, (Me.Top + Me.Height / 2) / Screen.TwipsPerPixelY
        Else
            frmAgent.oChar.MoveTo (Me.Left + Me.Width + 120) / Screen.TwipsPerPixelX, (Me.Top + Me.Height / 2) / Screen.TwipsPerPixelY
        End If
        frmAgent.oChar.Show
        frmAgent.oChar.GestureAt (Me.Left + Frame1.Left) / Screen.TwipsPerPixelX, (Me.Top + Frame1.Top) / Screen.TwipsPerPixelY
        
        frmAgent.oChar.Speak "Welcome to project alpha, Please enter in your username and password to enter the system, this would have been provided to you by the carrier."
        bAgent = True
    End If
    
Exit Sub





End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_QueryUnload"
    Const ContainerName = "frmLogin"
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
bDebug = 1
    If bDebug = -1 Then
        On Error GoTo 0
    ElseIf bDebug = 1 Then
        On Error Resume Next
    Else
        'On Error GoTo ErrorOccur
    End If


    If UnloadMode <> vbFormCode Then Cancel = True
    
Exit Sub


End Sub
