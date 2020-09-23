VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H0084E8E8&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   4035
   ClientLeft      =   4365
   ClientTop       =   1875
   ClientWidth     =   5205
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
   ScaleHeight     =   4035
   ScaleWidth      =   5205
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1740
      Top             =   1500
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H0084E8E8&
      Caption         =   "&Close and Exit"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3060
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3360
      Width           =   1995
   End
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H0084E8E8&
      Caption         =   "&Login"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3360
      Width           =   2115
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0084E8E8&
      Caption         =   "&Password"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BA3F3F&
      Height          =   1185
      Left            =   150
      TabIndex        =   3
      Top             =   1980
      Width           =   4905
      Begin VB.CheckBox chkRemember 
         BackColor       =   &H0084E8E8&
         Caption         =   "&Remember Password"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BA3F3F&
         Height          =   285
         Left            =   150
         TabIndex        =   5
         Top             =   810
         Width           =   3315
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H0084E8E8&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BA3F3F&
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   150
         PasswordChar    =   "›"
         TabIndex        =   4
         Top             =   360
         Width           =   4635
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0084E8E8&
      Caption         =   "&Username"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BA3F3F&
      Height          =   915
      Left            =   150
      TabIndex        =   1
      Top             =   990
      Width           =   4905
      Begin VB.ComboBox cmbUsername 
         BackColor       =   &H0084E8E8&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BA3F3F&
         Height          =   525
         Left            =   150
         TabIndex        =   2
         Top             =   300
         Width           =   4635
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sysop Login"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   1080
      TabIndex        =   0
      Top             =   30
      Width           =   3030
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
      X1              =   -330
      X2              =   6000
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H001394F2&
      FillStyle       =   0  'Solid
      Height          =   915
      Left            =   0
      Top             =   -30
      Width           =   5835
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
    If bDebug = -1 Then
        On Error GoTo 0
    ElseIf bDebug = 1 Then
        On Error Resume Next
    Else
        On Error GoTo ErrorOccur
    End If

    Dim myByte(3) As Byte
    Dim Cnt As Long
    CopyMemory myByte(0), longAddr, 4
    For Cnt = 0 To 3
        ConvertAddressToString = ConvertAddressToString + CStr(myByte(Cnt)) + "."
    Next Cnt
    ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
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
    If bDebug = -1 Then
        On Error GoTo 0
    ElseIf bDebug = 1 Then
        On Error Resume Next
    Else
        On Error GoTo ErrorOccur
    End If


    End
    
Exit Sub



ErrorOccur:
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

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
    If bDebug = -1 Then
        On Error GoTo 0
    ElseIf bDebug = 1 Then
        On Error Resume Next
    Else
        On Error GoTo ErrorOccur
    End If
   
    Dim ix As Integer
    Dim bFound As Boolean
    Dim bResult As Boolean
    Dim bPassed As Boolean
    Dim rsVISP As ADODB.Recordset
    Dim rsload As New ADODB.Recordset
        
    If InStr(cmbUsername.Text, "%") > 0 Then Exit Sub
    
    Dim SQL As String
    
    SQL = "select distinct sysops.* ,decode(sysops.Password,'" + odb.colSalts.ReturnSalt(PWSalt) + "') as decPassword, AES_DECRYPT('PublicKey','" & odb.colSalts.ReturnSalt("PublicKey") & "') as PubKey from sysops where Username = '" & cmbUsername.Text & "'"
    
    If ADOConn.State = adStateOpen Then rsload.Open SQL, ADOConn, adOpenDynamic, adLockReadOnly
        
    If rsload.State = adStateClosed Then bPassed = False
    
    If rsload.State = adStateOpen Then
        
            While Not rsload.EOF And Err.Number = 0
                If rsload!DecPassword = txtPassword.Text Then
                    Login.PublicKey = IIf(IsNull(rsload!PubKey) = True, "", rsload!PubKey)
                    Login.lSysopID = rsload!RecID
                    Login.sUsername = cmbUsername.Text
                    Login.lVirtualID = rsload!VirtualID
                    Login.lLevel = rsload!SecurityLevel
                    Login.bCreateSysop = IIf(Val(rsload!bCreateSysop) = 0, False, True)
                    Login.bAgency = Val(rsload!bAgency)
                    Login.lAgencyID = rsload!AgencyID
                    Login.bPrimary = Val(rsload!bPrimary)
                    Login.bTemplates = Val(rsload!bTemplates)
                    Login.bMaster = IIf(IsNull(rsload!Master), 0, Val(rsload!Master))
                    Login.bVISP = IIf(IsNull(rsload!bVISP), 0, Val(rsload!bVISP))
                    Login.bVISPFiscal = IIf(IsNull(rsload!bVISPFiscal), 0, Val(rsload!bVISPFiscal))
                    Login.bRunMaintenance = IIf(IsNull(rsload!bMaintain), 0, Val(rsload!bMaintain))
                    Login.bRecievables = IIf(IsNull(rsload!bRecievables), 0, Val(rsload!bRecievables))
                    Login.bInvoice = IIf(IsNull(rsload!bInvoice), 0, Val(rsload!bInvoice))
                    Login.bExpenditure = IIf(IsNull(rsload!bExpenditure), 0, Val(rsload!bExpenditure))
                    Login.bHoldings = IIf(IsNull(rsload!bHoldings), 0, Val(rsload!bHoldings))
                    Login.bComm = IIf(IsNull(rsload!bComm), 0, Val(rsload!bComm))
                    Login.bRefund = IIf(IsNull(rsload!bRefund), 0, Val(rsload!bRefund))
                    Login.bAddCust = IIf(IsNull(rsload!bAddCust), 0, Val(rsload!bAddCust))
                    Login.bOwnership = IIf(IsNull(rsload!bOwnership), 0, Val(rsload!bOwnership))
                    Login.bAccSettings = IIf(IsNull(rsload!bAccSettings), 0, Val(rsload!bAccSettings))
                    Login.bVendors = IIf(IsNull(rsload!bVendors), 0, Val(rsload!bVendors))
                                        
                    Call oViSP.PopulateMe(Login.lVirtualID, ADOConn)
                    Call oResell.PopulateResellers(ADOConn, fs_LoadHeader, Login.lVirtualID, "VISPDetails")
                    
                    If Login.lVirtualID <> 1000 Then
                        bResult = MySQL.OpenTable(ADOConn, rsVISP, , "select RecID, Realm, ABN, Description, bTaxMode, cTaxCode, cTaxCountry, MISCFee, SysopID from virtualisp Where RecID = " & Login.lVirtualID)
                        If rsVISP.State = adStateOpen Then
                            Login.TaxMode = Val(rsVISP!bTaxMode)
                            Login.TaxCountry = rsVISP!cTaxCountry
                            Login.TaxCode = rsVISP!cTaxCode
                            Login.ViSPDesc = IIf(IsNull(rsVISP!Description), "", rsVISP!Description)
                            Login.ViSPDesc = IIf(IsNull(rsVISP!ABN), "", rsVISP!ABN) & " - " & Login.ViSPDesc
                            Login.sVISPDomain = rsVISP!Realm
                            Login.MISCFee = rsVISP!MISCFee
                            If Login.lSysopID = rsVISP!SysopID Then Login.bVISPPrimary = True
                        End If
                    Else
                            Login.ViSPDesc = "Exitstencil Press Pty. Ltd."
                            Login.TaxMode = 0
                            Login.TaxCountry = "AUS0001"
                            Login.TaxCode = "GST"
                        
                        Login.sVISPDomain = "ep.net.au"
                        If Login.lSysopID = 3 Then Login.bVISPPrimary = True
                    End If
                    
                    App.LogEvent ":: Sysop Login :: " & cmbUsername.Text & " ::", vbLogEventTypeInformation
                    bPassed = True
                    rsload.MoveLast
                End If
                rsload.MoveNext
            Wend
        
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
                Call MySQL.OpenTable(ADOConn, rsEnc, , "select AES_ENCRYPT('" & txtPassword.Text & "','" & odb.colSalts.ReturnSalt("PublicKey") & "') as nResult")
                SaveSetting "projectalpha", "Login", "Password", rsEnc!nResult
            Else
                SaveSetting "projectalpha", "Login", "Password", ""
            End If
            
           'This example was created by George Bernier (bernig@dinomail.qc.ca)
            Dim error As Long
            Dim FixedInfoSize As Long
            Dim AdapterInfoSize As Long
            Dim i As Integer
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
                    Info = Info + vbCrLf + Trim(sSTR.ReplaceString("DNS Servers:  " & FixedInfo.DnsServerList.IPAddress, Chr$(0), " "))
                    pAddrStr = FixedInfo.DnsServerList.Next
                    Do While pAddrStr <> 0 And Err.Number = 0
                          CopyMemory Buffer, ByVal pAddrStr, Len(Buffer)
                          Info = Info + vbCrLf + Trim(sSTR.ReplaceString("DNS Servers:  " & Buffer.IPAddress, Chr$(0), " ")) 'dns server IP
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
        
            For i = 0 To Buffer2.AddressLength - 1
                   PhysicalAddress = PhysicalAddress & Hex(Buffer2.Address(i))
                    If i < Buffer2.AddressLength - 1 Then
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
                   Info = Info + vbCrLf + Trim(sSTR.ReplaceString("IP Address: " & Buffer.IPAddress, Chr$(0), " "))
                   Info = Info + vbCrLf + Trim(sSTR.ReplaceString("Subnet Mask: " & Buffer.IpMask, Chr$(0), " "))
                   pAddrStr = Buffer.Next
                   If pAddrStr <> 0 Then
                    CopyMemory Buffer2.IpAddressList, ByVal pAddrStr, Len(Buffer2.IpAddressList)
                   End If
           Loop
            Info = Info + vbCrLf + Trim(sSTR.ReplaceString("Default Gateway: " & Buffer2.GatewayList.IPAddress, Chr$(0), " "))
            pAddrStr = Buffer2.GatewayList.Next
            Do While pAddrStr <> 0
                    CopyMemory Buffer, Buffer2.GatewayList, Len(Buffer)
                    Info = Info + vbCrLf + Trim(sSTR.ReplaceString("IP Address: " & Buffer.IPAddress, Chr$(0), " "))
                    pAddrStr = Buffer.Next
                    If pAddrStr <> 0 Then
                    CopyMemory Buffer2.GatewayList, ByVal pAddrStr, Len(Buffer2.GatewayList)
                    End If
            Loop
        
            Info = Info + vbCrLf + Trim(sSTR.ReplaceString("DHCP Server: " & Buffer2.DhcpServer.IPAddress, Chr$(0), " "))
            Info = Info + vbCrLf + Trim(sSTR.ReplaceString("Primary WINS Server: " & Buffer2.PrimaryWinsServer.IPAddress, Chr$(0), " "))
            Info = Info + vbCrLf + Trim(sSTR.ReplaceString("Secondary WINS Server: " & Buffer2.SecondaryWinsServer.IPAddress, Chr$(0), " "))
        
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
            'MsgBox MySQL.ESC(Info)
            'mysql.execute adoconn,  "update sysops set IP='" & Winsock1.LocalIP & ", LastNetworkAdd = '" & MySQL.ESC(Info) & "', project alphaVersion = " '" & App.Major & "." & App.Minor & "." & App.Revision & "' where RecID = " & rsload!RecID
            
            Dim rsSave As ADODB.Recordset
            
            
                MySQL.Execute ADOConn, "update sysops set IP = '" & Winsock1.LocalIP & "' where RecID = " & rsload!RecID, True
                MySQL.Execute ADOConn, "update sysops set LastNetworkAdd = '" & Info & "' where RecID = " & rsload!RecID, True
                MySQL.Execute ADOConn, "update sysops set prjAlphaVersion = '" & App.Major & "." & App.Minor & "." & App.Revision & "' where RecID = " & rsload!RecID, True
            
            
            frmAgent.oChar.StopAll
            Unload Me
        Else
            frmAgent.oChar.Play "Decline"
            frmAgent.oChar.Speak "Password or Username was incorrect!"
        End If
    Else
        If rsload.RecordCount = 0 Then
            frmAgent.oChar.Play "Decline"
            frmAgent.oChar.Speak "Username not found!"
        Else
            frmAgent.oChar.Play "Decline"
            frmAgent.oChar.Speak "Password was Incorrect!"
        End If
    End If
            
    
   
        
Exit Sub



ErrorOccur:
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

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
    If bDebug = -1 Then
        On Error GoTo 0
    ElseIf bDebug = 1 Then
        On Error Resume Next
    Else
        On Error GoTo ErrorOccur
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
    
    Call MySQL.OpenTable(ADOConn, rsEnc, , "select AES_DECRYPT('" & GetSetting("projectalpha", "Login", "Password", "") & "','" & odb.colSalts.ReturnSalt("PublicKey") & "') as nResult")
    
    If rsEnc.State = adStateOpen Then
        If Not rsEnc.EOF And Not rsEnc.BOF Then
            If IsNull(rsEnc!nResult) Then txtPassword.Text = "" Else txtPassword.Text = rsEnc!nResult
        End If
    End If
    If txtPassword.Text <> "" Then chkRemember.Value = 1
            
Exit Sub



ErrorOccur:
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

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
    If bDebug = -1 Then
        On Error GoTo 0
    ElseIf bDebug = 1 Then
        On Error Resume Next
    Else
        On Error GoTo ErrorOccur
    End If


    Static bAgent As Boolean
    
    If bAgent = False Then
        'frmAgent.LoadChar "merlin.asc", 200, 200
        If frmAgent.oChar Is Nothing Then
            frmAgent.LoadChar DirName + "msagent\chars\robby.acs", (Me.Left + Me.Width + 120) / Screen.TwipsPerPixelX, (Me.Top + Me.Height / 2) / Screen.TwipsPerPixelY
        Else
            frmAgent.oChar.MoveTo (Me.Left + Me.Width + 120) / Screen.TwipsPerPixelX, (Me.Top + Me.Height / 2) / Screen.TwipsPerPixelY
        End If
        frmAgent.oChar.Show
        frmAgent.oChar.GestureAt (Me.Left + Frame1.Left) / Screen.TwipsPerPixelX, (Me.Top + Frame1.Top) / Screen.TwipsPerPixelY
        
        frmAgent.oChar.Speak "Welcome to project alpha, Please enter in your username and password to enter the system, this would have been provided to you by the carrier."
        bAgent = True
    End If
    
Exit Sub



ErrorOccur:
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

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
    If bDebug = -1 Then
        On Error GoTo 0
    ElseIf bDebug = 1 Then
        On Error Resume Next
    Else
        On Error GoTo ErrorOccur
    End If


    If UnloadMode <> vbFormCode Then Cancel = True
    
Exit Sub



ErrorOccur:
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Sub
