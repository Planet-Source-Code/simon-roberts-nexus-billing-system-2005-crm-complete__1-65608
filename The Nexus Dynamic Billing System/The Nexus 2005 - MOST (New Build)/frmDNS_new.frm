VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmDNS 
   Appearance      =   0  'Flat
   BackColor       =   &H00BA3F3F&
   Caption         =   "Form1"
   ClientHeight    =   9960
   ClientLeft      =   1680
   ClientTop       =   3270
   ClientWidth     =   9360
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0084E8E8&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9960
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H0084E8E8&
      Caption         =   "&Save and Close"
      Height          =   435
      Index           =   1
      Left            =   7170
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   9420
      Width           =   2025
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0084E8E8&
      Caption         =   "&Add Docs"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9420
      Width           =   1935
   End
   Begin MSComctlLib.ListView lsDocs 
      Height          =   4545
      Left            =   270
      TabIndex        =   14
      Top             =   4800
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   8017
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   0
      BackColor       =   14718607
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Doc Type"
         Object.Width           =   5080
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Document Description"
         Object.Width           =   10583
      EndProperty
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA3F3F&
      Height          =   735
      Left            =   8730
      Picture         =   "frmDNS_new.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3990
      Width           =   435
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00BA3F3F&
      Caption         =   "Technicians Password"
      ForeColor       =   &H0084E8E8&
      Height          =   765
      Index           =   5
      Left            =   4500
      TabIndex        =   11
      Top             =   3930
      Width           =   4185
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         Height          =   345
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   150
         PasswordChar    =   "ž"
         TabIndex        =   12
         Top             =   270
         Width           =   3885
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00BA3F3F&
      Caption         =   "Technicians Name"
      ForeColor       =   &H0084E8E8&
      Height          =   765
      Index           =   4
      Left            =   240
      TabIndex        =   9
      Top             =   3930
      Width           =   4185
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         Height          =   345
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   150
         TabIndex        =   10
         Top             =   270
         Width           =   3885
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00BA3F3F&
      Caption         =   "Key for domain"
      ForeColor       =   &H0084E8E8&
      Height          =   765
      Index           =   3
      Left            =   240
      TabIndex        =   7
      Top             =   3150
      Width           =   8925
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         Height          =   345
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   150
         PasswordChar    =   "ž"
         TabIndex        =   8
         Top             =   270
         Width           =   8655
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00BA3F3F&
      Caption         =   "Contact Name"
      ForeColor       =   &H0084E8E8&
      Height          =   765
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   2370
      Width           =   8925
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         Height          =   345
         Index           =   2
         Left            =   150
         TabIndex        =   6
         Top             =   270
         Width           =   8655
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00BA3F3F&
      Caption         =   "Admin's Email Address"
      ForeColor       =   &H0084E8E8&
      Height          =   765
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1590
      Width           =   8925
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         Height          =   345
         Index           =   1
         Left            =   150
         TabIndex        =   4
         Top             =   270
         Width           =   8655
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00BA3F3F&
      Caption         =   "Domain Name (without prefix)"
      ForeColor       =   &H0084E8E8&
      Height          =   765
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   810
      Width           =   8925
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   150
         TabIndex        =   2
         Top             =   270
         Width           =   8655
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8610
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNS_new.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNS_new.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNS_new.frx":0CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNS_new.frx":1138
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNS_new.frx":158A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNS_new.frx":19DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNS_new.frx":1E2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNS_new.frx":2280
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNS_new.frx":26D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNS_new.frx":2B24
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNS_new.frx":2F76
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNS_new.frx":33C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNS_new.frx":381A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNS_new.frx":3B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNS_new.frx":3F86
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNS_new.frx":43D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNS_new.frx":482A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNS_new.frx":4C7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNS_new.frx":50CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNS_new.frx":5520
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDNS_new.frx":5972
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   810
      X2              =   9510
      Y1              =   420
      Y2              =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Domain Name Service"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0084E8E8&
      Height          =   360
      Left            =   870
      TabIndex        =   0
      Top             =   60
      Width           =   3390
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   90
      Picture         =   "frmDNS_new.frx":5DC4
      Stretch         =   -1  'True
      Top             =   60
      Width           =   720
   End
End
Attribute VB_Name = "frmDNS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lv As ListView

Public dKey As String
Public oCNT As clsSubscriber
Public SESSION As String
Public iCloseState As frm_CloseStates
Private Sub cmdSave_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdSave_Click"
    Const ContainerName = "frmDNS"
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


    iCloseState = frmCloseSave
    Unload Me
    
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


Private Sub cmdClose_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdClose_Click"
    Const ContainerName = "frmDNS"
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


    iCloseState = frmCloseCancel
    Unload Me

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

Private Sub Command1_Click()

    If txtField(5).Text = txtField(5).Tag Then
        txtField(3).PasswordChar = ""
        txtField(5).PasswordChar = ""
    End If
    
End Sub

Private Sub Command2_Click(Index As Integer)

    Select Case Index
    Case 0
        Dim fNew As New frmDNSDocType
        
        fNew.Show 1
        
        If Trim(fNew.sDesc) <> "" Then
            Dim itmX As ListItem
            
            Set itmX = lsDocs.ListItems.Add(, UCase(Left(fNew.sType, 3)) & oIP.colDNSDocs.Count + 1, fNew.sType, , fNew.bIcon + 1)
            itmX.SubItems(1) = fNew.sDesc
            oIP.colDNSDocs.Add UCase(Left(fNew.sType, 3)) & oIP.colDNSDocs.Count + 1, 0, 0, fNew.sType, "", fNew.bIcon + 1, itmX.SubItems(1), itmX.Text
        End If
    Case 1
    
        If dKey = "" Then
            With oCNT.col_Domains.Add("NEW" & SESSION & oCNT.col_Domains.Count + 1, 0, txtField(0), txtField(1), "N", True, 10, txtField(3), oCNT.fRecID, txtField(4), _
                txtField(5).Text, Login.lSysopID, Login.lVirtualID, "NEW" & SESSION & oCNT.col_Domains.Count + 1)
                .ContactName = txtField(2)
            End With
        Else
            oCNT.col_Domains(dKey).ContactName = txtField(2)
            oCNT.col_Domains(dKey).Domain = txtField(0)
            oCNT.col_Domains(dKey).AdminEmail = txtField(1)
            oCNT.col_Domains(dKey).vKey = txtField(3)
            oCNT.col_Domains(dKey).TechName = txtField(4)
            oCNT.col_Domains(dKey).TechPass = IIf(txtField(5).Tag = "" Or txtField(5).Tag <> oCNT.col_Domains(dKey).TechPass, pass(txtField(5).Tag, "Verify Technicians Password"), txtField(5).Tag)
            Me.iCloseState = 2
            Unload Me
        End If
    End Select
    
End Sub

Private Sub Form_Load()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Load"
    Const ContainerName = "frmDNS"
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

    If SESSION = "" Then SESSION = GetSessionChar(SESSION, Me.hwnd, 14)
    
    If Login.lLevel > 90 Then
        txtField(3).PasswordChar = ""
    End If
    
    If oIP.colDNSDocs.Count > 0 Then
    
        lsDocs.ListItems.Clear
        Dim Xnt As Integer
        Dim itmX As ListItem
        For Xnt = 1 To oIP.colDNSDocs.Count
            
        
            Set itmX = lsDocs.ListItems.Add(, UCase(Left(oIP.colDNSDocs(Xnt).DocType, 3)) & Xnt, oIP.colDNSDocs(Xnt).ItemText, , oIP.colDNSDocs(Xnt).bIcon)
            
            itmX.SubItems(1) = oIP.colDNSDocs(Xnt).Description
            itmX.Tag = oIP.colDNSDocs(Xnt).DocText
        
        Next
    
    End If
    
    If lRecID = 0 And oIP.colDNSDocs.Count = 0 Then
    
        lsDocs.ListItems.Add , "ADD" & oIP.colDNSDocs.Count + 1, "Primary Address", , 21
        lsDocs.ListItems(1).SubItems(1) = "This is the main mailing address and referal address for this DNS"
        oIP.colDNSDocs.Add "ADD" & oIP.colDNSDocs.Count + 1, 0, 0, "ADDRESS", "", 21, lsDocs.ListItems(1).SubItems(1), lsDocs.ListItems(1).Text
        lsDocs.ListItems.Add , "ADD" & oIP.colDNSDocs.Count + 1, "Billing Address", , 21
        lsDocs.ListItems(2).SubItems(1) = "This is the billing address that renewal notices are sent to."
        oIP.colDNSDocs.Add "ADD" & oIP.colDNSDocs.Count + 1, 0, 0, "ADDRESS", "", 21, lsDocs.ListItems(2).SubItems(1), lsDocs.ListItems(2).Text
        lsDocs.ListItems.Add , "ADD" & oIP.colDNSDocs.Count + 1, "Technicians Address", , 21
        lsDocs.ListItems(3).SubItems(1) = "This is the address for the technician."
        oIP.colDNSDocs.Add "ADD" & oIP.colDNSDocs.Count + 1, 0, 0, "ADDRESS", "", 21, lsDocs.ListItems(3).SubItems(1), lsDocs.ListItems(3).Text
        
    Else
    
        txtField(2) = oCNT.col_Domains(dKey).ContactName
        txtField(0) = oCNT.col_Domains(dKey).Domain
        txtField(1) = oCNT.col_Domains(dKey).AdminEmail
        txtField(3) = oCNT.col_Domains(dKey).vKey
        txtField(4) = oCNT.col_Domains(dKey).TechName
        txtField(5).Tag = oCNT.col_Domains(dKey).TechPass
    End If
    
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_QueryUnload"
    Const ContainerName = "frmDNS"
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

    If UnloadMode = 0 Then
        oIP.colDNSDocs.Clear
    End If
    
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

Private Sub lsDocs_AfterLabelEdit(Cancel As Integer, NewString As String)


    oIP.colDNSDocs(Val(Mid(lsDocs.SelectedItem.Key, 4))).ItemText = NewString

End Sub

Private Sub lsDocs_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Call GUI.ColumnSort(ColumnHeader, lsDocs)
    
End Sub

Private Sub lsDocs_DblClick()

    Dim XML As String
    
    If lsDocs.SelectedItem Is Nothing Then
    
    Else
    
        XML = lsDocs.SelectedItem.Tag
        'MsgBox XML
        
        Select Case Left(lsDocs.SelectedItem.Key, 3)
        Case "ADD"
            
            Dim ffrmSnailMail As frmSnailMail
            Set ffrmSnailMail = New frmSnailMail

        
            ffrmSnailMail.sContactName = XMLTag(1, XML, "Contact")
            ffrmSnailMail.sStreetLine1 = XMLTag(1, XML, "Street1")
            ffrmSnailMail.sStreetLine2 = XMLTag(1, XML, "Street2")
            ffrmSnailMail.sSuburb = XMLTag(1, XML, "Suburb")
            ffrmSnailMail.sState = XMLTag(1, XML, "State")
            ffrmSnailMail.sPostcode = XMLTag(1, XML, "Postcode")
            ffrmSnailMail.sCountry = XMLTag(1, XML, "Country")
    
            ffrmSnailMail.Show 1

            XML = "<Address>"
            XML = XML + vbTab + "<Contact>" & ffrmSnailMail.sContactName & "</Contact>" & vbCrLf
            XML = XML + vbTab + "<Street1>" & ffrmSnailMail.sStreetLine1 & "</Street1>" & vbCrLf
            XML = XML + vbTab + "<Street2>" & ffrmSnailMail.sStreetLine2 & "</Street2>" & vbCrLf
            XML = XML + vbTab + "<Suburb>" & ffrmSnailMail.sSuburb & "</Suburb>" & vbCrLf
            XML = XML + vbTab + "<State>" & ffrmSnailMail.sState & "</State>" & vbCrLf
            XML = XML + vbTab + "<Postcode>" & ffrmSnailMail.sPostcode & "</Postcode>" & vbCrLf
            XML = XML + vbTab + "<Country>" & ffrmSnailMail.sCountry & "</Country>" & vbCrLf
            XML = XML + "</Address>"
            
            lsDocs.SelectedItem.Tag = XML
            
        Case "PHO"
        
            Dim ffrmPhoneNo As frmPhoneNumber
            Set ffrmPhoneNo = New frmPhoneNumber
            
            ffrmPhoneNo.sContactName = XMLTag(1, XML, "Contact")
            ffrmPhoneNo.sPhonenumber = XMLTag(1, XML, "PhoneNumber")
            ffrmPhoneNo.sExtension = XMLTag(1, XML, "Extension")
            ffrmPhoneNo.sNote = XMLTag(1, XML, "Note")
            ffrmPhoneNo.Show 1
            
            XML = "<Phone>"
            XML = XML + vbTab + "<Contact>" & ffrmPhoneNo.sContactName & "</Contact>" & vbCrLf
            XML = XML + vbTab + "<PhoneNumber>" & ffrmPhoneNo.sPhonenumber & "</PhoneNumber>" & vbCrLf
            XML = XML + vbTab + "<Extension>" & ffrmPhoneNo.sExtension & "</Extension>" & vbCrLf
            XML = XML + vbTab + "<Note>" & ffrmPhoneNo.sNote & "</Note>" & vbCrLf
            XML = XML + "</Phone>"
            
            lsDocs.SelectedItem.Tag = XML
            
        Case "XML"
        
            Dim fXML As New XMLEdit
            
            fXML.XML = XML
            
            fXML.Show 1
            
            lsDocs.SelectedItem.Tag = fXML.XML
            
        End Select
        
        oIP.colDNSDocs(Val(Mid(lsDocs.SelectedItem.Key, 4))).DocText = lsDocs.SelectedItem.Tag
    End If
End Sub

Private Sub txtField_GotFocus(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtField_GotFocus"
    Const ContainerName = "frmDNS"
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


    txtField(Index).SelStart = 0
    txtField(Index).SelLength = Len(txtField(Index).Text)
        
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

Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtField_KeyPress"
    Const ContainerName = "frmDNS"
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


    Select Case KeyAscii
    Case 13
        KeyAscii = 0
        SendKeys "{TAB}"
    End Select
    
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

Private Sub txtField_LostFocus(Index As Integer)

    Select Case Index
    Case 5
        Dim iBox As String
        If oCNT.col_Domains(dKey).TechPass <> "" And txtField(5).Text <> "" Then
            iBox = InputBox("Please enter in the old password to change the technicians password (Case Sensitive).")
            If oCNT.col_Domains(dKey).TechPass = iBox Then
            
                txtField(5).Tag = txtField(5).Text
            Else
                If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.StopAll
                If Not frmAgent.oChar Is Nothing Then frmAgent.oChar.Speak "Passwords do not match, please attempt to change the password again"
                txtField(5).Text = ""
                txtField(5).SetFocus
            End If
        End If
    End Select

End Sub
