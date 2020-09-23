VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSnailMail 
   BackColor       =   &H00BA3F3F&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Address"
   ClientHeight    =   5715
   ClientLeft      =   3690
   ClientTop       =   3855
   ClientWidth     =   9675
   ControlBox      =   0   'False
   Icon            =   "frmSnailMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList il32x32 
      Left            =   9060
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSnailMail.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSnailMail.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSnailMail.frx":0CE6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tb 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   1429
      ButtonWidth     =   1667
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "il32x32"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Address"
            ImageIndex      =   3
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Designation"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00AF94D3&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   7710
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5100
      Width           =   1815
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00AF94D3&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5100
      Width           =   1905
   End
   Begin VB.Frame frame1 
      BackColor       =   &H00BA3F3F&
      Height          =   4035
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   930
      Width           =   9405
      Begin VB.ComboBox cmbCountry 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3030
         Width           =   8985
      End
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   3930
         MaxLength       =   20
         TabIndex        =   4
         Top             =   2310
         Width           =   2235
      End
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   6240
         MaxLength       =   50
         TabIndex        =   5
         Top             =   2310
         Width           =   2925
      End
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   180
         MaxLength       =   50
         TabIndex        =   3
         Top             =   2310
         Width           =   3675
      End
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   180
         MaxLength       =   100
         TabIndex        =   2
         Top             =   1590
         Width           =   9015
      End
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   180
         MaxLength       =   100
         TabIndex        =   1
         Top             =   930
         Width           =   9045
      End
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   180
         MaxLength       =   100
         TabIndex        =   0
         Top             =   270
         Width           =   9045
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Country"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0084E8E8&
         Height          =   225
         Index           =   6
         Left            =   210
         TabIndex        =   16
         Top             =   3510
         Width           =   8955
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "State"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0084E8E8&
         Height          =   195
         Index           =   5
         Left            =   3900
         TabIndex        =   15
         Top             =   2700
         Width           =   2235
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Postcode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0084E8E8&
         Height          =   225
         Index           =   4
         Left            =   6240
         TabIndex        =   14
         Top             =   2730
         Width           =   2925
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Suburb"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0084E8E8&
         Height          =   225
         Index           =   3
         Left            =   240
         TabIndex        =   13
         Top             =   2700
         Width           =   3585
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Street Address Line 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0084E8E8&
         Height          =   225
         Index           =   2
         Left            =   180
         TabIndex        =   12
         Top             =   1980
         Width           =   9015
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Street Address Line 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0084E8E8&
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   11
         Top             =   1290
         Width           =   9045
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0084E8E8&
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   10
         Top             =   630
         Width           =   9075
      End
   End
   Begin VB.Frame frame1 
      BackColor       =   &H00BA3F3F&
      Caption         =   "Choose your assigned designation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4065
      Index           =   1
      Left            =   90
      TabIndex        =   18
      Top             =   930
      Visible         =   0   'False
      Width           =   9435
      Begin MSComctlLib.ListView lvFlag 
         Height          =   3615
         Left            =   180
         TabIndex        =   19
         Top             =   330
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   6376
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   16777215
         BackColor       =   12205887
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frmSnailMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sContactName As String
Public sStreetLine1 As String
Public sStreetLine2 As String
Public sSuburb As String
Public sState As String
Public sCountry As String
Public sPostcode As String
Public FlagID As Long
Public iCloseState As frm_CloseStates


Private Sub cmdSave_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdSave_Click"
    Const ContainerName = "frmSnailMail"
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

    sContactName = txtField(0).Text
    sStreetLine1 = txtField(1).Text
    sStreetLine2 = txtField(2).Text
    sSuburb = txtField(3).Text
    sState = txtField(4).Text
    sPostcode = txtField(6).Text
    sCountry = cmbCountry.Text
    
    
    iCloseState = frmCloseSave
    Unload Me
    
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

Private Sub cmdCancel_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdCancel_Click"
    Const ContainerName = "frmSnailMail"
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


    iCloseState = frmCloseCancel
    Unload Me
        
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
    Const ContainerName = "frmSnailMail"
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

    cmbCountry.AddItem "australia"
    cmbCountry.AddItem "albania"
    cmbCountry.AddItem "algeria"
    cmbCountry.AddItem "andorra"
    cmbCountry.AddItem "angola"
    cmbCountry.AddItem "anguilla"
    cmbCountry.AddItem "antigua and barbuda"
    cmbCountry.AddItem "argentina"
    cmbCountry.AddItem "armenia"
    cmbCountry.AddItem "ashmore and cartier islands"
    cmbCountry.AddItem "austria"
    cmbCountry.AddItem "azerbaijan"
    cmbCountry.AddItem "the bahamas"
    cmbCountry.AddItem "bahrain"
    cmbCountry.AddItem "bangladesh"
    cmbCountry.AddItem "barbados"
    cmbCountry.AddItem "belarus"
    cmbCountry.AddItem "belgium"
    cmbCountry.AddItem "belize"
    cmbCountry.AddItem "benin"
    cmbCountry.AddItem "bhutan"
    cmbCountry.AddItem "bolivia"
    cmbCountry.AddItem "bosnia and herzegovina"
    cmbCountry.AddItem "botswana"
    cmbCountry.AddItem "bouvet island"
    cmbCountry.AddItem "brazil"
    cmbCountry.AddItem "british indian ocean territory"
    cmbCountry.AddItem "british virgin islands"
    cmbCountry.AddItem "brunei"
    cmbCountry.AddItem "bulgaria"
    cmbCountry.AddItem "burkinafaso"
    cmbCountry.AddItem "burma"
    cmbCountry.AddItem "burundi"
    cmbCountry.AddItem "cambodia"
    cmbCountry.AddItem "cameroon"
    cmbCountry.AddItem "canada"
    cmbCountry.AddItem "capeverde"
    cmbCountry.AddItem "cayman islands"
    cmbCountry.AddItem "central african republic"
    cmbCountry.AddItem "chad"
    cmbCountry.AddItem "chile"
    cmbCountry.AddItem "china"
    cmbCountry.AddItem "clipperton island"
    cmbCountry.AddItem "cocos keeling islands"
    cmbCountry.AddItem "colombia"
    cmbCountry.AddItem "congo"
    cmbCountry.AddItem "cook islands"
    cmbCountry.AddItem "coralsea islands"
    cmbCountry.AddItem "costarica"
    cmbCountry.AddItem "croatia"
    cmbCountry.AddItem "ctedivoire"
    cmbCountry.AddItem "cuba"
    cmbCountry.AddItem "cyprus"
    cmbCountry.AddItem "czechrepublic"
    cmbCountry.AddItem "denmark"
    cmbCountry.AddItem "dialects"
    cmbCountry.AddItem "djibouti"
    cmbCountry.AddItem "dominica"
    cmbCountry.AddItem "dominican republic"
    cmbCountry.AddItem "ecuador"
    cmbCountry.AddItem "egypt"
    cmbCountry.AddItem "elsalvador"
    cmbCountry.AddItem "equatorial guinea"
    cmbCountry.AddItem "eritrea"
    cmbCountry.AddItem "estonia"
    cmbCountry.AddItem "ethiopia"
    cmbCountry.AddItem "europa island"
    cmbCountry.AddItem "falkl and islands islasmalvinas"
    cmbCountry.AddItem "fiji"
    cmbCountry.AddItem "finland"
    cmbCountry.AddItem "france"
    cmbCountry.AddItem "gabon"
    cmbCountry.AddItem "the gambia"
    cmbCountry.AddItem "republic of georgia"
    cmbCountry.AddItem "germany"
    cmbCountry.AddItem "ghana"
    cmbCountry.AddItem "gibraltar"
    cmbCountry.AddItem "glorioso islands"
    cmbCountry.AddItem "greece"
    cmbCountry.AddItem "grenada"
    cmbCountry.AddItem "guate mala"
    cmbCountry.AddItem "guernsey"
    cmbCountry.AddItem "guinea"
    cmbCountry.AddItem "guinea bissau"
    cmbCountry.AddItem "guyana"
    cmbCountry.AddItem "haiti"
    cmbCountry.AddItem "heard islandand mcdonald islands"
    cmbCountry.AddItem "honduras"
    cmbCountry.AddItem "hong kong"
    cmbCountry.AddItem "hungary"
    cmbCountry.AddItem "iceland"
    cmbCountry.AddItem "india"
    cmbCountry.AddItem "indonesia"
    cmbCountry.AddItem "iran"
    cmbCountry.AddItem "iraq"
    cmbCountry.AddItem "ireland"
    cmbCountry.AddItem "isle of man"
    cmbCountry.AddItem "israel"
    cmbCountry.AddItem "italy"
    cmbCountry.AddItem "jamaica"
    cmbCountry.AddItem "japan"
    cmbCountry.AddItem "jersey"
    cmbCountry.AddItem "jordan"
    cmbCountry.AddItem "juande nova island"
    cmbCountry.AddItem "kazakhstan"
    cmbCountry.AddItem "kenya"
    cmbCountry.AddItem "kiribati"
    cmbCountry.AddItem "kuwait"
    cmbCountry.AddItem "kyrgyzstan"
    cmbCountry.AddItem "laos"
    cmbCountry.AddItem "latvia"
    cmbCountry.AddItem "lebanon"
    cmbCountry.AddItem "lesotho"
    cmbCountry.AddItem "liberia"
    cmbCountry.AddItem "libya"
    cmbCountry.AddItem "liechtenstein"
    cmbCountry.AddItem "lithuania"
    cmbCountry.AddItem "luxembourg"
    cmbCountry.AddItem "macau"
    cmbCountry.AddItem "the former yugoslav republic of macedonia"
    cmbCountry.AddItem "madagascar"
    cmbCountry.AddItem "malawi"
    cmbCountry.AddItem "malaysia"
    cmbCountry.AddItem "maldives"
    cmbCountry.AddItem "mali"
    cmbCountry.AddItem "malta"
    cmbCountry.AddItem "mauritania"
    cmbCountry.AddItem "mauritius"
    cmbCountry.AddItem "mayotte"
    cmbCountry.AddItem "mexico"
    cmbCountry.AddItem "moldova"
    cmbCountry.AddItem "monaco"
    cmbCountry.AddItem "mongolia"
    cmbCountry.AddItem "montenegro"
    cmbCountry.AddItem "montserrat"
    cmbCountry.AddItem "morocco"
    cmbCountry.AddItem "mozambique"
    cmbCountry.AddItem "namibia"
    cmbCountry.AddItem "naoero"
    cmbCountry.AddItem "nepal"
    cmbCountry.AddItem "netherlands"
    cmbCountry.AddItem "new zealand"
    cmbCountry.AddItem "nicaragua"
    cmbCountry.AddItem "niger"
    cmbCountry.AddItem "nigeria"
    cmbCountry.AddItem "niue"
    cmbCountry.AddItem "norfolk island"
    cmbCountry.AddItem "north korea"
    cmbCountry.AddItem "norway"
    cmbCountry.AddItem "oman"
    cmbCountry.AddItem "pakistan"
    cmbCountry.AddItem "panama"
    cmbCountry.AddItem "papuanew guinea"
    cmbCountry.AddItem "paraguay"
    cmbCountry.AddItem "peru"
    cmbCountry.AddItem "philippines"
    cmbCountry.AddItem "pitcairn islands"
    cmbCountry.AddItem "poland"
    cmbCountry.AddItem "portugal"
    cmbCountry.AddItem "qatar"
    cmbCountry.AddItem "romania"
    cmbCountry.AddItem "russia"
    cmbCountry.AddItem "rwanda"
    cmbCountry.AddItem "saint kitts and nevis"
    cmbCountry.AddItem "saint lucia"
    cmbCountry.AddItem "saint vincent and the grenadines"
    cmbCountry.AddItem "samoa"
    cmbCountry.AddItem "sanmarino"
    cmbCountry.AddItem "saotome and principe"
    cmbCountry.AddItem "saudiarabia"
    cmbCountry.AddItem "senegal"
    cmbCountry.AddItem "serbia"
    cmbCountry.AddItem "seychelles"
    cmbCountry.AddItem "sierraleone"
    cmbCountry.AddItem "singapore"
    cmbCountry.AddItem "slovakia"
    cmbCountry.AddItem "slovenia"
    cmbCountry.AddItem "solomon islands"
    cmbCountry.AddItem "somalia"
    cmbCountry.AddItem "south africa"
    cmbCountry.AddItem "south georgia"
    cmbCountry.AddItem "the south sandwich islands"
    cmbCountry.AddItem "south korea"
    cmbCountry.AddItem "spain"
    cmbCountry.AddItem "srilanka"
    cmbCountry.AddItem "sudan"
    cmbCountry.AddItem "suriname"
    cmbCountry.AddItem "swaziland"
    cmbCountry.AddItem "sweden"
    cmbCountry.AddItem "switzerland"
    cmbCountry.AddItem "syria"
    cmbCountry.AddItem "taiwan"
    cmbCountry.AddItem "tajikistan"
    cmbCountry.AddItem "tanzania"
    cmbCountry.AddItem "thailand"
    cmbCountry.AddItem "togo"
    cmbCountry.AddItem "trinidad and tobago"
    cmbCountry.AddItem "trist and acunha"
    cmbCountry.AddItem "tromelin island"
    cmbCountry.AddItem "tunisia"
    cmbCountry.AddItem "turkey"
    cmbCountry.AddItem "turkmenistan"
    cmbCountry.AddItem "turks and caicos islands"
    cmbCountry.AddItem "tuvalu"
    cmbCountry.AddItem "uganda"
    cmbCountry.AddItem "ukraine"
    cmbCountry.AddItem "united arabemirates"
    cmbCountry.AddItem "united kingdom"
    cmbCountry.AddItem "uruguay"
    cmbCountry.AddItem "uzbekistan"
    cmbCountry.AddItem "vanuatu"
    cmbCountry.AddItem "vatican city"
    cmbCountry.AddItem "venezuela"
    cmbCountry.AddItem "vietnam"
    cmbCountry.AddItem "western sahara"
    cmbCountry.AddItem "yemen"
    cmbCountry.AddItem "zambia"
    cmbCountry.AddItem "zimbabwe"
    
    txtField(0).Text = sContactName
    txtField(1).Text = sStreetLine1
    txtField(2).Text = sStreetLine2
    txtField(3).Text = sSuburb
    txtField(4).Text = sState
    txtField(6).Text = sPostcode
    
    Dim CX As Integer
    For CX = 0 To cmbCountry.ListCount - 1
        If cmbCountry.List(CX) = sCountry Then
            cmbCountry.ListIndex = CX
            Exit For
        End If
    Next


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
    Const ContainerName = "frmSnailMail"
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


    sContactName = txtField(0).Text
    sStreetLine1 = txtField(1).Text
    sStreetLine2 = txtField(2).Text
    sSuburb = txtField(3).Text
    sState = txtField(4).Text
    sCountry = cmbCountry.Text
    sPostcode = txtField(6).Text

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

Private Sub tb_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
    Case 1
        Frame1(0).Visible = True
        Frame1(1).Visible = False
    Case 2
        Frame1(0).Visible = False
        Frame1(1).Visible = True
    End Select
End Sub

Private Sub txtField_GotFocus(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtField_GotFocus"
    Const ContainerName = "frmSnailMail"
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


    txtField(Index).SelStart = 0
    txtField(Index).SelLength = Len(txtField(Index).Text)
    
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

Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "txtField_KeyPress"
    Const ContainerName = "frmSnailMail"
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


    Select Case KeyAscii
    Case 13
        KeyAscii = 0
        SendKeys "{TAB}"
    End Select
            
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
