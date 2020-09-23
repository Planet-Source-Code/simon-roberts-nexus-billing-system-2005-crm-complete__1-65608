VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{683364A1-B37D-11D1-ADC5-006008A5848C}#1.0#0"; "dhtmled.ocx"
Begin VB.Form XMLEdit 
   BackColor       =   &H00BA3F3F&
   Caption         =   "XML Editor"
   ClientHeight    =   11835
   ClientLeft      =   2520
   ClientTop       =   2670
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "XMLEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11835
   ScaleWidth      =   9480
   Begin VB.Frame fraPreview 
      BackColor       =   &H00800000&
      Caption         =   "&DHTML Edit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5025
      Left            =   1380
      TabIndex        =   3
      Top             =   6720
      Width           =   7965
      Begin DHTMLEDLibCtl.DHTMLEdit DHTML 
         Height          =   4545
         Left            =   150
         TabIndex        =   4
         Top             =   330
         Width           =   7695
         ActivateApplets =   -1  'True
         ActivateActiveXControls=   -1  'True
         ActivateDTCs    =   -1  'True
         ShowDetails     =   -1  'True
         ShowBorders     =   -1  'True
         Appearance      =   1
         Scrollbars      =   -1  'True
         ScrollbarAppearance=   1
         SourceCodePreservation=   -1  'True
         AbsoluteDropMode=   0   'False
         SnapToGrid      =   -1  'True
         SnapToGridX     =   50
         SnapToGridY     =   50
         BrowseMode      =   0   'False
         UseDivOnCarriageReturn=   0   'False
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8310
      Top             =   420
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
            Picture         =   "XMLEdit.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "XMLEdit.frx":075C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "XMLEdit.frx":0BAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "XMLEdit.frx":1000
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "XMLEdit.frx":1452
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "XMLEdit.frx":18A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "XMLEdit.frx":1CF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "XMLEdit.frx":2148
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "XMLEdit.frx":259A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "XMLEdit.frx":29EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "XMLEdit.frx":2E3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "XMLEdit.frx":3290
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "XMLEdit.frx":36E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "XMLEdit.frx":39FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "XMLEdit.frx":3E4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "XMLEdit.frx":42A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "XMLEdit.frx":46F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "XMLEdit.frx":4B44
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "XMLEdit.frx":4F96
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "XMLEdit.frx":53E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "XMLEdit.frx":583A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer filler 
      Left            =   8940
      Top             =   570
   End
   Begin MSComctlLib.TabStrip ts 
      Height          =   11775
      Left            =   0
      TabIndex        =   1
      Top             =   30
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   20770
      MultiRow        =   -1  'True
      Style           =   1
      Placement       =   2
      Separators      =   -1  'True
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "<Tags Menu>"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox XMLWorkPad 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   1380
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1080
      Width           =   7995
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "XML Scetch Pad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0084E8E8&
      Height          =   480
      Left            =   2310
      TabIndex        =   2
      Top             =   90
      Width           =   3165
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0084E8E8&
      BorderWidth     =   3
      X1              =   2280
      X2              =   9480
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Image Image1 
      Height          =   945
      Left            =   1380
      Picture         =   "XMLEdit.frx":5C8C
      Stretch         =   -1  'True
      Top             =   60
      Width           =   885
   End
   Begin VB.Menu mnuVault 
      Caption         =   "&XML Vault"
      Begin VB.Menu mnuVault_tags 
         Caption         =   "<Tags menu>"
         Begin VB.Menu mnuVault_tags_Curr 
            Caption         =   "Tags Currently Propogated"
         End
      End
      Begin VB.Menu mnuVault_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVault_Save 
         Caption         =   "&Save && Close"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "XMLEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const bKeyPresses = 45
Dim lSelPos As Long
Public XML As String


Private Sub DHTML_DisplayChanged()

    XMLWorkPad.Text = DHTML.DocumentHTML
    
End Sub

Private Sub Form_Load()
Dim hSysMenu As Long     ' handle to the system menu
    Dim Count As Long        ' the number of items initially on the menu
    Dim mii As MENUITEMINFO  ' describes a menu item to add
    Dim retval As Long       ' return value
    
    ' Get a handle to the system menu.
    'hSysMenu = GetSystemMenu(Form1.hWnd, 0)
    ' See how many items are currently in it.
    
    
    Dim hMenu As Long, hSubMenu As Long, lngID As Long

    'Get the handle of the form's menu
    hMenu = GetMenu(Me.hwnd)
    'Get the handle of the form's submenu
    hMenu = GetSubMenu(hMenu, 0)
    hSubMenu = GetSubMenu(hMenu, 0)
    Count = GetMenuItemCount(hSubMenu)
    
    ' Add a separator bar and then Always On Top to the system menu.
    With mii
        ' The size of the structure.
        .cbSize = Len(mii)
        ' What parts of the structure to use.
        .fMask = MIIM_ID Or MIIM_TYPE
        ' This is a separator.
        .ftype = MFT_SEPARATOR
        ' It has an ID of 0.
        .wID = 0
    End With
    ' Add the separator to the end of the system menu.
    '******************************************************
    'retval = InsertMenuItem(hSysMenu, count, 1, mii)
    
    ' Likewise, add the Always On Top command.
    With mii
        .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE
        ' This is a regular text item.
        .ftype = MFT_STRING
        ' The option is enabled.
        .fState = MFS_ENABLED
        ' It has an ID of 1 (this identifies it in the window procedure).
        .wID = 1
        ' The text to place in the menu item.
        .dwTypeData = "<IP>"
        .cch = Len(.dwTypeData)
    End With
    ' Add this to the bottom of the system menu.
    retval = InsertMenuItem(hSubMenu, Count + 1, 1, mii)
    
    XMLWorkPad.Text = XML
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Select Case MsgBox("Do you want to save your work?", vbYesNo + vbInformation, "Save?")
    Case vbYes
        Call mnuVault_Save_Click
        
    End Select
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    XMLWorkPad.Move XMLWorkPad.Left, XMLWorkPad.Top, Me.ScaleWidth - XMLWorkPad.Left - 180, Me.ScaleHeight - XMLWorkPad.Top - 360 - fraPreview.Height
    fraPreview.Move XMLWorkPad.Left, XMLWorkPad.Top + XMLWorkPad.Height + 180, XMLWorkPad.Width
    DHTML.Move DHTML.Left, DHTML.Top, fraPreview.Width - DHTML.Left * 2, fraPreview.Height - DHTML.Top - 90
    
    Line1.X2 = Me.ScaleWidth
    ts.Move 0, 0, ts.Width, Me.ScaleHeight
    
End Sub

' Before unloading, restore the default system menu and remove the
' custom window procedure.
Private Sub Form_Unload(Cancel As Integer)
    Dim retval As Long  ' return value
    
    ' Replace the previous window procedure to prevent crashing.
    retval = SetWindowLong(Me.hwnd, GWL_WNDPROC, pOldProc)
    ' Remove the modifications made to the system menu.
    retval = GetSystemMenu(Me.hwnd, 1)
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub mnuVault_Save_Click()

    XML = XMLWorkPad.Text
    Unload Me
    
End Sub

Private Sub ts_Click()

    Select Case ts.SelectedItem.Index
    Case 1
        PopupMenu mnuVault_tags
    End Select
    
End Sub

Private Sub XMLWorkPad_Change()

    On Error Resume Next
    
    DHTML.DocumentHTML = XMLWorkPad.Text
End Sub

Private Sub XMLWorkPad_KeyPress(KeyAscii As Integer)

    On Error Resume Next
    
    Static KeyAfter As Byte
    Static bOn As Boolean
    Dim sTag As String
    Dim sCapture As String
    
    
    Select Case KeyAscii
    Case 8
    
        If bOn = True Then
            KeyAfter = KeyAfter - 1
        End If
    
    Case 13, 9
        bOn = False
        KeyAfter = 0
        lSelPos = XMLWorkPad.SelStart + 1
        
    Case Asc("<")
        bOn = True
        KeyAfter = 0
        lSelPos = XMLWorkPad.SelStart + 1
        
    Case Asc(">")
        
        If lSelPos + bKeyPresses >= XMLWorkPad.SelStart Then
            sTag = Mid(XMLWorkPad.Text, lSelPos, XMLWorkPad.SelStart - lSelPos + 1)
            sTag = MySQL.ReplaceString(sTag, Chr(9), "")
            
            sCapture = sTag
            If InStr(sTag, "/") > 0 Then Exit Sub
            
            While Asc(Left(sTag, 1)) <= 32
                sTag = Right(sTag, Len(sTag) - 1)
            Wend
            While Asc(Right(sTag, 1)) <= 32
                sTag = Left(sTag, Len(sTag) - 1)
            Wend
            
            While Left(sTag, 1) = "<"
                sTag = Right(sTag, Len(sTag) - 1)
            Wend
            While Right(sTag, 1) = ">"
                sTag = Left(sTag, Len(sTag) - 1)
            Wend
            
            
            sTag = Trim(sTag)
            If sTag = "" Or Len(sTag) < 4 Then Exit Sub
            
            Dim iPos As Long
        
        
            iPos = 0
            If InStr(XMLWorkPad.Text, sCapture) > 0 Then
                'Debug.Print XMLWorkPad.Text
                iPos = InStr(iPos + 1, XMLWorkPad.Text, sCapture)
                If iPos = 1 Then
                    XMLWorkPad.Text = sCapture + ">" & vbCrLf & vbTab & vbTab & vbCrLf & "</" & sTag & ">" & Mid(XMLWorkPad.Text, Len(sCapture) + 1)
                    KeyAscii = 0
                    XMLWorkPad.SelStart = XMLWorkPad.SelStart + 4 + Len(sCapture)
                    bOn = False
                Else
                    XMLWorkPad.Text = Left(XMLWorkPad.Text, iPos - 1) & sCapture + ">" & vbCrLf & vbTab & vbTab & vbCrLf & "</" & sTag & ">" & Mid(XMLWorkPad.Text, iPos + Len(sCapture))
                    XMLWorkPad.SelStart = XMLWorkPad.SelStart + 4 + Len(sCapture)
                    KeyAscii = 0
                    bOn = False
                End If
            End If
            
            If KeyAscii = 0 Then
             
             
            End If
            
        End If
    
    Case Else
    
        If bOn = True Then
            KeyAfter = KeyAfter + 1
        End If
    
        If KeyAfter = bKeyPresses Then
            bOn = False
            KeyAfter = 0
        End If
        
    End Select
    
End Sub
