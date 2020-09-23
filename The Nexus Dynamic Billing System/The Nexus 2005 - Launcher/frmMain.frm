VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "project alpha - global launcher"
   ClientHeight    =   9045
   ClientLeft      =   540
   ClientTop       =   2175
   ClientWidth     =   13935
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0442
   ScaleHeight     =   9045
   ScaleWidth      =   13935
   Begin VB.Timer tmZoom 
      Interval        =   25
      Left            =   9780
      Top             =   5670
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Way Points"
      Height          =   2445
      Left            =   5160
      TabIndex        =   3
      Top             =   6270
      Width           =   7125
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Your Location"
         Height          =   615
         Left            =   180
         TabIndex        =   5
         Top             =   1710
         Width           =   6795
         Begin VB.Label lblMouse 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   210
            Left            =   6630
            TabIndex        =   7
            Top             =   270
            Width           =   45
         End
      End
      Begin MSComctlLib.ListView lvWay 
         Height          =   1395
         Left            =   150
         TabIndex        =   4
         Top             =   270
         Width           =   6825
         _ExtentX        =   12039
         _ExtentY        =   2461
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "!Title!"
            Text            =   "Server Title"
            Object.Width           =   3087
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "!Longitude!"
            Text            =   "Longitude"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "!Latitude!"
            Text            =   "Latitude"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "!Description!"
            Text            =   "Description"
            Object.Width           =   6174
         EndProperty
      End
   End
   Begin VB.PictureBox footer 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   13875
      TabIndex        =   1
      Top             =   8760
      Width           =   13935
      Begin VB.CheckBox opt 
         Caption         =   "Big Font Mode"
         Height          =   210
         Index           =   0
         Left            =   12360
         TabIndex        =   6
         Tag             =   "/bigfont"
         Top             =   30
         Width           =   1485
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H009191F4&
         Height          =   210
         Left            =   60
         TabIndex        =   2
         Top             =   0
         Width           =   45
      End
   End
   Begin VB.CommandButton waypoint 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   0
      Left            =   13110
      MouseIcon       =   "frmMain.frx":18E788
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6060
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Timer tmCircle 
      Interval        =   25
      Left            =   10230
      Top             =   5670
   End
   Begin VB.Image Image1 
      Height          =   2535
      Left            =   270
      MouseIcon       =   "frmMain.frx":18E8DA
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   270
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Shape circleA 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      Height          =   495
      Index           =   0
      Left            =   12900
      Shape           =   3  'Circle
      Top             =   5850
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mX As Single
Dim mY As Single
Dim mBut As Integer

Const AW_HOR_POSITIVE = &H1 'Animates the window from left to right. This flag can be used with roll or slide animation.
Const AW_HOR_NEGATIVE = &H2 'Animates the window from right to left. This flag can be used with roll or slide animation.
Const AW_VER_POSITIVE = &H4 'Animates the window from top to bottom. This flag can be used with roll or slide animation.
Const AW_VER_NEGATIVE = &H8 'Animates the window from bottom to top. This flag can be used with roll or slide animation.
Const AW_CENTER = &H10 'Makes the window appear to collapse inward if AW_HIDE is used or expand outward if the AW_HIDE is not used.
Const AW_HIDE = &H10000 'Hides the window. By default, the window is shown.
Const AW_ACTIVATE = &H20000 'Activates the window.
Const AW_SLIDE = &H40000 'Uses slide animation. By default, roll animation is used.
Const AW_BLEND = &H80000 'Uses a fade effect. This flag can be used only if hwnd is a top-level window.
Private Declare Function AnimateWindow Lib "user32" (ByVal hwnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Boolean


Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1
Private Declare Function GetVersion Lib "kernel32" () As Long
Public Function GetWinVersion() As String
    Dim Ver As Long, WinVer As Long
    Ver = GetVersion()
    WinVer = Ver And &HFFFF&
    'retrieve the windows version
    GetWinVersion = Format((WinVer Mod 256) + ((WinVer \ 256) / 100), "Fixed")
End Function


Private Sub Form_Load()

    
    
    Image1.Move 0, 0, Me.ScaleWidth, footer.Top
    
    Image1.Stretch = True
    
    Image1.Picture = Me.Picture

     If Command = "" Or InStr(LCase(Command), "/main") > 0 Then
     
         Me.Caption = "Global Launcher - Project Alpha Terminal"
          
     ElseIf InStr(LCase(Command), "/nsa") > 0 Then
         
         Me.Caption = "Global Launcher - Sysop Control Center"
         
     ElseIf InStr(LCase(Command), "/deploy") > 0 Then
         
         Me.Caption = "Global Launcher - Deployment"
         
     ElseIf InStr(LCase(Command), "/template") > 0 Then
         
         Me.Caption = "Global Launcher - Plan and Template Configuration"
         
    End If

    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

    If Val(GetWinVersion) >= 5 Then
        AnimateWindow Me.hwnd, 900, AW_BLEND
    End If
       
    Me.Show
    Me.Refresh
    Frame1.Refresh
    Frame2.Refresh
    lblMouse.Refresh
    
    footer.Refresh
    
    DoEvents
    
    lblStat.Caption = "Connection to Server being established - www.projectalpha.com.au"
    
     oMysql.Connection Crypt("±¿Õ±Á•±‰", False, "None", "None"), Crypt("•±¶ÏÙ.ˆÏ¶ˆ±Ù.ˆÏ¶.»Æ", False, "None", "None"), Crypt("±¿Õ±Á•±‰", False, "None", "None"), Crypt("Õ¢¯•¦¢¶Ô", False, "None", "None"), oConn
    
    If oConn.State = adStateClosed Then
        
        lblStat.Caption = "Error connecting to server, check internet connection and MyODBC Connection"
        Exit Sub
        
    Else
    
        lblStat.Caption = "Connected to Server - Getting Way Points of Servers"
    
        Dim rsload As ADODB.Recordset
        
        Call oMysql.OpenTable(oConn, rsload, , "select Title, Description, Longitude, Latitude, ID from map_paservers")
        
        If rsload.State = adStateOpen Then
        
            Call oMysql.fillLV(oConn, rsload, lvWay, False, , "ID")
            
            Dim lx As Long
            
            Dim xPos As Single
            Dim yPos As Single
    
            lblStat.Caption = "There where " & lvWay.ListItems.Count & " waypoints of servers found"
            
            For lx = 1 To lvWay.ListItems.Count
                Load circleA(lx)
                circleA(lx).Left = -5000
                circleA(lx).Top = -5000
                Load waypoint(lx)
                waypoint(lx).Tag = 600
    
                xPos = Me.ScaleWidth / 2
                yPos = footer.Top / 2
        
                yPos = yPos + (-Val(lvWay.ListItems(lx).SubItems(2)) / 85) * yPos
                xPos = xPos + (Val(lvWay.ListItems(lx).SubItems(1)) / 170) * xPos
            
                waypoint(lx).Move xPos - waypoint(lx).Width / 2, yPos - waypoint(lx).Height / 2
                waypoint(lx).ToolTipText = lvWay.ListItems(lx).Text + " - " & lvWay.ListItems(lx).SubItems(3)
                waypoint(lx).Visible = True
                circleA(lx).Visible = True
                
                
            Next
        
        Else
            lblStat.Caption = "Error getting Way Points"
        End If
    End If
    
    lblStat.Caption = ""
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mX = X
    mY = Y
    mBut = Shift
    tmZoom.Enabled = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim xPos As Single
    Dim yPos As Single
    Dim iDisa As Long
    Dim iDisb As Long
    Dim bFound As Boolean
    
            xPos = Me.ScaleWidth / 2
            yPos = Me.ScaleHeight / 2
    
            
            xPos = IIf(xPos - X > 0, -((xPos - X) / xPos) * 170, ((X - xPos) / xPos) * 170)
            yPos = IIf(yPos - Y > 0, ((yPos - Y) / yPos) * 85, -((Y - yPos) / yPos) * 85)
            
            
            
    lblMouse.Caption = "(" & Format(xPos, "####.##") & ", " & Format(yPos, "####.##") & ")"
    DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)

    AnimateWindow Me.hwnd, 200, AW_VER_POSITIVE Or AW_HOR_NEGATIVE Or AW_HIDE
    oConn.Close
    End
    
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mX = -1
    mY = -1
    mBut = -1
    tmZoom.Enabled = False
    
End Sub

Private Sub tmCircle_Timer()

    On Error Resume Next
    
    Dim lx As Long
    
    Static R As Byte
    Static B As Byte
    Static G As Byte
    Static RDir As Boolean
    Static BDir As Boolean
    Static GDir As Boolean
    Static lMode As Integer
    
    For lx = waypoint.LBound To waypoint.UBound
        If waypoint(lx).Visible = True Then
            Randomize Now / 900 ^ 3
            
                R = R + IIf(RDir = False, Round(Rnd * 10), -Round(Rnd * 10))
                B = B + IIf(BDir = False, Round(Rnd * 10), -Round(Rnd * 10))
                G = G + IIf(GDir = False, Round(Rnd * 10), -Round(Rnd * 10))
                If R >= 230 Then RDir = True
                If B >= 230 Then BDir = True
                If G >= 230 Then GDir = True
                If R <= 20 Then RDir = False
                If B <= 20 Then BDir = False
                If G <= 20 Then GDir = False
            
            waypoint(lx).BackColor = RGB(R, G, B)
            
            If circleA(lx).Height >= Val(waypoint(lx).Tag) And circleA(lx).Width >= Val(waypoint(lx).Tag) Then
                circleA(lx).Move waypoint(lx).Left + (waypoint(lx).Width / 2) - 1, waypoint(lx).Top + (waypoint(lx).Height / 2) - 1, 1, 1
            Else
                circleA(lx).Move circleA(lx).Left - 5, circleA(lx).Top - 5, circleA(lx).Width + 10, circleA(lx).Height + 10
                circleA(lx).BorderColor = RGB(R, G, B)
            End If
        End If
        
    Next lx
    
End Sub

Private Sub waypoint_Click(Index As Integer)

    If GetSetting("projectalpha", "PublicKey", "B-0", -1) = -1 Then
        
        Dim fKey As New frmPublicKey
        
        fKey.KeyName = "PublicKey"
        fKey.PrimaryKey = "projectalpha"
        fKey.Show 1
    
    End If
    
    
    Dim rsload As ADODB.Recordset
    
    Call oMysql.OpenTable(oConn, rsload, , "select ID, AES_DECRYPT(ADOConnSTR,RecUnlock) as ADOConn from map_paservers where ID = '" & Mid(lvWay.ListItems(Index).Key, 2) & "'")
    
    
    Dim path As String, extension As String
    
    path = IIf(Right(App.path, 1) = "\", App.path, App.path + "\")
    
    Dim lx As Byte
    
    For lx = opt.LBound To opt.UBound
        If opt(lx).Value = 1 Then
            extension = extension + " " + opt(lx).Tag
        End If
    Next

    
    If Dir(path + "projectalpha.exe", vbNormal) <> "" Then
            
            If Command = "" Or InStr(LCase(Command), "/main") > 0 Then
            
                SaveSetting "projectalpha", "db", "ConnectionString", Crypt(rsload!ADOConn, True, "PublicKey", "projectalpha")
                ShellExecute Me.hwnd, vbNullString, path + "projectalpha.exe " + extension, vbNullString, "C:\", SW_SHOWNORMAL
            
                                
            ElseIf InStr(LCase(Command), "/sysopcontrol") > 0 Then
            
                SaveSetting "projectalpha", "db", "ConnectionString", Crypt(rsload!ADOConn, True, "PublicKey", "projectalpha")
                ShellExecute Me.hwnd, vbNullString, path + "NSA.exe " + extension, vbNullString, "C:\", SW_SHOWNORMAL
            
            ElseIf InStr(LCase(Command), "/deploy") > 0 Then
                
                SaveSetting "projectalpha", "db", "ConnectionString", Crypt(rsload!ADOConn, True, "PublicKey", "projectalpha")
                ShellExecute Me.hwnd, vbNullString, path + "pa_deployment.exe " + extension, vbNullString, "C:\", SW_SHOWNORMAL
            
            ElseIf InStr(LCase(Command), "/template") > 0 Then
                
                SaveSetting "projectalpha", "db", "ConnectionString", Crypt(rsload!ADOConn, True, "PublicKey", "projectalpha")
                ShellExecute Me.hwnd, vbNullString, path + "templateeditor.exe " + extension, vbNullString, "C:\", SW_SHOWNORMAL
            
           End If
    End If
    
    Unload Me
    
End Sub
