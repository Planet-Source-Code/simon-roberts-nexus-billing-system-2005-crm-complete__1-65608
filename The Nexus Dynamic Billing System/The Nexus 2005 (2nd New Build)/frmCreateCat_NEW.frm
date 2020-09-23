VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCreateCat 
   BackColor       =   &H00A9C9AE&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Create Category"
   ClientHeight    =   9765
   ClientLeft      =   4620
   ClientTop       =   3465
   ClientWidth     =   6585
   Icon            =   "frmCreateCat_NEW.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9765
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00A9C9AE&
      Caption         =   "Do not create category"
      Height          =   405
      Index           =   1
      Left            =   3900
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9240
      Width           =   2565
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00A9C9AE&
      Caption         =   "Create Category"
      Height          =   405
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9240
      Width           =   2565
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00A9C9AE&
      Caption         =   "Create Category"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9015
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.Frame Frame5 
         BackColor       =   &H00A9C9AE&
         Caption         =   "Sub Node of Category"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   150
         TabIndex        =   11
         Top             =   5970
         Width           =   5985
         Begin VB.ListBox cmbNodes 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   2460
            Left            =   150
            TabIndex        =   12
            Top             =   300
            Width           =   5715
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00A9C9AE&
         Caption         =   "Minimum Security Level to Access Category"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   150
         TabIndex        =   7
         Top             =   5130
         Width           =   5985
         Begin MSComctlLib.Slider sldSec 
            Height          =   315
            Left            =   150
            TabIndex        =   8
            Top             =   300
            Width           =   5745
            _ExtentX        =   10134
            _ExtentY        =   556
            _Version        =   393216
            Max             =   100
            TickFrequency   =   3
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00A9C9AE&
         Caption         =   "Icon"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3675
         Left            =   150
         TabIndex        =   3
         Top             =   1410
         Width           =   5985
         Begin VB.PictureBox picIcon 
            BackColor       =   &H00A9C9AE&
            BorderStyle     =   0  'None
            Height          =   3345
            Left            =   120
            ScaleHeight     =   3345
            ScaleWidth      =   5505
            TabIndex        =   6
            Top             =   240
            Width           =   5505
            Begin VB.Image imgIcon 
               BorderStyle     =   1  'Fixed Single
               Height          =   540
               Index           =   0
               Left            =   60
               Picture         =   "frmCreateCat_NEW.frx":030A
               Top             =   60
               Width           =   540
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00A9C9AE&
            BorderStyle     =   0  'None
            Height          =   3375
            Left            =   90
            ScaleHeight     =   3375
            ScaleWidth      =   5835
            TabIndex        =   4
            Top             =   210
            Width           =   5835
            Begin VB.VScrollBar vslIcon 
               Height          =   3345
               Left            =   5580
               Max             =   100
               TabIndex        =   5
               Top             =   0
               Width           =   255
            End
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00A9C9AE&
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   150
         TabIndex        =   1
         Top             =   330
         Width           =   5955
         Begin VB.TextBox txtField 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   510
            Left            =   120
            TabIndex        =   2
            Top             =   270
            Width           =   5745
         End
      End
   End
End
Attribute VB_Name = "frmCreateCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public tv As TreeView
Public SubNode As Double
Public Description As String
Public SecLevel As Byte
Public iIcon As Integer
Public RecID As Double


Private Sub Command1_Click(Index As Integer)

    Select Case Index
    Case 0
    
        If Trim(txtField.Text) <> "" Then
        
            Dim ix As Integer
            
            For ix = imgIcon.LBound To imgIcon.UBound
                If imgIcon(ix).BorderStyle = 1 Then
                    Me.iIcon = ix
                    Exit For
                End If
            Next ix
            
            If cmbNodes.ListIndex <> -1 Then
                SubNode = cmbNodes.ItemData(cmbNodes.ListIndex)
            Else
                SubNode = 0
            End If
            
            Description = txtField.Text
            SecLevel = sldSec.Value
            
            If Me.RecID = 0 Then
                On Error Resume Next
                Do
                       Err.Clear
                       
                    Me.RecID = MySQL.GetTMPRecID("exp_categories", directConn)
                    MySQL.Execute directConn, "insert into exp_categories (RecID,VirtualID,SysopID) Values('" & Me.RecID & "','" & Login.lVirtualID & "','" & Login.lSysopID & "')"
                    
                Loop Until Err.Number = 0
                
                GUI.mapCategory.Add "r" & Me.RecID, Me.RecID, SubNode, Val(Login.lVirtualID), Val(Login.lSysopID), iIcon, Description, "exp001", sldSec.Value, "r" & Me.RecID
                
                Select Case SubNode
                Case 0
                    tv.NodeS.Add , , "r" & Me.RecID, Description, iIcon, iIcon
                    
                Case Else
                
                    tv.NodeS.Add "r" & SubNode, tvwChild, "r" & Me.RecID, Description, iIcon, iIcon
                    
                End Select
            End If
            
            MySQL.Execute directConn, "update exp_categories set SubRecID = '" & SubNode & "', Description = '" & MySQL.ESC(Description) & "', SecLevel = '" & SecLevel & "' where RecID = " & Me.RecID
                
            GUI.mapCategory(GUI.mapCategory.FindKey("r" & Me.RecID)).Description = Description
            GUI.mapCategory(GUI.mapCategory.FindKey("r" & Me.RecID)).SecLevel = sldSec.Value
            GUI.mapCategory(GUI.mapCategory.FindKey("r" & Me.RecID)).SubRecID = SubNode
            GUI.mapCategory(GUI.mapCategory.FindKey("r" & Me.RecID)).Icon = iIcon
            
            Unload Me
            
        End If
        
        
    Case 1
    
        Unload Me
        
    End Select
End Sub

Private Sub Form_Load()

    Dim i As Integer
    Static Level As Integer
    
    For i = 1 To fIcon.il32x32.ListImages.Count
        If i = 1 Then
            imgIcon(i - 1).Picture = fIcon.il32x32.ListImages(i).ExtractIcon
            imgIcon(i - 1).Move 120, 0
        Else
            Load imgIcon(i - 1)
            Set imgIcon(i - 1).Container = picIcon
            imgIcon(i - 1).Picture = fIcon.il32x32.ListImages(i).ExtractIcon
            imgIcon(i - 1).Move imgIcon(0).Left + imgIcon(i - 2).Left + imgIcon(0).Width, Level * (imgIcon(0).Width + 120)
            If imgIcon(i - 1).Left > picIcon.ScaleWidth - imgIcon(i - 1).Width Then
                Level = Level + 1
                imgIcon(i - 1).Move imgIcon(0).Left, Level * (imgIcon(0).Width + 120)
            End If
            imgIcon(i - 1).Visible = True
            imgIcon(i - 1).BorderStyle = 0
        End If
    
    Next i
    
    picIcon.Height = imgIcon(fIcon.il32x32.ListImages.Count - 1).Top + imgIcon(fIcon.il32x32.ListImages.Count - 1).Height
    Set picIcon.Container = Picture1
    picIcon.Move 0, 0
    Picture1.Move 90, 240, Frame3.Width - 180, Frame3.Height - 330
    
    cmbNodes.AddItem "[Primary Node]"
    cmbNodes.ItemData(cmbNodes.ListCount - 1) = 0
    
    
    If tv.NodeS.Count > 0 Then
        Dim lx As Long
       
        For lx = 1 To tv.NodeS.Count
            cmbNodes.AddItem tv.NodeS(lx).Text
            cmbNodes.ItemData(cmbNodes.ListCount - 1) = Val(Mid(tv.NodeS(lx).Key, 2))
            If tv.SelectedItem Is Nothing Then
            
            Else
            
                If tv.NodeS(lx).Index = tv.SelectedItem.Index Then
                    cmbNodes.ListIndex = cmbNodes.ListCount - 1
                End If
                
            End If
        Next lx
        
    End If
    
End Sub

Private Sub imgIcon_Click(Index As Integer)

    Dim ix As Integer
    
    For ix = imgIcon.LBound To imgIcon.UBound
        imgIcon(ix).BorderStyle = 0
    Next
    
    imgIcon(Index).BorderStyle = 1
    
End Sub

Private Sub Slider1_Change()

    If Slider1.Value > Login.lLevel Then Slider1.Value = Login.lLevel
    
End Sub

Private Sub vslIcon_Change()

    picIcon.Top = -((vslIcon.Value / 100) * (picIcon.Height - Picture1.Height))
    
    
End Sub
